import json
import os
import re
import tempfile
from pathlib import Path

from portal_history import build_latest_news, parse_timestamp, read_excel_state

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# --------------------------------------------------
# パス
# --------------------------------------------------

PROJECT_DIR = Path(__file__).resolve().parent.parent

CLIENT_SECRET = PROJECT_DIR / "client_secret.json"
TOKEN_FILE = PROJECT_DIR / "token.json"

DATA_DIR = PROJECT_DIR / "data"
OUTPUT_FILE = DATA_DIR / "videos.json"


# --------------------------------------------------
# OAuth
# チャットのURL自動リンク化対策で文字列を分割
# --------------------------------------------------

GOOGLE_API = "www." + "googleapis." + "com"

SCOPES = [
    "https:" + "//" + GOOGLE_API + "/auth/youtube.upload",
    "https:" + "//" + GOOGLE_API + "/auth/youtube.readonly",
]


# --------------------------------------------------
# ポータルで使用するタグ
# --------------------------------------------------

TAG_CATEGORIES = [
    "大会",
    "プレー",
    "選手",
    "ポジション",
    "課題",
]


def get_youtube():
    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(
            TOKEN_FILE,
            SCOPES
        )

    if not creds or not creds.valid:

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET,
                SCOPES
            )
            creds = flow.run_local_server(port=0)

        TOKEN_FILE.write_text(
            creds.to_json(),
            encoding="utf-8"
        )

    return build(
        "youtube",
        "v3",
        credentials=creds
    )


def get_uploads_playlist_id(youtube):
    response = youtube.channels().list(
        part="contentDetails",
        mine=True
    ).execute()

    items = response.get("items", [])

    if not items:
        raise RuntimeError("YouTubeチャンネルを取得できませんでした。")

    return (
        items[0]["contentDetails"]
        ["relatedPlaylists"]["uploads"]
    )


def get_video_ids(youtube, playlist_id):
    video_ids = []
    page_token = None

    while True:
        response = youtube.playlistItems().list(
            part="snippet",
            playlistId=playlist_id,
            maxResults=50,
            pageToken=page_token
        ).execute()

        for item in response.get("items", []):
            video_id = (
                item.get("snippet", {})
                .get("resourceId", {})
                .get("videoId")
            )

            if video_id:
                video_ids.append(video_id)

        page_token = response.get("nextPageToken")

        if not page_token:
            break

    return video_ids


def parse_tags(tags):
    categories = {
        category: []
        for category in TAG_CATEGORIES
    }

    for tag in tags:
        if ":" not in tag:
            continue

        category, value = tag.split(":", 1)

        category = category.strip()
        value = value.strip()

        if category in categories and value:
            if value not in categories[category]:
                categories[category].append(value)

    return categories


def parse_event_date(title):
    """
    タイトル先頭の
    2026年9月26日 ...
    から 2026-09-26 を生成する。
    """

    match = re.match(
        r"^(\d{4})年(\d{1,2})月(\d{1,2})日",
        title
    )

    if not match:
        return None

    year, month, day = map(int, match.groups())

    return f"{year:04d}-{month:02d}-{day:02d}"


def get_video_data(youtube, video_ids):
    videos = []

    # videos.list は最大50 IDずつ
    for start in range(0, len(video_ids), 50):

        batch = video_ids[start:start + 50]

        response = youtube.videos().list(
            part="snippet,status",
            id=",".join(batch)
        ).execute()

        for item in response.get("items", []):

            video_id = item["id"]
            snippet = item.get("snippet", {})
            status = item.get("status", {})

            title = snippet.get("title", "")
            tags = snippet.get("tags", [])

            categories = parse_tags(tags)

            # 勝北VBCポータル用タグがない動画は除外
            # 大会・プレーの両方が設定された動画のみ対象
            if (
                not categories["大会"]
                or not categories["プレー"]
            ):
                continue

            thumbnails = snippet.get("thumbnails", {})

            thumbnail = ""

            for size in (
                "maxres",
                "standard",
                "high",
                "medium",
                "default",
            ):
                if size in thumbnails:
                    thumbnail = thumbnails[size].get("url", "")
                    break

            videos.append({
                "id": video_id,
                "title": title,
                "eventDate": parse_event_date(title),
                "publishedAt": snippet.get("publishedAt"),
                "thumbnail": thumbnail,
                "privacyStatus": status.get("privacyStatus"),
                "tags": tags,
                "categories": categories,
            })

    return videos


def sort_videos(videos):
    """アップロード日時の降順。日時が不明な動画は最後。"""
    def key(video):
        timestamp = parse_timestamp(video.get("publishedAt"))
        return (
            timestamp is None,
            -timestamp if timestamp is not None else 0,
            video["title"], video["id"],
        )
    return sorted(videos, key=key)


def make_portal_output(videos, saved_ids, records):
    excluded = [video for video in videos if video["id"] not in saved_ids]
    if excluded:
        print(f"ExcelにIDが未保存の動画を除外: {len(excluded)}本")
    videos = sort_videos([video for video in videos if video["id"] in saved_ids])
    return {
        "videoCount": len(videos),
        "videos": videos,
        "news": build_latest_news(records, {video["id"] for video in videos}),
    }


def write_portal_output(output, path=OUTPUT_FILE):
    """完成したUTF-8 JSONだけを置換し、書き込み失敗時は旧JSONを残す。"""
    content = json.dumps(output, ensure_ascii=False, indent=2) + "\n"
    if path.exists() and path.read_text(encoding="utf-8") == content:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", newline="\n", dir=path.parent,
            prefix=".videos-", suffix=".tmp", delete=False,
        ) as handle:
            temporary = Path(handle.name)
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def main():
    # Excel保存済みのIDだけを採用する。読めない場合は既存JSONを更新しない。
    saved_ids, records = read_excel_state()
    print("YouTube動画情報を取得します。")

    youtube = get_youtube()

    playlist_id = get_uploads_playlist_id(youtube)

    print(f"Uploads playlist: {playlist_id}")

    video_ids = get_video_ids(
        youtube,
        playlist_id
    )

    print(
        f"チャンネルのアップロード動画: "
        f"{len(video_ids)}本"
    )

    videos = get_video_data(
        youtube,
        video_ids
    )

    output = make_portal_output(videos, saved_ids, records)
    write_portal_output(output)

    print()
    print(
        f"ポータル対象動画: "
        f"{output['videoCount']}本"
    )
    print(
        f"生成完了: "
        f"{OUTPUT_FILE}"
    )


if __name__ == "__main__":
    main()
