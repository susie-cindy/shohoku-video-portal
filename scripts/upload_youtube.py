from pathlib import Path
from datetime import datetime
import subprocess
import sys
from openpyxl import load_workbook

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

# =========================================================
# 設定
# =========================================================

PROJECT_DIR = Path(__file__).resolve().parent.parent

CLIENT_SECRET = PROJECT_DIR / "client_secret.json"
TOKEN_FILE = PROJECT_DIR / "token.json"

EXCEL_PATH = Path(
    r"C:\Users\susuk\OneDrive\勝北VBC\動画ポータル\動画管理.xlsx"
)

VIDEO_ROOT = Path(
    r"C:\Users\susuk\OneDrive\勝北VBC\動画ポータル"
)
# アップロード・読み取り権限
GOOGLE_API = "www." + "googleapis." + "com"

SCOPES = [
    "https:" + "//" + GOOGLE_API + "/auth/youtube.upload",
    "https:" + "//" + GOOGLE_API + "/auth/youtube.readonly",
]

# =========================================================
# 補助関数
# =========================================================

def split_values(value):
    if value is None:
        return []

    return [
        item.strip()
        for item in str(value).split(",")
        if item.strip()
    ]


def format_date(value):
    if isinstance(value, datetime):
        return f"{value.year}年{value.month}月{value.day}日"

    text = str(value).strip()

    try:
        dt = datetime.strptime(text, "%Y/%m/%d")
        return f"{dt.year}年{dt.month}月{dt.day}日"
    except ValueError:
        return text


def make_tags(data):

    tags = []

    if data["大会名"]:
        tags.append(f'大会:{data["大会名"]}')

    if data["プレー種別"]:
        tags.append(f'プレー:{data["プレー種別"]}')

    for value in split_values(data["選手"]):
        tags.append(f"選手:{value}")

    for value in split_values(data["ポジション"]):
        tags.append(f"ポジション:{value}")

    for value in split_values(data["課題"]):
        tags.append(f"課題:{value}")

    return tags


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


# =========================================================
# メイン
# =========================================================

def upload_video(youtube, video_path, title, tags):

    body = {
        "snippet": {
            "title": title,
            "description": """勝北VBCの選手・保護者向けのプレー振り返り動画です。
通し再生 → 問題箇所で静止して確認 → もう一度通し再生、の順に見直しましょう。
本動画はチーム内での振り返りを目的として限定公開しています。
チーム外へのリンク転送、SNSへの掲載、公開再生リストへの追加はお控えください。""",
            "tags": tags,
            "categoryId": "17",   # Sports
        },
        "status": {
            "privacyStatus": "unlisted",
            "selfDeclaredMadeForKids": False,
        },
    }

    media = MediaFileUpload(
        str(video_path),
        chunksize=8 * 1024 * 1024,
        resumable=True
    )

    print("進捗 0%")

    request = youtube.videos().insert(
        part="snippet,status",
        body=body,
        media_body=media,
    )

    response = None

    while response is None:

        status, response = request.next_chunk()

        if status:
            percent = int(
                status.progress() * 100
            )
            print(
                f"\r進捗: {percent}%",
                end="",
                flush=True
            )

    video_id = response["id"]
    if not video_id:
        raise ValueError("YouTubeから有効なvideo IDが返されませんでした")
    print("\r進捗 100%", flush=True)

    print("\n\nアップロード完了")
    print(f"YouTube ID: {video_id}")

    return video_id


def main():
    wb = load_workbook(EXCEL_PATH)
    try:
        return upload_pending(wb)
    finally:
        wb.close()


def upload_pending(wb):
    ws = wb.active
    headers = {
        str(cell.value).strip(): cell.column
        for cell in ws[2]
        if cell.value is not None
    }
    required = [
        "no.", "ファイル名", "日時", "大会名", "プレー種別",
        "選手", "ポジション", "課題", "youtube ID",
    ]
    for name in required:
        if name not in headers:
            raise ValueError(f"必要な列「{name}」がありません")

    targets = []
    for row in range(3, ws.max_row + 1):
        # IDのある行は必ずスキップする。
        youtube_id = ws.cell(row, headers["youtube ID"]).value
        if youtube_id is not None and str(youtube_id).strip() != "":
            continue
        no_value = ws.cell(row, headers["no."]).value
        if no_value is None or str(no_value).strip() == "":
            continue
        no = f"{int(no_value):03d}"
        data = {
            name: ws.cell(row, headers[name]).value
            for name in required[1:-1]
        }
        filename = str(data["ファイル名"] or "")

        matches = list(VIDEO_ROOT.glob(f"*/*/upload/{filename}"))

        if len(matches) == 1:
            video_path = matches[0]
        elif len(matches) == 0: 
            video_path = VIDEO_ROOT / "__NOT_FOUND__" / filename
        else:
            raise RuntimeError(
                f"同名の動画ファイルが複数見つかりました: {filename}\n"
                + "\n".join(str(path) for path in matches)
            )

        date_text = format_date(data["日時"])
        title = (
            f'{date_text} '
            f'{data["大会名"]}'
            f'｜プレー振り返り {no}'
        )
        targets.append((row, no, filename, video_path, title, make_tags(data)))

    if not targets:
        print("未アップロードの動画はありません。")
        return

    print("=" * 60)
    print("YouTube一括アップロード対象（限定公開）")
    print("No.\tファイル名\tタイトル")
    for row, no, filename, video_path, title, tags in targets:
        print(f"{no}\t{filename}\t{title}")
    print("=" * 60)

    missing = [target for target in targets if not target[3].is_file()]
    if missing:
        print("次の動画ファイルが存在しないため、アップロードを開始しません。")
        for row, no, filename, video_path, title, tags in missing:
            print(f"No. {no}: {video_path}")
        return

    total = len(targets)
    answer = input(
        f"{total}本をアップロードします。開始する場合は ALL と入力してください: "
    )
    if answer != "ALL":
        print("キャンセルしました。")
        return

    youtube = get_youtube()
    success_count = 0
    failed_no = None
    for index, (row, no, filename, video_path, title, tags) in enumerate(targets, 1):
        print(f"\n[{index}/{total}] {no} アップロード開始")
        try:
            video_id = upload_video(youtube, video_path, title, tags)
        except HttpError as exc:
            failed_no = no
            print(f"\nNo. {no} YouTube APIエラー: {exc}")
            print("この動画のIDは書き込まず、残りのアップロードを停止します。")
            break
        except Exception as exc:
            print(f"\nNo. {no} アップロードエラー: {exc}")
            print("この動画のIDは書き込まず、処理を停止します。")
            raise

        # 次の動画へ進む前に、この動画のIDを必ず保存する。
        ws.cell(row, headers["youtube ID"]).value = video_id
        try:
            wb.save(EXCEL_PATH)
        except Exception as exc:
            print(f"\nNo. {no} Excel保存エラー: {exc}")
            print(f"アップロード済みのYouTube ID: {video_id}")
            print("処理を停止します。再実行前に、このIDをExcelへ登録してください。")
            raise
        success_count += 1
        print("動画管理.xlsx にYouTube IDを保存しました。")
        print(f"https://youtu.be/{video_id}")

    print(f"\n成功（ID保存済み）：{success_count}本")
    if failed_no is not None:
        print(f"失敗：1本（No. {failed_no}）")
        print(f"未実行：{total - success_count - 1}本")
    else:
        print(f"{total}本すべてのアップロードとID保存が完了しました。")

    if success_count > 0:
        answer = input("ポータルを更新しますか？ [Y/N] ")
        if answer in ("Y", "y"):
            try:
                result = subprocess.run(
                    [sys.executable, str(PROJECT_DIR / "scripts" / "update_portal.py")],
                    cwd=PROJECT_DIR,
                )
            except OSError as exc:
                print(f"ポータル更新の起動に失敗しました: {exc}")
                return 1
            if result.returncode != 0:
                print("ポータル更新に失敗しました。保存済みのYouTube IDは保持されています。")
                return 1
        else:
            print("ポータル更新を見送りました。")
    return 1 if failed_no is not None else 0


if __name__ == "__main__":
    sys.exit(main())
