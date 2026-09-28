"""YouTube情報を同期し、動画一覧だけをコミット・送信する。"""

from pathlib import Path
import subprocess
import sys


PROJECT_DIR = Path(__file__).resolve().parent.parent
PORTAL_FILE = "data/videos.json"


def git(*args):
    return subprocess.run(
        ["git", *args], cwd=PROJECT_DIR, check=True,
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    ).stdout.strip()


def main():
    stage = "Gitの事前確認"
    try:
        # 他の作業や未送信コミットを巻き込まない。
        root = Path(git("rev-parse", "--show-toplevel")).resolve()
        if root != PROJECT_DIR:
            raise RuntimeError("プロジェクト直下のGitリポジトリではありません。")
        branch = git("symbolic-ref", "--quiet", "--short", "HEAD")
        upstream = git("rev-parse", "--abbrev-ref", "@{upstream}")
        remote = git("config", "--get", f"branch.{branch}.remote")
        remote_ref = git("config", "--get", f"branch.{branch}.merge")
        if remote == "." or not remote_ref.startswith("refs/heads/"):
            raise RuntimeError("通常のリモート追跡ブランチを設定してください。")
        if git("diff", "--cached", "--name-only"):
            raise RuntimeError("ステージ済みの変更があります。先に手動で整理してください。")
        if git("ls-files", "--unmerged"):
            raise RuntimeError("競合が未解決です。先に解決してください。")
        if git("rev-list", "--count", f"{upstream}..HEAD") != "0":
            raise RuntimeError(
                "未送信コミットがあります。内容を確認して手動で送信・整理してから再実行してください。"
            )
        for secret in ("client_secret.json", "token.json"):
            if git("ls-files", "--", secret):
                raise RuntimeError(f"{secret} がGit管理対象です。処理を停止します。")
            git("check-ignore", "--", secret)

        stage = "YouTube情報の同期（sync.py）"
        subprocess.run(
            [sys.executable, str(PROJECT_DIR / "scripts" / "sync.py")],
            cwd=PROJECT_DIR, check=True,
        )
        stage = "動画一覧の差分確認"
        if not git("status", "--porcelain", "--", PORTAL_FILE):
            print("動画一覧に変更はありません。正常終了しました。")
            return 0
        if not (PROJECT_DIR / PORTAL_FILE).is_file():
            raise RuntimeError("同期後の data/videos.json がありません。")

        stage = "動画一覧のステージ登録"
        git("add", "--", PORTAL_FILE)
        if git("diff", "--cached", "--name-only").splitlines() != [PORTAL_FILE]:
            raise RuntimeError("ステージ内容が想定と異なるため停止しました。")
        stage = "動画一覧のコミット"
        git("commit", "--only", "-m", "Update portal video list", "--", PORTAL_FILE)
        stage = "Gitへの送信（push）"
        # 設定による別ブランチ・タグの同時送信を避け、強制送信もしない。
        git("-c", "push.followTags=false", "push", remote, f"HEAD:{remote_ref}")
        print("動画一覧をコミットし、Gitへ送信しました。")
        return 0
    except (OSError, subprocess.CalledProcessError, RuntimeError) as exc:
        print(f"エラー: {stage}に失敗しました。", file=sys.stderr)
        if isinstance(exc, subprocess.CalledProcessError):
            print(f"終了コード: {exc.returncode}", file=sys.stderr)
            if exc.stderr:
                print(exc.stderr.strip(), file=sys.stderr)
        else:
            print(str(exc), file=sys.stderr)
        print("変更・コミットは自動で取り消しません。Gitの状態を確認してください。", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
