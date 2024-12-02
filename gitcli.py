import os
import subprocess
from datetime import datetime
import argparse

# مسیر پوشه گیت
GIT_REPO_PATH = r"C:\Users\HP\Documents\exel_010"


def git_push():
    """اجرای دستورات گیت برای افزودن، کامیت کردن و پوش کردن تغییرات."""
    commit_message = f"Auto commit on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    try:
        os.chdir(GIT_REPO_PATH)

        # افزودن تغییرات
        subprocess.run(["git", "add", "."], check=True)

        # کامیت کردن
        subprocess.run(["git", "commit", "-m", commit_message], check=True)

        # پوش کردن
        subprocess.run(["git", "push"], check=True)

        print("Changes pushed to GitHub successfully.")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")


def main():
    parser = argparse.ArgumentParser(description="A simple CLI for git automation.")
    parser.add_argument(
        "command",
        type=str,
        nargs="?",
        choices=["push"],
        default="push",  # پیش‌فرض به push تنظیم می‌شود
        help="Command to execute (e.g., 'push').",
    )
    args = parser.parse_args()

    if args.command == "push":
        git_push()


if __name__ == "__main__":
    main()
