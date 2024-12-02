import os
import subprocess
from datetime import datetime

# مسیر پوشه گیت
GIT_REPO_PATH = r"C:\Users\HP\Documents\exel_010"


# پیام کامیت
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


import os
username = os.getlogin()
print(f"Username: {username}")