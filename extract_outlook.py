# -*- coding: utf-8 -*-
"""
Outlook 받은 편지함(또는 다른 폴더)의 최근 이메일을
CSV 파일로 저장하는 스크립트입니다.
Outlook 데스크톱 앱이 실행되어 로그인된 상태여야 합니다.

실행 (uv 사용):
    uv run extract_outlook.py
"""

import win32com.client
import csv
import os
import datetime

# ===== 아래 설정값만 필요하면 바꾸세요 =====
MAX_EMAILS = 50    # 가져올 최근 이메일 개수
FOLDER_INDEX = 6   # 6 = 받은 편지함(Inbox), 5 = 보낸 편지함(Sent Items)
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "OutlookExtract")
# ==========================================


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.GetDefaultFolder(FOLDER_INDEX)
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # 최신 메일이 먼저 오도록 정렬

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = os.path.join(OUTPUT_DIR, "emails_{}.csv".format(timestamp))

    count = 0
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["받은시각", "보낸사람", "제목", "본문"])

        for msg in messages:
            if count >= MAX_EMAILS:
                break
            try:
                body = (msg.Body or "").replace("\n", " ").replace("\r", " ")
                writer.writerow([str(msg.ReceivedTime), msg.SenderName, msg.Subject, body])
                count += 1
            except Exception as e:
                print("메일 하나를 건너뜁니다 (오류: {})".format(e))
                continue

    print("완료! {}개의 이메일을 아래 파일에 저장했습니다:".format(count))
    print(csv_path)


if __name__ == "__main__":
    main()
