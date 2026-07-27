# Outlook 이메일 추출 가이드 (Windows, 비개발자용)

> **대상:** 회사 PC(Windows)에서 Outlook을 사용하는 누구나. 코딩 경험이 없어도 순서대로 따라 하면 됩니다.
> **방식:** Windows에 설치된 Outlook 프로그램(`outlook.exe`)에 이미 로그인되어 있는 상태를 그대로 이용해서,
> Python 스크립트가 Outlook에게 "받은 편지함 내용 좀 보여줘"라고 요청하는 방식입니다.
> 별도의 회사 승인(앱 등록)이나 API 키가 필요 없습니다.

---

## 목차

- [Outlook 이메일 추출 가이드 (Windows, 비개발자용)](#outlook-이메일-추출-가이드-windows-비개발자용)
  - [목차](#목차)
  - [1. 이 방식이 하는 일 (요약)](#1-이-방식이-하는-일-요약)
  - [2. 시작하기 전에 꼭 확인할 것 (보안 주의)](#2-시작하기-전에-꼭-확인할-것-보안-주의)
  - [3. 준비물 확인](#3-준비물-확인)
  - [4. Python 설치](#4-python-설치)
  - [5. 필요한 패키지 설치 (pywin32)](#5-필요한-패키지-설치-pywin32)
  - [6. 추출 스크립트 파일 만들기](#6-추출-스크립트-파일-만들기)
    - [6-1. 폴더 만들기](#6-1-폴더-만들기)
    - [6-2. 스크립트 파일 만들기](#6-2-스크립트-파일-만들기)
    - [6-3. 더블클릭 실행용 파일 만들기 (선택, 편의용)](#6-3-더블클릭-실행용-파일-만들기-선택-편의용)
  - [7. 실행하기](#7-실행하기)
  - [8. 결과 확인하기](#8-결과-확인하기)
  - [9. (선택) 매일 자동으로 실행되게 하기](#9-선택-매일-자동으로-실행되게-하기)
  - [10. (선택) 추출한 파일을 서버로 옮기기](#10-선택-추출한-파일을-서버로-옮기기)
  - [11. 자주 발생하는 문제](#11-자주-발생하는-문제)
    - [`python`, `pip` 명령이 안 먹혀요 ("찾을 수 없습니다" 오류)](#python-pip-명령이-안-먹혀요-찾을-수-없습니다-오류)
    - [`ModuleNotFoundError: No module named 'win32com'`](#modulenotfounderror-no-module-named-win32com)
    - [`Outlook.Application` 관련 오류 (예: `pywintypes.com_error`)](#outlookapplication-관련-오류-예-pywintypescom_error)
    - [CSV 파일을 엑셀로 열었더니 한글이 깨져요](#csv-파일을-엑셀로-열었더니-한글이-깨져요)
    - [아무 파일도 안 만들어지고 창이 바로 닫혀요](#아무-파일도-안-만들어지고-창이-바로-닫혀요)
  - [12. 설정값 바꾸는 방법 (요약표)](#12-설정값-바꾸는-방법-요약표)

---

## 1. 이 방식이 하는 일 (요약)

```mermaid
flowchart LR
    A[Windows PC<br/>Outlook.exe 로그인 상태] --> B[Python 스크립트 실행]
    B --> C[Outlook에게 메일 요청<br/>pywin32]
    C --> D[CSV 파일로 저장<br/>제목/보낸사람/시각/본문]
    D --> E[필요 시 서버로 전달]
```

- Outlook 프로그램이 이미 로그인되어 있으면, Python 스크립트가 그 옆에서 "최근 이메일 목록 보여줘"라고
  물어보고, 그 결과를 하나의 표(CSV 파일)로 저장합니다.
- 인터넷 API 키, 회사 IT 승인, 별도 로그인 없이 **지금 쓰고 있는 Outlook 창 그대로** 이용합니다.
- 단, **Outlook 프로그램이 켜져 있어야** 동작합니다 (백그라운드 실행 포함).

---

## 2. 시작하기 전에 꼭 확인할 것 (보안 주의)

> ⚠️ 회사 이메일은 업무 정보/개인정보가 포함될 수 있습니다. 아래를 반드시 먼저 확인하세요.

- [ ] 회사 보안 정책상 이메일 내용을 로컬 파일로 저장하거나 사내망 밖으로 옮기는 것이 허용되는지 확인했습니다.
- [ ] 이 스크립트로 만들어진 CSV 파일은 **본인 PC 안에만** 두고, 함부로 외부 서버/클라우드에 올리지 않습니다.
- [ ] 특정 메일함(예: 받은 편지함)만 대상으로 하며, 필요 이상으로 많은 메일을 추출하지 않습니다.
- [ ] 민감한 메일(인사, 법무, 기밀 문서 등)은 이 방식으로 다루지 않습니다.

확인이 끝났다면 아래 단계로 진행하세요.

---

## 3. 준비물 확인

| 항목 | 확인 방법 |
|------|----------|
| Windows PC | 회사에서 사용하는 PC (Windows 10/11) |
| Outlook 데스크톱 앱 | 시작 메뉴에서 "Outlook" 검색 후 실행되고, 회사 메일이 정상적으로 보이는지 확인 |
| 인터넷/사내망 연결 | Outlook이 메일을 정상 수신하고 있으면 OK |

> Outlook 웹(outlook.office.com)만 쓰고 계셨다면, **Windows용 Outlook 앱을 설치**하고 회사 계정으로 로그인해야
> 이 방식을 쓸 수 있습니다. (사내 포털의 소프트웨어 설치 센터 또는 Microsoft 365 앱 설치 메뉴를 이용하세요.)

---

## 4. Python 설치

이미 Python이 설치되어 있다면 이 단계는 건너뛰어도 됩니다. (확인법: 아래 6번 참고)

1. 웹 브라우저에서 **https://www.python.org/downloads/** 접속
2. 노란색 **"Download Python 3.x.x"** 버튼 클릭 → 설치 파일(.exe) 다운로드
3. 다운로드된 파일을 더블클릭해서 설치 시작
4. **설치 화면 맨 아래의 체크박스 "Add python.exe to PATH"를 반드시 체크** ✅ (이걸 빠뜨리면 나중에 명령이 안 먹힙니다)
5. **"Install Now"** 클릭 → 설치 완료까지 대기
6. 설치 확인:
   - 키보드에서 `Windows 키` 누르고 `cmd` 입력 → **"명령 프롬프트"** 실행
   - 검은 창(명령 프롬프트)에 아래처럼 입력하고 Enter:
     ```
     python --version
     ```
   - `Python 3.x.x` 같은 글자가 나오면 설치 성공입니다.

---

## 5. 필요한 패키지 설치 (pywin32)

`pywin32`는 Python이 Outlook 같은 Windows 프로그램과 대화할 수 있게 해주는 부품입니다.

1. 명령 프롬프트(cmd)를 엽니다 (4번 단계 6번 항목 참고)
2. 아래 명령을 입력하고 Enter:
   ```
   pip install pywin32
   ```
3. 화면에 `Successfully installed pywin32-...` 라는 글자가 보이면 성공입니다.

---

## 6. 추출 스크립트 파일 만들기

### 6-1. 폴더 만들기

1. 바탕화면(Desktop)에 새 폴더를 하나 만듭니다.
2. 폴더 이름을 **`OutlookExtract`** 로 지정합니다.
   - 결과적으로 `C:\Users\내계정\Desktop\OutlookExtract` 폴더가 생깁니다.

### 6-2. 스크립트 파일 만들기

1. 메모장(Notepad)을 엽니다. (`Windows 키` → `메모장` 검색 → 실행)
2. 아래 내용을 **그대로 복사해서** 메모장에 붙여넣습니다.

```python
# -*- coding: utf-8 -*-
"""
Outlook 받은 편지함(또는 다른 폴더)의 최근 이메일을
CSV 파일로 저장하는 스크립트입니다.
Outlook 데스크톱 앱이 실행되어 로그인된 상태여야 합니다.
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
```

3. 메모장 메뉴에서 **"파일" → "다른 이름으로 저장"** 클릭
4. 저장 위치: 방금 만든 `OutlookExtract` 폴더 선택
5. 파일 이름: **`extract_outlook.py`** 입력
   - "파일 형식"은 **"모든 파일"** 로 선택해야 합니다 (안 그러면 `.txt`가 자동으로 붙습니다)
   - 인코딩은 **UTF-8** 선택
6. **"저장"** 클릭

### 6-3. 더블클릭 실행용 파일 만들기 (선택, 편의용)

매번 명령 프롬프트를 여는 게 번거롭다면, 더블클릭만으로 실행되는 파일을 만들 수 있습니다.

1. 메모장을 새로 엽니다.
2. 아래 내용을 그대로 붙여넣습니다.
   ```bat
   @echo off
   cd /d %~dp0
   python extract_outlook.py
   pause
   ```
3. **"파일" → "다른 이름으로 저장"**
4. 저장 위치: 같은 `OutlookExtract` 폴더
5. 파일 이름: **`run.bat`** (파일 형식은 "모든 파일")
6. 저장

이제 `OutlookExtract` 폴더 안에 `extract_outlook.py`와 `run.bat` 두 개 파일이 있어야 합니다.

---

## 7. 실행하기

1. **Outlook 프로그램을 실행하고 로그인된 상태**로 열어둡니다. (창을 최소화해도 괜찮습니다, 꺼지지만 않으면 됩니다)
2. `OutlookExtract` 폴더로 이동합니다.
3. **`run.bat` 파일을 더블클릭**합니다.
4. 검은 창(명령 프롬프트)이 뜨고 잠깐 실행되다가 다음과 같은 메시지가 나오면 성공입니다:
   ```
   완료! 50개의 이메일을 아래 파일에 저장했습니다:
   C:\Users\내계정\Desktop\OutlookExtract\emails_20260727_101530.csv
   ```
5. 아무 키나 누르면 창이 닫힙니다.

> `run.bat`을 만들지 않았다면, 명령 프롬프트(cmd)를 열고 아래처럼 직접 실행해도 됩니다.
> ```
> cd Desktop\OutlookExtract
> python extract_outlook.py
> ```

---

## 8. 결과 확인하기

- `OutlookExtract` 폴더 안에 `emails_날짜시각.csv` 형태의 파일이 생성됩니다.
- 더블클릭하면 엑셀(Excel)로 자동으로 열립니다.
- 열이 4개입니다: **받은시각 / 보낸사람 / 제목 / 본문**
- 실행할 때마다 새 파일이 생성되므로(파일명에 시각이 포함됨), 예전 파일이 덮어써지지 않습니다.

---

## 9. (선택) 매일 자동으로 실행되게 하기

Windows의 "작업 스케줄러"를 이용하면 사람이 직접 더블클릭하지 않아도 매일 자동으로 실행됩니다.

1. `Windows 키` → **"작업 스케줄러"** 검색 후 실행
2. 오른쪽의 **"기본 작업 만들기..."** 클릭
3. 이름 입력 (예: `Outlook 이메일 추출`) → **다음**
4. 트리거: **"매일"** 선택 → **다음** → 원하는 시각 설정 (예: 매일 오전 9시) → **다음**
5. 동작: **"프로그램 시작"** 선택 → **다음**
6. **"프로그램/스크립트"** 칸에 아래 내용 입력:
   ```
   C:\Users\내계정\Desktop\OutlookExtract\run.bat
   ```
   (`내계정` 부분은 실제 Windows 로그인 계정명으로 바꿔주세요)
7. **다음** → **마침**

> 주의: 작업 스케줄러로 실행되는 시각에 Outlook 프로그램이 열려 있어야 정상 동작합니다.
> Outlook이 꺼져 있으면 스크립트가 오류를 내며 실패할 수 있습니다.

---

## 10. (선택) 추출한 파일을 서버로 옮기기

CSV 파일을 Linux 서버(hermes)나 다른 Python 프로그램에서 활용하려면, 파일을 옮기는 절차가 한 번 더 필요합니다.
아래 방법 중 회사 보안 정책에서 허용하는 방법을 사용하세요.

| 방법 | 설명 | 난이도 |
|------|------|--------|
| 회사 공유 드라이브/OneDrive 동기화 폴더 | `OutlookExtract` 폴더 대신 동기화되는 폴더에 저장하도록 6-2번 스크립트의 `OUTPUT_DIR` 값을 그 폴더 경로로 변경 | 쉬움 |
| WinSCP (SFTP) | Windows용 무료 프로그램으로 서버에 파일 업로드 (https://winscp.net) | 보통 |
| USB/사내 파일 전송 시스템 | 회사에서 승인한 방식으로 수동 전달 | 쉬움 |

> 이 문서는 "Windows에서 이메일을 추출하는 단계"까지만 다룹니다. 추출된 CSV 파일을 hermes 서버가
> 자동으로 읽어서 활용하게 하려면, 서버 쪽에서 해당 폴더를 주기적으로 확인하는 별도 작업(예: cron,
> kanban 연동)이 추가로 필요하며, 이는 다음 단계 작업으로 진행합니다.

---

## 11. 자주 발생하는 문제

### `python`, `pip` 명령이 안 먹혀요 ("찾을 수 없습니다" 오류)
- Python 설치 시 **"Add python.exe to PATH"** 체크를 안 했을 가능성이 큽니다.
- Python을 다시 설치하면서 그 체크박스를 반드시 체크하세요.

### `ModuleNotFoundError: No module named 'win32com'`
- `pip install pywin32` 가 제대로 실행되지 않은 것입니다. 5번 단계를 다시 실행하세요.

### `Outlook.Application` 관련 오류 (예: `pywintypes.com_error`)
- Outlook 프로그램이 켜져 있지 않거나 로그인되어 있지 않은 상태입니다.
- Outlook을 먼저 실행하고 받은 편지함이 정상적으로 보이는지 확인한 뒤 다시 시도하세요.

### CSV 파일을 엑셀로 열었더니 한글이 깨져요
- 스크립트에서 이미 `utf-8-sig` 인코딩으로 저장하므로 보통 문제가 없습니다.
- 만약 깨진다면, 엑셀에서 파일을 열지 말고 **"데이터" → "텍스트/CSV에서"** 메뉴로 불러오면서
  인코딩을 "UTF-8"로 지정해서 열어보세요.

### 아무 파일도 안 만들어지고 창이 바로 닫혀요
- `run.bat`을 더블클릭 대신, cmd 창을 직접 열고 `run.bat`을 실행해서 오류 메시지를 확인하세요.
  ```
  cd Desktop\OutlookExtract
  run.bat
  ```

---

## 12. 설정값 바꾸는 방법 (요약표)

`extract_outlook.py` 파일 상단의 아래 3줄만 메모장으로 열어서 수정하면 동작을 바꿀 수 있습니다.

| 설정 | 기본값 | 설명 |
|------|--------|------|
| `MAX_EMAILS` | `50` | 한 번에 가져올 최근 이메일 개수 |
| `FOLDER_INDEX` | `6` | `6`=받은 편지함, `5`=보낸 편지함 |
| `OUTPUT_DIR` | 바탕화면의 `OutlookExtract` 폴더 | 결과 CSV 파일이 저장될 위치 |

수정 후에는 메모장에서 저장(`Ctrl+S`)하고, `run.bat`을 다시 더블클릭하면 새 설정으로 실행됩니다.
