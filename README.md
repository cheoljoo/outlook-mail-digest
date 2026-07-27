# outlook-mail-digest

Outlook 메일함에서 이메일 데이터(제목/발신자/본문/첨부 등)를 추출해 정리하는 도구

## 문서 구성

같은 목표(메일 자동 추출)를 여러 방식으로 시도한 기록이 문서별로 남아있습니다. 어떤 기준으로 나뉘어 있는지 정리하면:

| 문서 | 목적/기준 |
|------|-----------|
| [outlook_data_extraction.md](outlook_data_extraction.md) | **최초 시도(폐기됨)**: Windows Outlook 데스크톱 앱 + `pywin32` COM 자동화로 로컬에서 직접 추출하는 방식. 이 회사 PC에서는 새 Outlook이 COM을 지원하지 않고, 클래식 Outlook은 설치는 되지만 사내 보안 에이전트가 실행을 막아 포기함. 시도/트러블슈팅 기록 보존용. |
| [outlook_power_automate_setup.md](outlook_power_automate_setup.md) | 회사 승인(앱 등록) 없이 가능한 대안으로 채택한 **Power Automate 기반 방식의 범용 가이드(한국어)**. 특정 개인 정보(실제 서버 주소, 계정명 등) 없이 누구나 따라 할 수 있도록 일반화한 버전. |
| [outlook_power_automate_setup.en.md](outlook_power_automate_setup.en.md) | 위 문서와 내용/구성은 동일하지만, 실제 OneDrive·Power Automate UI가 영문이라 버튼/메뉴 이름을 영문 그대로 표기한 버전. |
| [my_outlook_power_automate_setup.md](my_outlook_power_automate_setup.md) | **실제로 이 프로젝트에서 쓰고 있는 개인화된 작업 버전.** 실제 서버(`tiger02.lge.com`), 실제 스크립트(`sync_outlook_digest.py`, `run.bat`) 경로가 그대로 들어가 있고, 시도했다가 안 된 방법(rclone 관리자 승인 차단 등)과 실제 겪은 트러블슈팅이 함께 기록된 **살아있는 문서**. 설정을 바꾸거나 재현할 때는 이 문서를 기준으로 봅니다. |

## 자동화 스크립트

- [sync_outlook_digest.py](sync_outlook_digest.py) — OneDrive에 동기화된 `OutlookDigest.xlsx`를 SSH 키 인증으로 원격 서버(tiger02)에 전송. SSH 키 생성/서버 등록, Windows 작업 스케줄러 등록 여부를 스스로 확인해 필요할 때만 처리(idempotent). 모든 단계를 `sync_outlook_digest.log`에 시간과 함께 기록.
- [run.bat](run.bat) — `uv`가 없으면 설치한 뒤 `uv run sync_outlook_digest.py` 실행. Windows 작업 스케줄러에 이 파일을 등록해 주기적으로 돌림.
- [extract_outlook.py](extract_outlook.py) — (폐기된 방식) 클래식 Outlook COM 자동화로 로컬에서 직접 추출하던 스크립트. 참고용으로 보존.

Licensed under the Apache License, Version 2.0.
