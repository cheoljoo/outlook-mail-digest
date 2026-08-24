# browser_session — 사내 로그인 세션을 유지하는 headless 브라우저

이 서버(`tiger`)는 여러 사람이 같이 쓰는 **공용 서버**입니다. 그래서 로그인 세션(쿠키)은
반드시 이 저장소 **밖**, 소유자 본인만 읽을 수 있는 `~/.secure_browser_profile`
(`chmod 700`)에 저장합니다. 이 디렉터리와 그 안의 세션은 **본인 계정 로그인 상태와 동일한
권한**을 가지므로 비밀번호처럼 취급하세요 — 절대 git에 커밋하거나 다른 사람과 공유하지 않습니다.

## 준비 (이미 완료됨)

- `.venv/`: `uv venv --python 3.11` + `uv pip install playwright` (root/sudo 불필요)
- Chromium 바이너리: `.venv/bin/playwright install chromium` 로 `~/.cache/ms-playwright/`에 설치됨
- `~/.secure_browser_profile`: `chmod 700`으로 생성된 로그인 세션 저장 경로

## 1. 최초 1회 로그인

```bash
cd /data01/cheoljoo.lee/code/outlook-mail-digest/browser_session
.venv/bin/python start_login_session.py
```

터미널에 뜨는 안내대로, Windows에서 SSH 로컬 포트포워딩 + `chrome://inspect`로 접속해서
실제로 로그인(SSO/MFA 포함)합니다. **비밀번호/OTP는 그 DevTools 창에만 입력하고, 이 터미널이나
Claude와의 채팅에는 절대 입력하지 마세요.**

## 2. 이후 재사용 (headless)

```bash
.venv/bin/python fetch_page.py "<URL>" <출력파일 접두사>
```

`out/<접두사>.png`(스크린샷), `out/<접두사>.html`(렌더링된 HTML)을 남깁니다.
로그인 화면으로 리다이렉트되면(세션 만료) 1번을 다시 실행하세요.

## 보안 메모

- 디버그 포트(`127.0.0.1:9222`)는 loopback에만 바인딩되어, SSH 접속 없이는 이 서버 자체에서도
  네트워크로 닿지 않습니다.
- `out/` 폴더의 스크린샷/HTML에는 사내 민감정보가 그대로 담길 수 있으므로 `.gitignore`에서
  제외되어 있고, 필요 없어지면 직접 정리하세요.
