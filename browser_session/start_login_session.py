# -*- coding: utf-8 -*-
"""
사내 SSO 로그인이 필요한 페이지(SharePoint 등)에 최초 1회 로그인하기 위한 헬퍼.

이 스크립트는 headless Chromium을 원격 디버깅 포트(기본 9222)를 열어서 실행합니다.
포트는 반드시 127.0.0.1(loopback)에만 바인딩되므로, 이 서버 자체에 접속 가능한 사람이라도
SSH 세션 없이는 이 포트에 닿을 수 없습니다.

사용 흐름:
1. 이 스크립트를 실행 (이 창은 로그인이 끝날 때까지 열어둡니다)
2. Windows PowerShell에서 별도로 SSH 로컬 포트포워딩 실행:
     ssh -L 9222:localhost:9222 cheoljoo.lee@<tiger 접속 주소>
   (이 창도 로그인이 끝날 때까지 열어둡니다)
3. Windows Chrome에서 주소창에 chrome://inspect/#devices 입력
   → "Configure..." 클릭 → localhost:9222 추가
   → 페이지가 "Remote Target" 목록에 나타나면 "inspect" 클릭
4. 열린 DevTools 창에서 실제로 마우스/키보드로 로그인 (SSO/MFA 포함) 진행
   ⚠️ 비밀번호/OTP는 이 창에만 입력하세요. Claude와의 대화창에는 절대 입력하지 마세요.
5. 로그인이 끝나 원하는 페이지(OutlookDigest.xlsx 등)가 보이면, 이 스크립트가 실행 중인
   터미널에서 Enter를 눌러 세션을 저장하고 종료합니다.

세션(쿠키 등)은 --- 이 계정으로 로그인한 것과 동일한 권한 ---이므로 profile 디렉터리는
소유자 본인만 읽을 수 있도록 권한이 잠겨 있어야 합니다 (setup에서 이미 chmod 700 처리).
"""

import subprocess
import sys
from pathlib import Path

PROFILE_DIR = Path.home() / ".secure_browser_profile"
DEBUG_PORT = 9222


def find_chromium_executable() -> str:
    cache_dir = Path.home() / ".cache" / "ms-playwright"
    candidates = sorted(cache_dir.glob("chromium-*/chrome-linux*/chrome"))
    if not candidates:
        print("[ERROR] chromium 실행 파일을 찾을 수 없습니다. "
              "browser_session/.venv/bin/playwright install chromium 을 먼저 실행하세요.")
        sys.exit(1)
    return str(candidates[-1])


def main() -> None:
    if not PROFILE_DIR.exists():
        print(f"[ERROR] 프로필 디렉터리가 없습니다: {PROFILE_DIR}")
        sys.exit(1)

    chromium = find_chromium_executable()
    print(f"[INFO] Chromium 실행: {chromium}")
    print(f"[INFO] 프로필: {PROFILE_DIR}")
    print(f"[INFO] 디버그 포트: 127.0.0.1:{DEBUG_PORT} (loopback 전용)")

    proc = subprocess.Popen(
        [
            chromium,
            "--headless=new",
            f"--remote-debugging-port={DEBUG_PORT}",
            "--remote-debugging-address=127.0.0.1",
            f"--user-data-dir={PROFILE_DIR}",
            "--no-first-run",
            "about:blank",
        ]
    )

    print()
    print("=" * 70)
    print("이제 Windows에서:")
    print(f"  1) ssh -L {DEBUG_PORT}:localhost:{DEBUG_PORT} cheoljoo.lee@<tiger 주소>")
    print("  2) Windows Chrome 주소창에 chrome://inspect/#devices")
    print(f"  3) Configure -> localhost:{DEBUG_PORT} 추가 -> Remote Target에서 inspect 클릭")
    print("  4) 그 창에서 실제로 로그인 (비밀번호/OTP는 절대 이 터미널이나")
    print("     Claude 채팅에 입력하지 마세요)")
    print("=" * 70)
    input("\n로그인을 마쳤으면 여기서 Enter를 눌러 종료하세요 (세션은 이미 디스크에 저장되어 있습니다)...")

    proc.terminate()
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        proc.kill()
    print("[INFO] 종료했습니다. 이제 fetch_page.py 로 헤드리스 재사용이 가능합니다.")


if __name__ == "__main__":
    main()
