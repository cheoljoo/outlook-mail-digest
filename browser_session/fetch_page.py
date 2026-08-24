# -*- coding: utf-8 -*-
"""
start_login_session.py 로 저장해 둔 로그인 세션을 재사용해서, headless로 페이지에
접속하고 스크린샷/HTML을 남깁니다. 세션이 만료되면(로그인 화면으로 리다이렉트되면)
start_login_session.py를 다시 한 번 실행해서 재로그인해야 합니다.

사용법:
  .venv/bin/python fetch_page.py "<URL>" [출력파일 접두사]
"""

import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

PROFILE_DIR = Path.home() / ".secure_browser_profile"
OUT_DIR = Path(__file__).resolve().parent / "out"


def main() -> None:
    if len(sys.argv) < 2:
        print("사용법: fetch_page.py <URL> [출력파일 접두사]")
        sys.exit(1)

    url = sys.argv[1]
    prefix = sys.argv[2] if len(sys.argv) > 2 else "page"
    OUT_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=True,
        )
        page = context.new_page()
        page.goto(url, wait_until="networkidle", timeout=30000)

        screenshot_path = OUT_DIR / f"{prefix}.png"
        html_path = OUT_DIR / f"{prefix}.html"
        page.screenshot(path=str(screenshot_path), full_page=True)
        html_path.write_text(page.content(), encoding="utf-8")

        print(f"현재 URL: {page.url}")
        print(f"스크린샷: {screenshot_path}")
        print(f"HTML: {html_path}")

        if "login.microsoftonline.com" in page.url or "adfs" in page.url.lower():
            print("[WARN] 로그인 페이지로 리다이렉트된 것으로 보입니다. "
                  "세션이 만료됐을 수 있으니 start_login_session.py를 다시 실행하세요.")

        context.close()


if __name__ == "__main__":
    main()
