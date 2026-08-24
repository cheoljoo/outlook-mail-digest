# -*- coding: utf-8 -*-
"""
start_login_session.py 로 저장해 둔 로그인 세션을 재사용해서, SharePoint/OneDrive 공유
링크로부터 실제 파일(.xlsx 등) 자체를 다운로드합니다. 화면 스크린샷이 아니라 진짜 바이너리
파일을 받으므로 openpyxl 등으로 바로 파싱할 수 있습니다.

동작 원리: SharePoint 공유 링크에 &download=1 을 붙이면 뷰어 대신 파일 다운로드가
트리거되는 경우가 많습니다. 이 스크립트는 그 방식을 먼저 시도하고, 실패하면(다운로드
이벤트가 안 잡히면) 원래 URL을 그대로 열어서 무슨 일이 있었는지 스크린샷/HTML로 남깁니다.

사용법:
  .venv/bin/python download_file.py "<공유 URL>" <저장할 파일명>
"""

import sys
import urllib.parse
from pathlib import Path

from playwright.sync_api import sync_playwright

PROFILE_DIR = Path.home() / ".secure_browser_profile"
OUT_DIR = Path(__file__).resolve().parent / "out"


def _with_download_param(url: str) -> str:
    parts = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qs(parts.query)
    query["download"] = ["1"]
    new_query = urllib.parse.urlencode(query, doseq=True)
    return urllib.parse.urlunsplit(
        (parts.scheme, parts.netloc, parts.path, new_query, parts.fragment)
    )


def main() -> None:
    if len(sys.argv) < 3:
        print("사용법: download_file.py <공유 URL> <저장할 파일명>")
        sys.exit(1)

    url = sys.argv[1]
    out_name = sys.argv[2]
    OUT_DIR.mkdir(exist_ok=True)
    download_url = _with_download_param(url)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=True,
            accept_downloads=True,
        )
        page = context.new_page()

        try:
            with page.expect_download(timeout=15000) as download_info:
                page.goto(download_url, wait_until="commit", timeout=30000)
            download = download_info.value
            target = OUT_DIR / out_name
            download.save_as(str(target))
            print(f"[OK] 다운로드 완료: {target} ({target.stat().st_size} bytes)")
        except Exception as e:
            print(f"[WARN] 다운로드 이벤트가 발생하지 않았습니다 ({e!r}).")
            page.goto(url, wait_until="networkidle", timeout=30000)
            fallback_png = OUT_DIR / f"{out_name}.fallback.png"
            page.screenshot(path=str(fallback_png), full_page=True)
            print(f"대신 현재 화면을 저장했습니다: {fallback_png}")
            print(f"현재 URL: {page.url}")

            blocked_msg = "조직에서 이 장치를 사용한 다운로드"
            for fr in page.frames:
                try:
                    if fr.get_by_text(blocked_msg).count() > 0:
                        print(
                            "[INFO] 조직 정책(Conditional Access/DLP)이 이 장치(관리되지 않는 "
                            "기기로 인식됨)에서 다운로드/인쇄/동기화를 명시적으로 차단하고 "
                            "있습니다. 파일 자체를 받는 것은 이 방식으로는 불가능하고, "
                            "화면을 보고 읽는(스크린샷/HTML) 용도로만 쓸 수 있습니다."
                        )
                        break
                except Exception:
                    pass

        context.close()


if __name__ == "__main__":
    main()
