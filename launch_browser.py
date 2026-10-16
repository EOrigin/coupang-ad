#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
쿠팡 로그인 전용 브라우저 실행 스크립트 (npm run login)
- 브라우저 창을 띄워 사용자가 직접 안전하게 로그인할 수 있도록 지원
- 쿠키, URL, DOM 요소를 다각도로 실시간 감지하여 로그인 완료 즉시 감지
- 로그인 완료 시 세션을 session.json 및 browser_profile에 영구 저장 후 안전하게 종료
"""

import os
import json
import time
from playwright.sync_api import sync_playwright
from login_manager import (
    PROFILE_DIR,
    ADS_WING_LOGIN_URL,
    save_session,
    load_session,
    dismiss_popups,
    close_extra_popup_pages,
    setup_page_popup_handlers,
    wait_and_detect_login,
)

LOG_FILE = os.path.abspath("./captured_api.jsonl")


def handle_response(response):
    try:
        url = response.url
        content_type = response.headers.get("content-type", "")
        if "application/json" in content_type:
            # 광고 관련 유의미한 API 요청 필터링
            keywords = ["ad", "report", "campaign", "product", "item", "wing", "stat", "search", "metric", "roas", "performance"]
            if any(k in url.lower() for k in keywords):
                try:
                    data = response.json()
                    with open(LOG_FILE, "a", encoding="utf-8") as f:
                        record = {
                            "timestamp": time.time(),
                            "url": url,
                            "status": response.status,
                            "data": data,
                        }
                        f.write(json.dumps(record, ensure_ascii=False) + "\n")
                except Exception:
                    pass
    except Exception:
        pass


def main():
    os.makedirs(PROFILE_DIR, exist_ok=True)
    print("\n" + "=" * 60, flush=True)
    print("쿠팡 로그인 전용 브라우저를 실행합니다.", flush=True)
    print(f"- 프로필 디렉토리: {PROFILE_DIR}", flush=True)
    print(f"- CDP 디버깅 포트: 9222", flush=True)
    print("=" * 60 + "\n", flush=True)

    with sync_playwright() as p:
        context = None
        browser_managed_by_us = True

        # 1. 기존 실행 중인 브라우저(포트 9222)가 있는지 확인
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
            browser_managed_by_us = False
            print("기존에 열려있는 브라우저(포트 9222)에 성공적으로 연결되었습니다.", flush=True)
        except Exception:
            # 신규 브라우저 실행
            context = p.chromium.launch_persistent_context(
                user_data_dir=PROFILE_DIR,
                headless=False,
                args=[
                    "--remote-debugging-port=9222",
                    "--start-maximized",
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                ],
                no_viewport=True,
            )

        load_session(context)

        def setup_page(pg):
            pg.on("response", handle_response)
            setup_page_popup_handlers(pg)

        context.on("page", setup_page)
        page = context.pages[0] if context.pages else context.new_page()
        setup_page(page)

        # 쿠팡 광고 판매자 로그인 페이지로 이동
        print(f"로그인 페이지 이동 중: {ADS_WING_LOGIN_URL}", flush=True)
        try:
            page.goto(ADS_WING_LOGIN_URL, wait_until="domcontentloaded")
        except Exception as e:
            print(f"페이지 이동 안내: {e}", flush=True)

        print("\n" + "=" * 60, flush=True)
        print("★ 쿠팡 로그인 창이 열렸습니다. ★", flush=True)
        print("1. 브라우저에서 아이디와 비밀번호를 입력하고 로그인을 완료해 주세요.", flush=True)
        print("2. 로그인이 완료되면 자동으로 감지하여 세션을 저장하고 창을 닫습니다.", flush=True)
        print("=" * 60 + "\n", flush=True)

        try:
            detected_page = wait_and_detect_login(page, context, timeout_sec=300, is_headless=False)
            if detected_page:
                save_session(context)
                print("세션 저장이 완료되었습니다. 다음 명령어를 실행할 수 있습니다 (예: npm run crawl).\n", flush=True)
                time.sleep(2)
            else:
                print("\n브라우저 창이 닫혀 프로그램을 종료합니다.", flush=True)
        except KeyboardInterrupt:
            print("\n[알림] 사용자에 의해 중단되었습니다.", flush=True)
            save_session(context)
        except Exception as e:
            print(f"\n[알림] 로그인 대기 중 오류 발생: {e}", flush=True)
        finally:
            if browser_managed_by_us and context:
                try:
                    context.close()
                except Exception:
                    pass


if __name__ == "__main__":
    main()
