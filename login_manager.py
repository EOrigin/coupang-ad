#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
쿠팡 로그인 세션 관리 및 인증 상태 감지 공통 모듈
- 쿠키(CAP_AUTH_SESSION, sc_vid, KEYCLOAK 등) 기반 정밀 인증 감지
- URL(경로 및 도메인) 및 DOM 요소(로그아웃 버튼, 대시보드 테이블 등) 다각도 검사
- 쿠팡 윙/xauth 로그인 후 광고 대시보드 자동 연동
- 비밀번호 변경 팝업, AML, 공지사항 모달 자동 닫기
- 세션 영구 보존 및 복원 (session.json)
"""

import os
import sys
import json
import time
import urllib.parse

SESSION_FILE = os.path.abspath("./session.json")
PROFILE_DIR = os.path.abspath("./browser_profile")
ADS_DASHBOARD_URL = "https://advertising.coupang.com/marketing/dashboard/sales"
ADS_WING_LOGIN_URL = "https://advertising.coupang.com/user/login?_cap_client=WING&returnUrl=%2Fmarketing%2Fdashboard%2Fsales"
WING_URL = "https://wing.coupang.com"


def save_session(context):
    """세션 쿠키 및 스토리지를 session.json에 영구 보존 저장"""
    if not context:
        return False
    try:
        context.storage_state(path=SESSION_FILE)
        cookies = context.cookies()
        persist_cookies = []
        for c in cookies:
            if c.get("expires", 0) <= 0:
                c_copy = dict(c)
                c_copy["expires"] = int(time.time()) + 86400 * 30  # 30일 보존
                persist_cookies.append(c_copy)
        if persist_cookies:
            context.add_cookies(persist_cookies)
            context.storage_state(path=SESSION_FILE)
        return True
    except Exception as e:
        print(f"  [세션 저장 안내] {e}", flush=True)
        return False


def load_session(context):
    """저장된 session.json에서 세션 복원"""
    if not context or not os.path.exists(SESSION_FILE):
        return False
    try:
        with open(SESSION_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
        cookies = state.get("cookies", [])
        if cookies:
            for c in cookies:
                if c.get("expires", 0) <= 0:
                    c["expires"] = int(time.time()) + 86400 * 30
            context.add_cookies(cookies)
            return True
    except Exception as e:
        print(f"  [세션 로드 안내] {e}", flush=True)
    return False


def dismiss_popups(page):
    """대시보드 공지, 비밀번호 변경 알림 및 팝업 모달 자동 닫기"""
    if not page:
        return
    popup_selectors = [
        # 비밀번호 다음에 변경하기
        'button:has-text("다음에 변경")',
        'a:has-text("다음에 변경")',
        'button:has-text("다음에 변경하기")',
        'button:has-text("90일 후 변경")',
        'button:has-text("다음에 하기")',
        # 오늘 하루 보지 않기 / 공지
        'button:has-text("오늘 하루 보지 않기")',
        'button:has-text("다시 보지 않기")',
        'button:has-text("닫기")',
        'button:has-text("확인")',
        '.ant-modal-close',
        'button[aria-label="Close"]',
        'button[aria-label="close"]',
    ]
    for sel in popup_selectors:
        try:
            btn = page.locator(sel).first
            if btn.is_visible(timeout=250):
                btn.click()
                page.wait_for_timeout(200)
        except Exception:
            pass


def get_auth_cookies_status(context):
    """
    Playwright 컨텍스트 쿠키 기반 인증 상태 상세 진단
    """
    if not context:
        return {
            "has_ads_session": False,
            "has_wing_session": False,
            "has_xauth_session": False,
            "sc_vid": "",
            "cookie_count": 0,
        }
    try:
        cookies = context.cookies()
        c_dict = {c.get("name"): c for c in cookies}
        c_names = set(c_dict.keys())

        # 1. 광고 센터 인증 쿠키
        # CAP_AUTH_SESSION이 핵심, aid + cmgbdgw-sid 조합도 광고 인증 세션임
        has_ads_session = ("CAP_AUTH_SESSION" in c_names) or (
            "aid" in c_names and "cmgbdgw-sid" in c_names
        )

        # 2. 쿠팡 윙/판매자 인증 쿠키
        # sc_vid (판매자 ID 예: A01547111) 및 sc_uid
        sc_vid = c_dict.get("sc_vid", {}).get("value", "")
        has_wing_session = bool("sc_vid" in c_names or "sc_uid" in c_names)

        # 3. xauth SSO 인증 쿠키
        has_xauth_session = bool(
            "KEYCLOAK_IDENTITY" in c_names
            or "KEYCLOAK_SESSION" in c_names
            or "AUTH_SESSION_ID" in c_names
        )

        return {
            "has_ads_session": has_ads_session,
            "has_wing_session": has_wing_session,
            "has_xauth_session": has_xauth_session,
            "sc_vid": sc_vid,
            "cookie_count": len(cookies),
        }
    except Exception:
        return {
            "has_ads_session": False,
            "has_wing_session": False,
            "has_xauth_session": False,
            "sc_vid": "",
            "cookie_count": 0,
        }


def is_active_login_form(page) -> bool:
    """현재 페이지에 활성화된 로그인 입력 폼(아이디/비밀번호 필드)이 있는지 확인"""
    try:
        if page.locator("input[type='password']:visible").count() > 0:
            return True
        if page.locator("input[name='password']:visible").count() > 0:
            return True
    except Exception:
        pass
    return False


def is_ads_logged_in(page, context=None) -> bool:
    """
    쿠팡 광고 센터 로그인 완료 여부 종합 판별
    - URL 경로 기반 확인 (쿼리스트링 오탐 배제)
    - 쿠키 기반 확인 (CAP_AUTH_SESSION 등)
    - DOM 셀렉터 기반 확인 (대시보드 캠페인 목록, 테이블, 로그아웃 버튼 등)
    """
    if not page:
        return False

    try:
        url = page.url or ""
        parsed = urllib.parse.urlparse(url)
        host = (parsed.hostname or "").lower()
        path = parsed.path.lower()

        # 1. 활성 비밀번호 입력창이 있다면 아직 로그인 폼 단계임
        if is_active_login_form(page):
            return False

        # 2. 광고 센터 도메인인 경우
        if "advertising.coupang.com" in host:
            # 명시적인 로그인 페이지 진입 중인 경우
            if path in ["/user/login", "/login_sxauth"]:
                return False

            # (A) 광고 대시보드 핵심 경로 진입 완료
            dashboard_paths = [
                "/marketing",
                "/dashboard",
                "/campaign",
                "/home",
                "/report",
                "/from-wing-lnb",
            ]
            if any(path.startswith(dp) or dp in path for dp in dashboard_paths):
                return True

            # (B) 대시보드 DOM 요소 존재 확인
            try:
                if (
                    page.locator('[data-test-id="pa-cam-dd-name-value"]').count() > 0
                    or page.locator(".rt-tbody").count() > 0
                    or page.locator('a[href*="logout"]').count() > 0
                    or page.locator('button:has-text("로그아웃")').count() > 0
                ):
                    return True
            except Exception:
                pass

        # 3. 쿠키 기반 추가 검증: 광고 세션 쿠키가 이미 발급되어 있는 경우
        if context:
            c_status = get_auth_cookies_status(context)
            if c_status["has_ads_session"] and "advertising.coupang.com" in host:
                if path not in ["/user/login", "/login_sxauth"]:
                    return True

    except Exception:
        pass

    return False


def is_wing_logged_in(page, context=None) -> bool:
    """쿠팡 윙 판매자 센터 로그인 여부 확인"""
    if not page:
        return False

    try:
        url = page.url or ""
        parsed = urllib.parse.urlparse(url)
        host = (parsed.hostname or "").lower()
        path = parsed.path.lower()

        # 1. 활성 비밀번호 입력창이 있다면 미로그인
        if is_active_login_form(page):
            return False

        # 2. 윙 도메인인 경우
        if "wing.coupang.com" in host:
            login_paths = ["/login", "/sso/login", "/signin", "/mfa", "/otp", "/verify"]
            if any(lp in path for lp in login_paths):
                return False

            title = page.title()
            if (
                any(t in title for t in ["Coupang Wing", "윙", "고객정보확인제도", "AML"])
                or path in ["", "/", "/seller", "/seller/home"]
                or "/sc-ui" in path
            ):
                return True

            try:
                if (
                    page.locator('a[href*="logout"]').count() > 0
                    or page.locator(".wing-seller-name").count() > 0
                    or page.locator("#header-wing-seller-name").count() > 0
                ):
                    return True
            except Exception:
                pass

        # 3. 쿠키 기반: 판매자 sc_vid 쿠키가 정상적으로 존재하는 경우
        if context:
            c_status = get_auth_cookies_status(context)
            if c_status["has_wing_session"] and not is_active_login_form(page):
                return True

    except Exception:
        pass

    return False


def is_ads_login_portal(page) -> bool:
    """쿠팡 광고 로그인 선택 포털 화면 (오픈마켓 판매자 / 로켓배송 공급사 / 광고전용 선택 화면)"""
    if not page:
        return False
    try:
        url = page.url or ""
        parsed = urllib.parse.urlparse(url)
        host = (parsed.hostname or "").lower()
        path = parsed.path.lower()
        query = parsed.query

        if "advertising.coupang.com" in host and path == "/user/login":
            if "_cap_client=" not in query and "client_id=" not in query:
                return True
    except Exception:
        pass
    return False


def wait_and_detect_login(page, context, timeout_sec=300, is_headless=False, on_status_update=None):
    """
    브라우저 컨텍스트의 모든 탭과 쿠키를 지속적으로 모니터링하여 로그인을 감지하고
    광고 대시보드로 안전하게 진입하도록 처리하는 메인 대기 함수
    """
    start_time = time.time()
    last_feedback_time = 0
    auto_forward_count = 0
    max_auto_forwards = 5

    # 초기 페이지 접근 시도
    current_page = page
    try:
        current_url = current_page.url or ""
        if not current_url or current_url == "about:blank":
            current_page.goto(ADS_DASHBOARD_URL, timeout=30000, wait_until="domcontentloaded")
    except Exception as e:
        print(f"  [안내] 초기 페이지 이동 알림: {e}", flush=True)

    while time.time() - start_time < timeout_sec:
        pages = context.pages if context and context.pages else [current_page]
        if not pages:
            print("  [알림] 브라우저 창이 모두 닫혔습니다.", flush=True)
            return None

        # 1. 팝업 / 공지 닫기
        for p in pages:
            dismiss_popups(p)

        # 2. 쿠키 상태 분석
        c_status = get_auth_cookies_status(context)

        # 3. 모든 열린 탭에서 광고 대시보드 로그인 여부 확인
        for p in pages:
            if is_ads_logged_in(p, context):
                # 광고 세션 확인 완료!
                save_session(context)
                dismiss_popups(p)

                # 만약 아직 매출 대시보드 경로가 아니라면 이동
                if "/marketing/dashboard/sales" not in p.url:
                    print(f"\n[로그인 감지] 광고 세션 확인 완료! 매출 성장 대시보드로 이동합니다...", flush=True)
                    try:
                        p.goto(ADS_DASHBOARD_URL, timeout=30000, wait_until="domcontentloaded")
                        p.wait_for_timeout(1500)
                        dismiss_popups(p)
                    except Exception:
                        pass

                save_session(context)
                print("\n" + "=" * 60, flush=True)
                print("[성공] 쿠팡 광고 센터 로그인이 정상 감지되었습니다!", flush=True)
                if c_status.get("sc_vid"):
                    print(f"- 감지된 판매자(벤더) ID: {c_status['sc_vid']}", flush=True)
                print(f"- 현재 페이지 URL: {p.url}", flush=True)
                print("=" * 60 + "\n", flush=True)
                return p

        # 헤드리스(백그라운드) 모드인데 세션이 없는 경우 안내 후 조기 종료
        if is_headless and (time.time() - start_time > 8):
            if not c_status["has_ads_session"]:
                print("\n" + "=" * 65, flush=True)
                print("[오류] 백그라운드(Headless) 모드에서는 브라우저 화면이 보이지 않아 수동 로그인을 진행할 수 없습니다.", flush=True)
                print("현재 저장된 로그인 세션이 없거나 만료되었습니다.")
                print("화면이 표시되는 일반 모드로 먼저 1회 로그인하여 세션을 저장해 주세요:")
                print("  ▶ 1회 로그인 실행: npm run login  또는  npm run crawl", flush=True)
                print("=" * 65 + "\n", flush=True)
                sys.exit(1)

        # 4. 광고 로그인 선택 포털(3개 버튼 화면)인 경우 오픈마켓 판매자 로그인으로 자동 연결
        for p in pages:
            if is_ads_login_portal(p):
                print("\n[안내] 광고 로그인 화면 감지: 오픈마켓/로켓그로스 판매자 로그인으로 자동 연결합니다...", flush=True)
                try:
                    wing_btn = p.locator('a[href*="_cap_client=WING"]').first
                    if wing_btn.is_visible(timeout=1000):
                        wing_btn.click()
                    else:
                        p.goto(ADS_WING_LOGIN_URL, timeout=30000, wait_until="domcontentloaded")
                except Exception:
                    p.goto(ADS_WING_LOGIN_URL, timeout=30000, wait_until="domcontentloaded")
                break

        # 5. 쿠팡 윙/xauth 로그인 완료 상태이나 광고 대시보드로 아직 안 넘어간 경우 자동 연동
        # (판매자 sc_vid 쿠키 또는 xauth 쿠키가 있고 비밀번호 입력창이 없는 경우)
        can_forward_to_ads = (
            (c_status["has_wing_session"] or c_status["has_xauth_session"])
            and not any(is_active_login_form(p) for p in pages)
        )
        if can_forward_to_ads and auto_forward_count < max_auto_forwards:
            # 윙 페이지 또는 완료 상태의 탭 찾기
            wing_or_portal_page = None
            for p in pages:
                if is_wing_logged_in(p, context) or "xauth.coupang.com" in (p.url or "") or is_ads_login_portal(p):
                    wing_or_portal_page = p
                    break

            if wing_or_portal_page:
                auto_forward_count += 1
                seller_info = f" (판매자ID: {c_status['sc_vid']})" if c_status['sc_vid'] else ""
                print(f"\n[로그인 감지] 쿠팡 판매자 인증 확인{seller_info} -> 광고 대시보드 세션 자동 연동 중... ({auto_forward_count}/{max_auto_forwards})", flush=True)
                try:
                    wing_or_portal_page.goto(ADS_DASHBOARD_URL, timeout=30000, wait_until="domcontentloaded")
                    wing_or_portal_page.wait_for_timeout(2000)
                except Exception:
                    try:
                        wing_or_portal_page.goto(ADS_WING_LOGIN_URL, timeout=30000, wait_until="domcontentloaded")
                    except Exception:
                        pass

        # 6. 실시간 사용자 피드백 (대기 상태 및 현재 열린 탭 안내)
        now = time.time()
        if now - last_feedback_time >= 5:
            urls_summary = " | ".join([p.url[:60] + "..." if len(p.url or "") > 60 else (p.url or "empty") for p in pages])
            auth_tags = []
            if c_status["has_ads_session"]:
                auth_tags.append("광고세션:O")
            if c_status["has_wing_session"]:
                auth_tags.append(f"윙세션:O({c_status['sc_vid'] or '인증'})")
            if c_status["has_xauth_session"]:
                auth_tags.append("xauth인증:O")
            auth_str = f" [{', '.join(auth_tags)}]" if auth_tags else ""

            print(f"  [로그인 대기 중...{auth_str}] 열린 탭 {len(pages)}개: {urls_summary}", flush=True)
            last_feedback_time = now

        time.sleep(1)

    raise TimeoutError("로그인 대기 시간(300초)이 초과되었습니다.")
