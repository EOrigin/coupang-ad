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


def dismiss_popups(page) -> list:
    """
    대시보드 공지, 정책위반 경고창, 프로모션 팝업, 비밀번호 변경 알림 및 모달 자동 닫기
    - 정책위반 경고창: 필수 체크박스(확인/동의) 자동 체크 후 확인/동의 버튼 클릭
    - 프로모션/마케팅: '오늘 하루 보지 않기', '다시 보지 않기', '다음에 하기', '닫기' 우선 클릭
    - Ant Design 모달 및 일반 다이얼로그 지원 (다중 레이어 팝업 반복 처리)
    - 백드롭(마스크) 제거 및 스크롤 락 해제
    """
    if not page:
        return []
    try:
        if page.is_closed():
            return []
    except Exception:
        return []

    dismissed_all = []

    # 1. 인페이지 JS 기반 다중 모달 자동 탐색 및 해제 (최대 3회 반복)
    for _ in range(3):
        try:
            res = page.evaluate('''() => {
                const dismissed = [];
                // 모달/다이얼로그/레이어 팝업 최상위 컨테이너 셀렉터 (중복 매칭 방지)
                const modalSelectors = [
                    '.ant-modal',
                    '.ant-modal-confirm',
                    '[role="dialog"]',
                    '[role="alertdialog"]',
                    '.modal.show',
                    '.modal-dialog',
                    '.layer-popup',
                    '.wing-modal',
                    '.sc-modal',
                    '.pop-layer',
                    'div[id*="popup" i]',
                    'div[id*="modal" i]',
                    'div[class*="popup" i]:not([style*="display: none"])',
                    'div[class*="modal" i]:not([style*="display: none"])'
                ];

                // 닫기/확인 버튼 우선순위 텍스트 패턴
                const priorityTexts = [
                    // 1순위: 프로모션 보지 않기 / 건너뛰기
                    '오늘 하루 보지 않기',
                    '오늘 하루 동안 보지 않기',
                    '다시 보지 않기',
                    '다시 보지 않음',
                    '다시 열지 않음',
                    '7일간 보지 않기',
                    '일주일간 보지 않기',
                    '다음에 하기',
                    '다음에 변경하기',
                    '다음에 변경',
                    '다음에 보기',
                    '다음에 참여하기',
                    '다음에',
                    '나중에 하기',
                    '나중에',
                    '90일 후 변경',
                    '건너뛰기',
                    '참여하지 않음',
                    '참여 안함',
                    // 2순위: 정책위반/경고/공지 확인 및 동의
                    '확인하였습니다',
                    '확인했습니다',
                    '동의하고 계속하기',
                    '동의합니다',
                    '내용을 확인했습니다',
                    '인지하였습니다',
                    '동의',
                    '계속하기',
                    // 3순위: 일반 닫기 / 취소 / 확인
                    '닫기',
                    '창닫기',
                    '확인하고 닫기',
                    '취소',
                    '확인'
                ];

                const closeButtonSelectors = [
                    '.ant-modal-close',
                    '.ant-modal-close-x',
                    'button[aria-label="Close" i]',
                    'button[aria-label="close" i]',
                    'button[aria-label="닫기"]',
                    '[data-test-id*="close" i]',
                    '[data-test-id*="dismiss" i]',
                    '.modal-close',
                    '.btn-close',
                    '.close-btn',
                    '.close',
                    'a.close',
                    '[data-dismiss="modal"]'
                ];

                const foundModals = document.querySelectorAll(modalSelectors.join(', '));
                foundModals.forEach(modal => {
                    if (!modal || !modal.isConnected) return;

                    const rect = modal.getBoundingClientRect();
                    const style = window.getComputedStyle(modal);
                    if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0' || rect.width <= 0 || rect.height <= 0) {
                        return;
                    }

                    // (A) 모달 내 체크박스(확인/동의/보지않기)가 있다면 먼저 모두 체크
                    const checkboxes = modal.querySelectorAll('input[type="checkbox"]:not(:checked)');
                    checkboxes.forEach(cb => {
                        try {
                            cb.click();
                            cb.checked = true;
                            cb.dispatchEvent(new Event('change', { bubbles: true }));
                            cb.dispatchEvent(new Event('input', { bubbles: true }));
                        } catch (e) {}
                    });
                    const antCbs = modal.querySelectorAll('.ant-checkbox:not(.ant-checkbox-checked), .ant-checkbox-wrapper:not(.ant-checkbox-wrapper-checked)');
                    antCbs.forEach(acb => {
                        try { acb.click(); } catch (e) {}
                    });

                    // (B) 실제 인터랙티브 버튼/링크 요소만 탐색 (div/span 등 컨테이너 오탐 방지)
                    let handled = false;
                    const clickableElements = Array.from(modal.querySelectorAll('button, a, input[type="button"], input[type="submit"], [role="button"], .ant-btn'));

                    for (const target of priorityTexts) {
                        for (const el of clickableElements) {
                            const txt = (el.innerText || el.textContent || el.value || '').trim();
                            if (txt === target || (txt.includes(target) && txt.length <= target.length + 8)) {
                                if (!el.disabled && !el.classList.contains('disabled') && !el.getAttribute('aria-disabled')) {
                                    try {
                                        el.click();
                                        dismissed.push(target);
                                        handled = true;
                                        break;
                                    } catch (e) {}
                                }
                            }
                        }
                        if (handled) break;
                    }

                    // (C) 텍스트 매칭 실패 시 닫기 아이콘 클릭
                    if (!handled) {
                        for (const cSel of closeButtonSelectors) {
                            const closeEl = modal.querySelector(cSel);
                            if (closeEl) {
                                try {
                                    closeEl.click();
                                    dismissed.push('닫기버튼');
                                    handled = true;
                                    break;
                                } catch (e) {}
                            }
                        }
                    }

                    // (D) 모달 확인/푸터 기본 버튼 (ant-modal-confirm 등)
                    if (!handled) {
                        const footerBtn = modal.querySelector('.ant-modal-confirm-btns button.ant-btn-primary, .ant-modal-footer button.ant-btn-primary, .ant-modal-confirm-btns button, .ant-modal-footer button');
                        if (footerBtn && !footerBtn.disabled) {
                            try {
                                const btnText = (footerBtn.innerText || '확인').trim();
                                footerBtn.click();
                                dismissed.push(btnText);
                                handled = true;
                            } catch (e) {}
                        }
                    }
                });

                // (E) 화면에 남은 모달 마스크(백드롭) 정리 및 스크롤 락 해제
                const visibleModals = document.querySelectorAll('.ant-modal:not([style*="display: none"]), [role="dialog"]:not([style*="display: none"]), [role="alertdialog"]:not([style*="display: none"])');
                if (visibleModals.length === 0) {
                    const masks = document.querySelectorAll('.ant-modal-mask, .ant-modal-wrap, .modal-backdrop');
                    masks.forEach(m => {
                        if (!m.querySelector('.ant-modal-content, .modal-content, [role="dialog"], [role="alertdialog"]')) {
                            try { m.remove(); } catch (e) {}
                        }
                    });
                    if (document.body && document.body.style.overflow === 'hidden') {
                        document.body.style.overflow = '';
                    }
                }

                return dismissed;
            }''')
            if res:
                dismissed_all.extend(res)
                page.wait_for_timeout(150)
            else:
                break
        except Exception:
            break

    # 2. Playwright Locator 기반 Fallback (JS 이벤트 인터셉트 미동작 대비)
    fallback_selectors = [
        'button:has-text("오늘 하루 보지 않기")',
        'button:has-text("다시 보지 않기")',
        'button:has-text("다음에 하기")',
        'button:has-text("다음에 변경")',
        'button:has-text("다음에 변경하기")',
        'button:has-text("90일 후 변경")',
        'button:has-text("확인했습니다")',
        'button:has-text("확인하였습니다")',
        'button:has-text("동의하고 계속하기")',
        'button:has-text("동의합니다")',
        'button:has-text("닫기")',
        'button:has-text("확인")',
        '.ant-modal-close',
        'button[aria-label="Close"]',
        'button[aria-label="close"]',
    ]
    for sel in fallback_selectors:
        try:
            btn = page.locator(sel).first
            if btn.is_visible(timeout=100):
                txt = btn.inner_text().strip() if btn.count() > 0 else sel
                btn.click(timeout=1000)
                dismissed_all.append(txt or sel)
                page.wait_for_timeout(150)
        except Exception:
            pass

    if dismissed_all:
        unique_dismissed = list(dict.fromkeys(dismissed_all))
        print(f"  [팝업 자동 닫기] 정책위반/프로모션/공지 팝업을 닫았습니다: {', '.join(unique_dismissed)}", flush=True)

    return dismissed_all


def close_extra_popup_pages(context, main_page=None):
    """
    브라우저 컨텍스트에 window.open 등으로 열린 불필요한 공지/프로모션/경고 팝업 창 자동 닫기
    - 메인 광고 대시보드 페이지는 보호
    - 로그인/인증 창은 인증 완료 전까지 보호
    """
    if not context:
        return
    try:
        pages = context.pages
        if len(pages) <= 1:
            return
        for p in pages:
            if main_page and p == main_page:
                continue
            if p.is_closed():
                continue
            url = (p.url or "").lower()
            # 로그인/인증 진행 중인 탭은 닫지 않음
            if any(k in url for k in ["xauth.coupang.com", "/user/login", "keycloak_callback"]):
                continue

            title = ""
            try:
                title = p.title() or ""
            except Exception:
                pass

            popup_indicators = ["popup", "notice", "event", "promotion", "warning", "policy", "banner"]
            is_popup = any(ind in url for ind in popup_indicators) or any(ind in title for ind in ["공지", "안내", "이벤트", "프로모션", "경고", "정책"])

            # 팝업 표시가 있거나, 메인 페이지가 지정되어 있고 메인 대시보드가 아닌 별도 창인 경우
            if is_popup or (main_page and p != main_page and "/marketing/dashboard" not in url):
                try:
                    dismiss_popups(p)
                except Exception:
                    pass
                p_desc = title or url[:50] or "별도 팝업창"
                print(f"  [팝업창 자동 닫기] 새 창으로 열린 팝업을 닫았습니다: {p_desc}", flush=True)
                try:
                    p.close()
                except Exception:
                    pass
    except Exception:
        pass


def setup_page_popup_handlers(page):
    """
    Playwright 페이지에 대화상자(alert/confirm) 자동 수락 및 add_locator_handler 등록
    크롤링 액션 수행 중 불시에 팝업이 뜨더라도 자동 감지하여 닫고 크롤링을 계속 진행
    """
    if not page or page.is_closed():
        return

    # 1. 브라우저 네이티브 alert / confirm 대화상자 자동 수락
    try:
        def on_dialog(dialog):
            print(f"  [알림창 자동 처리] {dialog.type}: {dialog.message}", flush=True)
            try:
                dialog.accept()
            except Exception:
                pass
        page.on("dialog", on_dialog)
    except Exception:
        pass

    # 2. Playwright add_locator_handler 등록 (액션 수행 중 팝업 발생 시 자동 해제)
    try:
        modal_locator = page.locator('.ant-modal, [role="dialog"], [role="alertdialog"], .modal-dialog').first
        page.add_locator_handler(
            modal_locator,
            lambda _: dismiss_popups(page)
        )
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

        # 3. 쿠키 기반: 판매자 sc_vid 쿠키가 정상적으로 존재하는 경우 (윙 도메인에 있을 때만)
        if context and "wing.coupang.com" in host:
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

    # 팝업 및 대화상자 이벤트 리스너 등록
    if context:
        try:
            context.on("page", lambda new_p: setup_page_popup_handlers(new_p))
        except Exception:
            pass

    # 초기 페이지 접근 시도
    current_page = page
    try:
        setup_page_popup_handlers(current_page)
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

        # 1. 팝업 / 공지 닫기 및 핸들러 연결
        for p in pages:
            setup_page_popup_handlers(p)
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

                # 로그인 완료 후 불필요한 공지/경고 팝업 창 정리
                close_extra_popup_pages(context, main_page=p)
                dismiss_popups(p)
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
            # 윙 페이지 또는 완료 상태의 탭 찾기 (xauth 인증 진행 중인 페이지는 방해하지 않음)
            wing_or_portal_page = None
            for p in pages:
                if is_wing_logged_in(p, context) or is_ads_login_portal(p):
                    wing_or_portal_page = p
                    break

            if wing_or_portal_page:
                auto_forward_count += 1
                seller_info = f" (판매자ID: {c_status['sc_vid']})" if c_status['sc_vid'] else ""
                print(f"\n[로그인 감지] 쿠팡 판매자 인증 확인{seller_info} -> 광고 대시보드 세션 자동 연동 중... ({auto_forward_count}/{max_auto_forwards})", flush=True)
                try:
                    wing_or_portal_page.goto(ADS_WING_LOGIN_URL, timeout=30000, wait_until="domcontentloaded")
                    wing_or_portal_page.wait_for_timeout(2000)
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
