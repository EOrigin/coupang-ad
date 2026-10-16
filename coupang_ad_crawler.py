#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
쿠팡 광고(AI 스마트광고) 캠페인 일괄 수집 크롤러
- 집행 광고비 기준(기본: 5,000원 이상)
- 광고 수익률 기준(기본: 600% 미만)
- 모든 스마트광고 캠페인을 자동으로 순회하여 해당 상품의 옵션ID 및 상세 정보 추출
- CSV, JSON, TXT(옵션ID 목록) 자동 저장
"""

import os
import re
import sys
import json
import time
import argparse
from datetime import datetime
from playwright.sync_api import sync_playwright
from excel_exporter import export_to_excel
from login_manager import (
    SESSION_FILE,
    ADS_DASHBOARD_URL,
    save_session,
    load_session,
    dismiss_popups,
    close_extra_popup_pages,
    setup_page_popup_handlers,
    wait_and_detect_login,
)

DASHBOARD_URL = ADS_DASHBOARD_URL

def parse_int(text: str) -> int:
    clean = re.sub(r'[^\d]', '', str(text))
    return int(clean) if clean else 0

def parse_float(text: str) -> float:
    clean = str(text).replace(',', '').replace('%', '').strip()
    try:
        return float(clean)
    except:
        return 0.0

def wait_for_login(page, context=None, timeout_sec=300, is_headless=False):
    """로그인 상태 확인 및 광고 대시보드 자동 진입 (login_manager 연동)"""
    return wait_and_detect_login(page, context, timeout_sec=timeout_sec, is_headless=is_headless)


def wait_table_loaded(page, timeout_ms=5000):
    """React Table 로딩 스피너 및 오버레이(-loading -active) 대기 및 돌발 팝업 자동 닫기"""
    dismiss_popups(page)
    try:
        page.locator('div.-loading.-active').wait_for(state='detached', timeout=timeout_ms)
    except Exception:
        pass
    dismiss_popups(page)
    try:
        page.wait_for_timeout(300)
    except Exception:
        pass


def extract_products_from_campaign(page, campaign_name: str, min_cost: int, max_roas: float):
    """현재 열려있는 캠페인 상세 페이지에서 모든 페이지를 순회하며 조건에 맞는 상품을 추출"""
    dismiss_popups(page)
    wait_table_loaded(page)

    # 100개씩 보기로 변경 시도
    try:
        dismiss_popups(page)
        page_size_selector = 'select[aria-label="rows per page"]'
        if page.locator(page_size_selector).count() > 0:
            current_size = page.locator(page_size_selector).input_value()
            if current_size != '100':
                page.select_option(page_size_selector, '100')
                page.wait_for_timeout(1000)
                wait_table_loaded(page)
    except Exception:
        dismiss_popups(page)

    results = []
    page_num = 1

    while True:
        dismiss_popups(page)
        wait_table_loaded(page)

        # 테이블 헤더 컬럼 인덱스 동적 탐색
        col_indices = page.evaluate('''() => {
            const headers = Array.from(document.querySelectorAll('.rt-thead.-header .rt-th')).map(th => th.innerText.trim());
            const nameIdx = headers.findIndex(h => h.includes('상품명'));
            const costIdx = headers.findIndex(h => h.includes('집행 광고비'));
            const roasIdx = headers.findIndex(h => h.includes('광고수익률'));
            const salesIdx = headers.findIndex(h => h.includes('광고 전환 매출'));
            return { nameIdx, costIdx, roasIdx, salesIdx, headers };
        }''')

        name_idx = col_indices['nameIdx']
        cost_idx = col_indices['costIdx']
        roas_idx = col_indices['roasIdx']
        sales_idx = col_indices['salesIdx']

        if name_idx == -1 or cost_idx == -1 or roas_idx == -1:
            print(f"    [경고] 테이블 헤더 컬럼을 식별할 수 없습니다. (헤더: {col_indices['headers']})")
            break

        # 현재 페이지의 모든 행 데이터 추출
        raw_rows = page.evaluate('''({ nameIdx, costIdx, roasIdx, salesIdx }) => {
            const groups = document.querySelectorAll('.rt-tbody .rt-tr-group');
            const data = [];
            groups.forEach(g => {
                const cells = g.querySelectorAll('.rt-td');
                if (!cells || cells.length <= Math.max(nameIdx, costIdx, roasIdx)) return;
                
                const nameCell = cells[nameIdx];
                const rawNameText = nameCell.innerText.trim();
                const link = nameCell.querySelector('a');
                const href = link ? link.getAttribute('href') : '';
                
                const costText = cells[costIdx].innerText.trim();
                const roasText = cells[roasIdx].innerText.trim();
                const salesText = (salesIdx !== -1 && cells[salesIdx]) ? cells[salesIdx].innerText.trim() : '0원';
                
                data.push({
                    rawNameText,
                    href,
                    costText,
                    roasText,
                    salesText
                });
            });
            return data;
        }''', {'nameIdx': name_idx, 'costIdx': cost_idx, 'roasIdx': roas_idx, 'salesIdx': sales_idx})

        for row in raw_rows:
            # 옵션ID 파싱 (href vendorItemId 또는 텍스트 ID)
            m_id = re.search(r'vendorItemId=(\d+)', row['href']) or re.search(r'ID\s*:\s*(\d+)', row['rawNameText'])
            option_id = m_id.group(1) if m_id else "알수없음"

            cost = parse_int(row['costText'])
            roas = parse_float(row['roasText'])
            sales = parse_int(row['salesText'])
            product_name = row['rawNameText'].split('\n')[0].strip()

            # 조건 검사: 광고비 >= min_cost AND ROAS < max_roas
            if cost >= min_cost and roas < max_roas:
                results.append({
                    "campaign": campaign_name,
                    "option_id": option_id,
                    "product_name": product_name,
                    "ad_cost": cost,
                    "roas": roas,
                    "sales": sales
                })

        # 페이징 상태 확인 및 다음 페이지 이동
        wait_table_loaded(page)
        pagination = page.evaluate('''() => {
            const totalEl = document.querySelector('.-totalPages');
            const totalPages = totalEl ? parseInt(totalEl.innerText.trim(), 10) : 1;
            const inputEl = document.querySelector('.-pageJump input');
            const curPage = inputEl ? parseInt(inputEl.value.trim(), 10) : 1;
            const nextBtn = document.querySelector('div.-next button.-btn');
            const nextContainer = document.querySelector('div.-next');
            const isDisabled = !nextBtn || nextBtn.disabled || nextBtn.hasAttribute('disabled') || 
                               (nextContainer && nextContainer.classList.contains('-disabled'));
            return {
                curPage: isNaN(curPage) ? 1 : curPage,
                totalPages: isNaN(totalPages) ? 1 : totalPages,
                canNext: !isDisabled
            };
        }''')

        # 더 이상 다음 페이지가 없거나 현재 페이지가 총 페이지 수 이상인 경우 종료
        if not pagination.get('canNext') or pagination.get('curPage', 1) >= pagination.get('totalPages', 1):
            break

        # 다음 버튼 클릭 (JS 기반으로 오버레이 인터셉트 없이 안전하게 클릭)
        try:
            clicked = page.evaluate('''() => {
                const btn = document.querySelector('div.-next button.-btn');
                if (btn && !btn.disabled) {
                    btn.click();
                    return true;
                }
                return false;
            }''')
            if not clicked:
                break
            page_num += 1
            page.wait_for_timeout(1000)
            wait_table_loaded(page)
        except Exception as e:
            break

    return results


def discover_all_campaigns(page):
    """
    대시보드 캠페인 테이블의 모든 페이지를 순회하여
    ON 상태인 캠페인과 OFF 상태인 캠페인을 식별 및 분류
    """
    print("\n[캠페인 목록 탐색] 전체 페이지를 순회하며 캠페인 및 ON/OFF 상태를 파악합니다...")
    dismiss_popups(page)
    wait_table_loaded(page)

    # 1. 한 페이지당 최대(20개) 보기 설정
    try:
        dismiss_popups(page)
        page_size_selector = 'select[aria-label="rows per page"]'
        if page.locator(page_size_selector).count() > 0:
            current_size = page.locator(page_size_selector).input_value()
            if current_size != '20':
                page.select_option(page_size_selector, '20')
                page.wait_for_timeout(1000)
                wait_table_loaded(page)
    except Exception:
        dismiss_popups(page)

    all_campaigns = []
    page_num = 1

    while True:
        dismiss_popups(page)
        wait_table_loaded(page)

        # 현재 페이지의 캠페인 행 정보 추출
        camps_on_page = page.evaluate('''() => {
            const rows = document.querySelectorAll('[data-test-id="pa-cam-dd-row"]');
            const list = [];
            rows.forEach(r => {
                const nameEl = r.querySelector('[data-test-id="pa-cam-dd-name-value"]');
                const switchBtn = r.querySelector('[data-test-id="pa-cam-dd-onoff-cell"] button[role="switch"]');
                const statusEl = r.querySelector('[data-test-id="pa-cam-dd-status-cell"]');

                if (!nameEl) return;
                const name = nameEl.innerText.trim();
                if (!name) return;

                const isOn = switchBtn ? (
                    switchBtn.getAttribute('aria-checked') === 'true' ||
                    switchBtn.classList.contains('ant-switch-checked') ||
                    switchBtn.innerText.includes('ON')
                ) : false;

                const status = statusEl ? statusEl.innerText.trim().replace(/\\s+/g, ' ') : '';
                list.push({ name, isOn, status });
            });
            return list;
        }''')

        for c in camps_on_page:
            c['table_page'] = page_num
            all_campaigns.append(c)

        # 페이징 상태 확인
        pagination = page.evaluate('''() => {
            const totalEl = document.querySelector('.-totalPages');
            const totalPages = totalEl ? parseInt(totalEl.innerText.trim(), 10) : 1;
            const inputEl = document.querySelector('.-pageJump input');
            const curPage = inputEl ? parseInt(inputEl.value.trim(), 10) : 1;
            const nextBtn = document.querySelector('div.-next button.-btn');
            const nextContainer = document.querySelector('div.-next');
            const isDisabled = !nextBtn || nextBtn.disabled || nextBtn.hasAttribute('disabled') ||
                               (nextContainer && nextContainer.classList.contains('-disabled'));
            return {
                curPage: isNaN(curPage) ? 1 : curPage,
                totalPages: isNaN(totalPages) ? 1 : totalPages,
                canNext: !isDisabled
            };
        }''')

        if not pagination.get('canNext') or pagination.get('curPage', 1) >= pagination.get('totalPages', 1):
            break

        # 다음 페이지로 이동
        dismiss_popups(page)
        clicked = page.evaluate('''() => {
            const btn = document.querySelector('div.-next button.-btn');
            if (btn && !btn.disabled) {
                btn.click();
                return true;
            }
            return false;
        }''')
        if not clicked:
            break
        page_num += 1
        page.wait_for_timeout(1000)
        wait_table_loaded(page)

    return all_campaigns


def open_campaign_detail(page, camp_name: str) -> bool:
    """
    대시보드 메인에서 특정 캠페인을 찾아 클릭하여 상세 페이지로 이동
    페이지네이션이 필요한 경우 페이지를 넘겨가며 탐색
    """
    dismiss_popups(page)

    # 1. 대시보드 메인으로 복귀 확인
    if not page.url.endswith("/marketing/dashboard/sales"):
        page.goto(DASHBOARD_URL)
        dismiss_popups(page)
        try:
            page.wait_for_selector('[data-test-id="pa-cam-dd-row"]', timeout=15000)
        except Exception:
            dismiss_popups(page)
            page.wait_for_selector('[data-test-id="pa-cam-dd-row"]', timeout=15000)
        page.wait_for_timeout(1000)
        dismiss_popups(page)

    # 2. 20개씩 보기 유지
    try:
        page_size_selector = 'select[aria-label="rows per page"]'
        if page.locator(page_size_selector).count() > 0:
            current_size = page.locator(page_size_selector).input_value()
            if current_size != '20':
                dismiss_popups(page)
                page.select_option(page_size_selector, '20')
                page.wait_for_timeout(1000)
                wait_table_loaded(page)
    except Exception:
        dismiss_popups(page)

    def try_click_on_current_page():
        dismiss_popups(page)
        return page.evaluate('''(targetName) => {
            const rows = document.querySelectorAll('[data-test-id="pa-cam-dd-row"]');
            for (const r of rows) {
                const nameEl = r.querySelector('[data-test-id="pa-cam-dd-name-value"]');
                if (nameEl && nameEl.innerText.trim() === targetName) {
                    const link = nameEl.closest('a') || nameEl;
                    link.click();
                    return true;
                }
            }
            return false;
        }''', camp_name)

    # 현재 페이지에서 먼저 클릭 시도
    if try_click_on_current_page():
        dismiss_popups(page)
        return True

    # 현재 페이지에 없다면 1페이지로 돌아가서 처음부터 순차 탐색
    pinfo = page.evaluate('''() => {
        const totalEl = document.querySelector('.-totalPages');
        const inputEl = document.querySelector('.-pageJump input');
        return {
            curPage: inputEl ? parseInt(inputEl.value.trim(), 10) : 1,
            totalPages: totalEl ? parseInt(totalEl.innerText.trim(), 10) : 1
        };
    }''')
    cur_p = pinfo.get('curPage', 1)
    total_p = pinfo.get('totalPages', 1)

    while cur_p > 1:
        dismiss_popups(page)
        prev_clicked = page.evaluate('''() => {
            const btn = document.querySelector('div.-previous button.-btn');
            if (btn && !btn.disabled) { btn.click(); return true; }
            return false;
        }''')
        if not prev_clicked:
            break
        page.wait_for_timeout(800)
        wait_table_loaded(page)
        cur_p -= 1

    for _ in range(total_p):
        if try_click_on_current_page():
            dismiss_popups(page)
            return True
        dismiss_popups(page)
        next_clicked = page.evaluate('''() => {
            const btn = document.querySelector('div.-next button.-btn');
            if (btn && !btn.disabled) { btn.click(); return true; }
            return false;
        }''')
        if not next_clicked:
            break
        page.wait_for_timeout(800)
        wait_table_loaded(page)

    return False


def run_crawler(args):
    os.makedirs(args.output_dir, exist_ok=True)
    os.makedirs(args.profile_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_file = os.path.join(args.output_dir, f"coupang_ads_{timestamp}.xlsx")
    json_file = os.path.join(args.output_dir, f"coupang_ads_{timestamp}.json")
    txt_file = os.path.join(args.output_dir, f"option_ids_{timestamp}.txt")

    print("\n" + "="*65)
    print(" 쿠팡 광고(AI 스마트광고) 상품 옵션ID 크롤러")
    print(f" - 집행 광고비 기준 : {args.min_cost:,}원 이상")
    print(f" - 광고 수익률 기준 : {args.max_roas:.1f}% 미만")
    print(f" - 결과 저장 폴더   : {os.path.abspath(args.output_dir)}")
    print("="*65 + "\n")

    with sync_playwright() as p:
        browser_managed_by_us = False
        context = None

        user_agent = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/130.0.0.0 Safari/537.36"
        )
        browser_args = [
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-setuid-sandbox",
        ]
        if not args.headless:
            browser_args.extend(["--remote-debugging-port=9222", "--start-maximized"])

        # 1. 기존 실행 중인 CDP 세션이 있는지 확인 (일반 GUI 모드일 때만)
        if not args.headless:
            try:
                browser = p.chromium.connect_over_cdp("http://localhost:9222")
                context = browser.contexts[0]
                page = context.pages[0] if context.pages else context.new_page()
                print("기존에 열려있는 브라우저(포트 9222)에 성공적으로 연결되었습니다.")
            except Exception:
                pass

        # 2. 없거나 헤드리스 모드인 경우 프로필 기반으로 브라우저 직접 실행
        if context is None:
            mode_str = "헤드리스 백그라운드 모드" if args.headless else "일반 GUI 모드"
            print(f"브라우저를 실행합니다 ({mode_str}, 저장된 세션 프로필 사용)...")
            launch_kwargs = {
                "user_data_dir": os.path.abspath(args.profile_dir),
                "headless": args.headless,
                "args": browser_args,
                "user_agent": user_agent,
            }
            if args.headless:
                launch_kwargs["viewport"] = {"width": 1920, "height": 1080}
            else:
                launch_kwargs["no_viewport"] = True

            context = p.chromium.launch_persistent_context(**launch_kwargs)
            context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")
            page = context.pages[0] if context.pages else context.new_page()
            browser_managed_by_us = True

        # 대시보드 접근 및 로그인 확인
        load_session(context)
        page = wait_for_login(page, context, is_headless=args.headless)

        # 팝업 자동 처리기 및 이벤트 리스너 등록
        setup_page_popup_handlers(page)
        dismiss_popups(page)
        close_extra_popup_pages(context, main_page=page)
        save_session(context)

        # 크롤링 중 발생하는 새 팝업 창 자동 닫기 리스너 등록
        def on_new_crawler_page(new_pg):
            try:
                new_pg.wait_for_load_state("domcontentloaded", timeout=3000)
            except Exception:
                pass
            setup_page_popup_handlers(new_pg)
            dismiss_popups(new_pg)
            close_extra_popup_pages(context, main_page=page)

        try:
            context.on("page", on_new_crawler_page)
        except Exception:
            pass

        # 대시보드 로딩 대기 (팝업 차단 대비 재시도)
        dismiss_popups(page)
        try:
            page.wait_for_selector('[data-test-id="pa-cam-dd-row"]', timeout=15000)
        except Exception:
            dismiss_popups(page)
            page.wait_for_selector('[data-test-id="pa-cam-dd-row"]', timeout=15000)

        # 1. 모든 페이지를 순회하며 캠페인 목록 및 ON/OFF 상태 수집
        all_campaigns = discover_all_campaigns(page)
        on_campaigns = [c for c in all_campaigns if c['isOn']]
        off_campaigns = [c for c in all_campaigns if not c['isOn']]

        print("\n" + "=" * 65)
        print(f" [캠페인 목록 분석 완료: 총 {len(all_campaigns)}개 발견]")
        print(f" ■ 🟢 ON  상태 캠페인 : {len(on_campaigns)}개 (모두 분석 대상)")
        for idx, c in enumerate(on_campaigns):
            print(f"    {idx+1}. {c['name']} (상태: {c['status']})")
        print(f" ■ ⚪ OFF 상태 캠페인 : {len(off_campaigns)}개 (제외됨)")
        for idx, c in enumerate(off_campaigns):
            print(f"    - {c['name']}")
        print("=" * 65)

        if not on_campaigns:
            print("\n[안내] 현재 ON 상태인 캠페인이 없습니다. 크롤링을 종료합니다.\n")
            if browser_managed_by_us:
                context.close()
            return

        all_results_by_campaign = {}
        all_option_ids = []
        total_matched = 0

        # 2. ON 상태인 모든 캠페인 순회 및 분석
        for idx, camp_info in enumerate(on_campaigns):
            camp_name = camp_info['name']
            print(f"\n[{idx+1}/{len(on_campaigns)}] 🟢 ON 캠페인 분석 진행 중: '{camp_name}'")
            try:
                dismiss_popups(page)
                close_extra_popup_pages(context, main_page=page)
                success = open_campaign_detail(page, camp_name)
                if not success:
                    # 팝업 가로막힘 대비 팝업 정리 후 1회 재시도
                    dismiss_popups(page)
                    success = open_campaign_detail(page, camp_name)
                if not success:
                    print(f"    [오류] 캠페인 '{camp_name}' 링크를 찾거나 클릭할 수 없습니다.")
                    all_results_by_campaign[camp_name] = []
                    continue

                try:
                    dismiss_popups(page)
                    page.wait_for_selector('.rt-tbody', timeout=15000)
                    page.wait_for_timeout(1000)
                except Exception:
                    dismiss_popups(page)
                    try:
                        page.wait_for_selector('.rt-tbody', timeout=5000)
                    except Exception:
                        print("    [안내] 상품 목록이 없거나 로딩되지 않았습니다.")
                        all_results_by_campaign[camp_name] = []
                        continue

                # 상품 추출
                matched = extract_products_from_campaign(page, camp_name, args.min_cost, args.max_roas)
                all_results_by_campaign[camp_name] = matched
                total_matched += len(matched)

                print(f"    => '{camp_name}' 분석 완료: 조건 만족 상품 {len(matched)}개")
                for item in matched:
                    all_option_ids.append(item['option_id'])
                    print(f"       • [ID: {item['option_id']}] {item['ad_cost']:,}원 | ROAS: {item['roas']:.2f}% | {item['product_name'][:35]}")
            except Exception as camp_err:
                print(f"    [오류] 캠페인 '{camp_name}' 처리 중 오류 발생 (다음 캠페인으로 계속 진행): {camp_err}")
                all_results_by_campaign[camp_name] = []


        # 파일 저장
        # 1. Excel (XLSX) - 캠페인별 요약, 전체 목록 및 개별 캠페인 시트
        export_to_excel(xlsx_file, all_results_by_campaign, args.min_cost, args.max_roas)

        # 2. JSON
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(all_results_by_campaign, f, ensure_ascii=False, indent=2)

        # 3. 옵션 ID 목록 TXT
        unique_option_ids = list(dict.fromkeys(all_option_ids))
        with open(txt_file, 'w', encoding='utf-8') as f:
            for opt_id in unique_option_ids:
                f.write(f"{opt_id}\n")

        flattened = [it for items in all_results_by_campaign.values() for it in items]

        # 결과 요약 터미널 출력
        print("\n" + "="*65)
        print("  크롤링 완료 요약")
        print("="*65)
        for cname, items in all_results_by_campaign.items():
            print(f"■ 캠페인: [{cname}] (총 {len(items)}개)")
            if not items:
                print("   (조건 만족 상품 없음)")
            else:
                for item in items:
                    print(f"   • 옵션ID: {item['option_id']} | 광고비: {item['ad_cost']:,}원 | ROAS: {item['roas']:.2f}% | 매출: {item['sales']:,}원")
                    print(f"     상품명: {item['product_name']}")
            print()

        print("-" * 65)
        print(f"총 수집된 조건 만족 상품 수: {len(flattened)}개 (고유 옵션ID: {len(unique_option_ids)}개)")
        print("\n[추출된 고유 옵션 ID 전체 목록 (쉼표 구분)]:")
        print(", ".join(unique_option_ids) if unique_option_ids else "(없음)")

        print("\n[생성된 파일 목록]:")
        print(f"1. Excel (XLSX) : {xlsx_file}")
        print(f"2. JSON         : {json_file}")
        print(f"3. TXT          : {txt_file} (옵션 ID만 줄바꿈 저장)")
        print("="*65 + "\n")

        if browser_managed_by_us:
            # 브라우저 직접 실행 시 세션 저장을 위해 정상 닫기
            context.close()

def main():
    parser = argparse.ArgumentParser(description="쿠팡 광고(AI 스마트광고) 옵션ID 크롤러")
    parser.add_argument("--min-cost", type=int, default=5000, help="최소 집행 광고비 (원, 기본값: 5000)")
    parser.add_argument("--max-roas", type=float, default=600.0, help="최대 광고수익률 ROAS (%%, 기본값: 600.0)")
    parser.add_argument("--headless", action="store_true", help="헤드리스 모드로 실행 (화면 미표시)")
    parser.add_argument("--output-dir", type=str, default="./output", help="결과 파일 저장 디렉토리 (기본값: ./output)")
    parser.add_argument("--profile-dir", type=str, default="./browser_profile", help="브라우저 세션 프로필 경로")

    args = parser.parse_args()
    run_crawler(args)

if __name__ == "__main__":
    main()
