import os
import re
import json
import time
from playwright.sync_api import sync_playwright
from excel_exporter import export_to_excel
from login_manager import (
    dismiss_popups,
    close_extra_popup_pages,
    setup_page_popup_handlers,
)

XLSX_OUTPUT = os.path.abspath("./crawl_results.xlsx")
JSON_OUTPUT = os.path.abspath("./crawl_results.json")

def parse_int(text):
    clean = re.sub(r'[^\d]', '', text)
    return int(clean) if clean else 0

def parse_float(text):
    clean = text.replace(',', '').replace('%', '').strip()
    try:
        return float(clean)
    except:
        return 0.0

def extract_products_from_current_campaign(page, campaign_name):
    """현재 열려있는 캠페인 상세 페이지에서 모든 페이지를 순회하며 조건에 맞는 상품을 추출"""
    # 1. 100개씩 보기로 변경 시도
    try:
        page_size_selector = 'select[aria-label="rows per page"]'
        if page.locator(page_size_selector).count() > 0:
            page.select_option(page_size_selector, '100')
            page.wait_for_timeout(2000)
    except Exception as e:
        print(f"  [안내] 100개씩 보기 변경 생략: {e}")

    results = []
    page_num = 1

    while True:
        # 헤더 인덱스 동적 탐색
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
            print(f"  [경고] 테이블 헤더를 찾을 수 없습니다. (헤더: {col_indices['headers']})")
            break

        # 행 데이터 일괄 추출
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

        # 필터링
        for row in raw_rows:
            # 옵션 ID 파싱
            m_id = re.search(r'vendorItemId=(\d+)', row['href']) or re.search(r'ID\s*:\s*(\d+)', row['rawNameText'])
            option_id = m_id.group(1) if m_id else "알수없음"

            cost = parse_int(row['costText'])
            roas = parse_float(row['roasText'])
            sales = parse_int(row['salesText'])
            product_name = row['rawNameText'].split('\n')[0].strip()

            # 조건: 광고비 >= 5,000원 AND ROAS < 600%
            if cost >= 5000 and roas < 600.0:
                results.append({
                    "campaign": campaign_name,
                    "option_id": option_id,
                    "product_name": product_name,
                    "ad_cost": cost,
                    "roas": roas,
                    "sales": sales
                })

        # 페이징 상태 확인 및 다음 페이지 이동
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

        print(f"  -> {pagination['curPage']}/{pagination['totalPages']} 페이지 완료 (현재 누적 {len(results)}건), 다음 페이지 이동...")

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
            try:
                page.locator('div.-loading.-active').wait_for(state='detached', timeout=5000)
            except Exception:
                pass
        except Exception:
            break

    return results

def discover_all_campaigns(page):
    """대시보드 캠페인 테이블의 모든 페이지를 순회하며 ON/OFF 상태를 식별 및 분류"""
    print("\n[캠페인 목록 탐색] 전체 페이지를 순회하며 캠페인 및 ON/OFF 상태를 파악합니다...")
    
    # 20개씩 보기 설정
    try:
        page_size_selector = 'select[aria-label="rows per page"]'
        if page.locator(page_size_selector).count() > 0:
            page.select_option(page_size_selector, '20')
            page.wait_for_timeout(1000)
    except Exception:
        pass

    all_campaigns = []
    page_num = 1

    while True:
        try:
            page.locator('div.-loading.-active').wait_for(state='detached', timeout=5000)
        except Exception:
            pass

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

    return all_campaigns


def open_campaign_detail(page, camp_name, dashboard_url):
    """대시보드 메인에서 특정 캠페인을 찾아 클릭하여 상세 페이지로 이동"""
    if not page.url.endswith("/marketing/dashboard/sales"):
        page.goto(dashboard_url)
        page.wait_for_selector('[data-test-id="pa-cam-dd-row"]', timeout=15000)
        page.wait_for_timeout(1000)

    try:
        page_size_selector = 'select[aria-label="rows per page"]'
        if page.locator(page_size_selector).count() > 0:
            if page.locator(page_size_selector).input_value() != '20':
                page.select_option(page_size_selector, '20')
                page.wait_for_timeout(1000)
    except Exception:
        pass

    def try_click_on_current_page():
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

    if try_click_on_current_page():
        return True

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
        prev_clicked = page.evaluate('''() => {
            const btn = document.querySelector('div.-previous button.-btn');
            if (btn && !btn.disabled) { btn.click(); return true; }
            return false;
        }''')
        if not prev_clicked:
            break
        page.wait_for_timeout(800)
        cur_p -= 1

    for _ in range(total_p):
        if try_click_on_current_page():
            return True
        next_clicked = page.evaluate('''() => {
            const btn = document.querySelector('div.-next button.-btn');
            if (btn && !btn.disabled) { btn.click(); return true; }
            return false;
        }''')
        if not next_clicked:
            break
        page.wait_for_timeout(800)

    return False


def main():
    print("="*60)
    print("쿠팡 광고 AI스마트광고 일괄 수집 크롤러 시작")
    print("조건: 집행 광고비 >= 5,000원 AND 광고 수익률 < 600%")
    print("대상: ON 상태인 캠페인 전체")
    print("="*60)

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        context = browser.contexts[0]
        page = context.pages[0]

        setup_page_popup_handlers(page)
        dismiss_popups(page)
        close_extra_popup_pages(context, main_page=page)

        # 대시보드 메인으로 이동
        dashboard_url = 'https://advertising.coupang.com/marketing/dashboard/sales'
        print(f"\n[1] 매출 성장 대시보드 접근 중: {dashboard_url}")
        page.goto(dashboard_url)
        dismiss_popups(page)
        try:
            page.wait_for_selector('[data-test-id="pa-cam-dd-row"]', timeout=15000)
        except Exception:
            dismiss_popups(page)
            page.wait_for_selector('[data-test-id="pa-cam-dd-row"]', timeout=15000)
        page.wait_for_timeout(2000)
        dismiss_popups(page)

        # 1. 모든 페이지 순회하며 캠페인 목록 및 ON/OFF 상태 수집
        all_campaigns = discover_all_campaigns(page)
        on_campaigns = [c for c in all_campaigns if c['isOn']]
        off_campaigns = [c for c in all_campaigns if not c['isOn']]

        print("\n" + "=" * 60)
        print(f"[캠페인 탐색 완료: 총 {len(all_campaigns)}개]")
        print(f"■ 🟢 ON 상태 캠페인 : {len(on_campaigns)}개 (모두 분석 대상)")
        for idx, c in enumerate(on_campaigns):
            print(f"   {idx+1}. {c['name']} (상태: {c['status']})")
        print(f"■ ⚪ OFF 상태 캠페인: {len(off_campaigns)}개 (제외됨)")
        for idx, c in enumerate(off_campaigns):
            print(f"   - {c['name']}")
        print("=" * 60)

        if not on_campaigns:
            print("\n[안내] 현재 ON 상태인 캠페인이 없습니다. 크롤링을 종료합니다.")
            return

        all_results_by_campaign = {}
        total_matched_count = 0

        # 2. ON 상태 캠페인 순회 및 상품 분석
        for idx, camp_info in enumerate(on_campaigns):
            camp_name = camp_info['name']
            print(f"\n[{idx+1}/{len(on_campaigns)}] 🟢 ON 캠페인 분석 진행 중: '{camp_name}'")

            # 해당 캠페인 상세 이동
            success = open_campaign_detail(page, camp_name, dashboard_url)
            if not success:
                print(f"  [주의] 캠페인 '{camp_name}' 링크 클릭 실패")
                all_results_by_campaign[camp_name] = []
                continue
            
            # 상품 테이블 로딩 대기
            try:
                page.wait_for_selector('.rt-tbody', timeout=15000)
                page.wait_for_timeout(1500)
            except Exception as e:
                print(f"  [주의] 상품 목록 로딩 실패 또는 상품 없음: {e}")
                all_results_by_campaign[camp_name] = []
                continue

            # 상품 추출
            matched_items = extract_products_from_current_campaign(page, camp_name)
            all_results_by_campaign[camp_name] = matched_items
            total_matched_count += len(matched_items)

            print(f"  => '{camp_name}' 캠페인 수집 완료: 조건 만족 상품 {len(matched_items)}개")

        # 결과 저장
        # 1. Excel (XLSX) 저장
        export_to_excel(XLSX_OUTPUT, all_results_by_campaign, min_cost=5000, max_roas=600.0)

        # 2. JSON 저장
        with open(JSON_OUTPUT, 'w', encoding='utf-8') as f:
            json.dump(all_results_by_campaign, f, ensure_ascii=False, indent=2)

        print("\n" + "="*60)
        print("수집 완료!")
        print(f"- 전체 조건 부합 상품 수: {total_matched_count}개")
        print(f"- Excel 저장 경로: {XLSX_OUTPUT}")
        print(f"- JSON 저장 경로: {JSON_OUTPUT}")
        print("="*60 + "\n")

if __name__ == "__main__":
    main()
