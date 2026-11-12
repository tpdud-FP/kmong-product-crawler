"""
쇼핑몰 상품 정보 자동 수집기 (포트폴리오 샘플)

대상: books.toscrape.com (크롤링 연습용으로 공개된 예시 쇼핑몰)
결과: 상품 목록 + 카테고리별 요약 + 차트가 들어간 엑셀 파일

실행 방법
  pip3 install requests beautifulsoup4 openpyxl
  python3 book_price_crawler.py
"""

import re
import time
from datetime import datetime
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_URL = "https://books.toscrape.com/"
GBP_TO_KRW = 1750          # 원화 환산용 환율 (예시 값)
REQUEST_DELAY = 0.3        # 서버 부담을 줄이기 위한 요청 간격(초)
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

session = requests.Session()
session.headers["User-Agent"] = "Mozilla/5.0 (portfolio sample crawler)"


def fetch(url, retries=3):
    """페이지를 가져와서 BeautifulSoup 객체로 반환. 실패하면 재시도."""
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(url, timeout=10)
            resp.raise_for_status()
            resp.encoding = "utf-8"
            time.sleep(REQUEST_DELAY)
            return BeautifulSoup(resp.text, "html.parser")
        except requests.RequestException as e:
            print(f"  ! 요청 실패 ({attempt}/{retries}): {url} - {e}")
            time.sleep(1.5 * attempt)
    raise RuntimeError(f"페이지를 가져오지 못했습니다: {url}")


def get_categories():
    """메인 페이지 사이드바에서 카테고리 이름과 주소를 가져온다."""
    soup = fetch(BASE_URL)
    links = soup.select("div.side_categories ul li ul li a")
    return [(a.get_text(strip=True), urljoin(BASE_URL, a["href"])) for a in links]


def parse_products(soup, page_url, category):
    """목록 페이지 한 장에서 상품 정보를 뽑는다."""
    items = []
    for pod in soup.select("article.product_pod"):
        link = pod.select_one("h3 a")
        price_text = pod.select_one("p.price_color").get_text()
        price_match = re.search(r"[\d.]+", price_text)
        rating_tag = pod.select_one("p.star-rating")
        rating_word = next((c for c in rating_tag.get("class", []) if c in RATING_MAP), None)
        stock_text = pod.select_one("p.availability").get_text(strip=True)

        price_gbp = float(price_match.group()) if price_match else None
        items.append({
            "category": category,
            "title": link.get("title", link.get_text(strip=True)),
            "price_gbp": price_gbp,
            "price_krw": round(price_gbp * GBP_TO_KRW) if price_gbp is not None else None,
            "rating": RATING_MAP.get(rating_word),
            "in_stock": "재고 있음" if "In stock" in stock_text else "품절",
            "url": urljoin(page_url, link["href"]),
        })
    return items


def crawl_category(name, url):
    """카테고리 하나를 마지막 페이지까지 수집한다."""
    products = []
    while url:
        soup = fetch(url)
        products.extend(parse_products(soup, url, name))
        next_link = soup.select_one("li.next a")
        url = urljoin(url, next_link["href"]) if next_link else None
    return products


# ---------------- 엑셀 저장 ----------------

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, row=1):
    for cell in ws[row]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def write_product_sheet(wb, products):
    ws = wb.active
    ws.title = "상품 목록"
    headers = ["번호", "카테고리", "상품명", "가격(£)", "가격(원)", "평점", "재고", "상품 링크"]
    ws.append(headers)
    for i, p in enumerate(products, start=1):
        ws.append([i, p["category"], p["title"], p["price_gbp"], p["price_krw"],
                   p["rating"], p["in_stock"], p["url"]])
    style_header(ws)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = BORDER
        row[0].alignment = Alignment(horizontal="center")
        row[3].number_format = '"£"#,##0.00'
        row[4].number_format = "#,##0\"원\""
        row[5].number_format = "0\"점\""
        row[5].alignment = Alignment(horizontal="center")
        row[6].alignment = Alignment(horizontal="center")

    widths = [9, 18, 55, 13, 14, 9, 11, 60]   # 필터 버튼 자리까지 고려한 너비
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_summary_sheet(wb, products, elapsed_sec):
    ws = wb.create_sheet("카테고리 요약")

    # 카테고리별 집계
    stats = {}
    for p in products:
        s = stats.setdefault(p["category"], {"count": 0, "price": 0.0, "rating": 0})
        s["count"] += 1
        s["price"] += p["price_krw"] or 0
        s["rating"] += p["rating"] or 0
    rows = sorted(stats.items(), key=lambda kv: kv[1]["count"], reverse=True)

    ws.append(["수집 정보"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(["수집 일시", datetime.now().strftime("%Y-%m-%d %H:%M")])
    ws.append(["총 상품 수", f"{len(products):,}개"])
    ws.append(["카테고리 수", f"{len(stats)}개"])
    ws.append(["소요 시간", f"{elapsed_sec:.0f}초"])
    ws.append([])

    ws.append(["카테고리", "상품 수", "평균 가격(원)", "평균 평점"])
    header_row = ws.max_row
    style_header(ws, header_row)
    for name, s in rows:
        ws.append([name, s["count"], round(s["price"] / s["count"]), round(s["rating"] / s["count"], 2)])
    last_row = ws.max_row

    for row in ws.iter_rows(min_row=header_row + 1, max_row=last_row):
        for cell in row:
            cell.border = BORDER
        row[2].number_format = "#,##0\"원\""
        row[3].number_format = "0.00\"점\""

    for col, w in zip("ABCD", [22, 10, 15, 11]):
        ws.column_dimensions[col].width = w

    # 상품 수 상위 10개 카테고리 막대 차트
    chart = BarChart()
    chart.type = "bar"
    chart.title = "카테고리별 상품 수 TOP 10"
    chart.y_axis.title = "상품 수"
    chart.legend = None
    top_end = min(header_row + 10, last_row)
    data = Reference(ws, min_col=2, min_row=header_row, max_row=top_end)
    cats = Reference(ws, min_col=1, min_row=header_row + 1, max_row=top_end)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.x_axis.scaling.orientation = "maxMin"   # 1위가 맨 위로
    chart.x_axis.delete = False                   # 카테고리 이름 표시
    chart.y_axis.delete = False                   # 상품 수 눈금 표시
    chart.y_axis.crosses = "max"                  # 눈금을 차트 아래쪽에
    chart.y_axis.majorGridlines = None
    chart.varyColors = False                      # 막대 색 하나로 통일
    chart.title.overlay = False                   # 제목이 막대를 가리지 않게
    series = chart.series[0]
    series.graphicalProperties.solidFill = "1F4E78"
    series.graphicalProperties.line.solidFill = "1F4E78"
    chart.gapWidth = 60
    chart.height, chart.width = 10, 18
    ws.add_chart(chart, "F2")

    wb.move_sheet(ws, offset=-1)   # 요약 시트를 맨 앞으로
    wb.active = 0


def main():
    start = time.time()
    print("카테고리 목록을 가져오는 중...")
    categories = get_categories()
    print(f"카테고리 {len(categories)}개 발견\n")

    products = []
    for i, (name, url) in enumerate(categories, start=1):
        items = crawl_category(name, url)
        products.extend(items)
        print(f"[{i:>2}/{len(categories)}] {name:<22} {len(items):>4}개  (누적 {len(products)}개)")

    elapsed = time.time() - start
    wb = Workbook()
    write_product_sheet(wb, products)
    write_summary_sheet(wb, products, elapsed)

    filename = f"상품수집결과_{datetime.now():%Y%m%d}.xlsx"
    wb.save(filename)
    print(f"\n완료! 총 {len(products)}개 상품 → {filename}  ({elapsed:.0f}초)")


if __name__ == "__main__":
    main()
