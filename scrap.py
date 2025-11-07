import asyncio
from pathlib import Path
from playwright.async_api import async_playwright
from dotenv import load_dotenv
import os
import csv
import re
from datetime import datetime
from zoneinfo import ZoneInfo  # built-in since Python 3.9
import argparse
import json
import sys

jst = ZoneInfo("Asia/Tokyo")

async def scraping(spreadsheet_id=None, sheet_name=None):
    # --- start time ---
    print("スクレイピングが始まりました...")
    start_time = datetime.now(jst)
    print(f"開始時刻: {start_time.strftime('%Y-%m-%d, %H:%M:%S')}")

    # Load environment variables
    load_dotenv()

    target_url = os.getenv("TARGET_URL")
    asins_file = os.getenv("ASINS_FILE")
    keywords_file = os.getenv("KEYWORDS_FILE")

    # Read ASINs (one column, no header)
    with open(asins_file, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        target_asins = [row[0].strip() for row in reader if row]  # remove empty lines

    # Read Keywords (one column, no header)
    with open(keywords_file, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        keywords = [row[0].strip() for row in reader if row]

    # chromium
    # Use exe directory when frozen, script directory otherwise
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
    browsers_dir = base_dir / ".playwright-browsers"
    
    chromium_dirs = list(browsers_dir.glob("chromium-*"))
    if not chromium_dirs:
        raise Exception("No Chromium browser found in .playwright-browsers")
    chromium_path = chromium_dirs[0] / "chrome-win" / "chrome.exe"

    # start playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            slow_mo=200,
            executable_path=str(chromium_path)
        )
        context = await browser.new_context(
            viewport={"width": 1600, "height": 1000}
        )
        page = await context.new_page()

        result = []  # [{"keyword": "自然検索", "SP": "", "SB": ""}, ...]

        for keyword in keywords:
            organic_products_info = []  # [{"asin": "", "page": ""}, ...]
            sponsored_products_info = []  # [{"asin": "", "page": ""}, ...]
            sb_products_info = [] # [{"asins": ["", ""], "page": number}, ...]
            for page_index in range(1, 3):  # Scrape first 2 pages for each keyword
                
                # Go to target URL page
                page_url = f"{target_url}s?k={keyword}&page={page_index}"
                await page.goto(page_url)

                # Get product elements (role is "listitem" and each product must have data-asin attribute in it)
                product_elements = await page.locator('[role="listitem"][data-asin]').all()

                # Get product elements for SB ads (data-asin attribute is existed in it, but it' s "")
                sb_ad_elements = await page.locator('[data-asin=""]').all()
                
                filtered_sb_ad_elements = []

                # Step 1: Cut from "関連検索キーワード"
                cut_index = None
                for i, element in enumerate(sb_ad_elements):
                    element_content = await element.inner_html()
                    if '<h2 class="a-size-medium-plus a-color-base">関連検索キーワード</h2>' in element_content:
                        cut_index = i
                        break

                if cut_index is not None:
                    sb_ad_elements = sb_ad_elements[:cut_index]

                # Step 2: Remove unwanted <h2> blocks
                unwanted_blocks = [
                    '<h2 class="a-size-medium-plus a-spacing-none a-color-base a-text-bold">結果</h2>',
                    '<h2 id="loom-desktop-inline-slot_featuredasins-heading" class="a-size-medium-plus a-color-base">高評価</h2>',
                    '<h2 class="a-size-medium-plus a-spacing-none a-color-base a-text-bold">その他の結果</h2>',
                    '<h2 id="loom-desktop-bottom-slot_featuredasins-heading" class="a-size-medium-plus a-color-base">開催中のタイムセール</h2>',
                    '<h2 id="loom-desktop-inline-slot_featuredasins-heading" class="a-size-medium-plus a-color-base">今のトレンド</h2>',
                ]

                for element in sb_ad_elements:
                    element_content = await element.inner_html()
                    if any(block in element_content for block in unwanted_blocks):
                        continue
                    filtered_sb_ad_elements.append(element)

                for product_element in product_elements:
                    asin = await product_element.get_attribute("data-asin")

                    element_content = await product_element.inner_html()
                    if "スポンサー" in element_content:
                        sponsored_products_info.append({
                            "asin": asin,
                            "page": page_index
                        })
                    else:
                        organic_products_info.append({
                            "asin": asin,
                            "page": page_index
                        })

                asin_pattern = re.compile(r'B0[A-Z0-9]{8}')
                for element in filtered_sb_ad_elements:
                    element_content = await element.inner_html()
                    asins = asin_pattern.findall(element_content)
                    asins = list(set(asins))  # remove duplicates
                    if asins:
                        sb_products_info.append({
                            "asins": asins,
                            "page": page_index
                        })
                await asyncio.sleep(3)
            
            
            gotRankedForOrganic = False
            gotProductForOrganic = {}  # {"asin": "", "page": number, "rank": number}

            for index, product in enumerate(organic_products_info, start=1):
                ASIN = product["asin"]

                if ASIN not in target_asins:
                    continue
                gotRankedForOrganic = True
                gotProductForOrganic = {
                    "asin": ASIN,
                    "page": product["page"],
                    "rank": index
                }
                break

            gotRankedForSponsored = False
            gotProductForSponsored = {}  # {"asin": "", "page": number, "rank": number}

            for index, product in enumerate(sponsored_products_info, start=1):
                ASIN = product["asin"]
                if ASIN not in target_asins:
                    continue
                gotRankedForSponsored = True
                gotProductForSponsored = {
                    "asin": ASIN,
                    "page": product["page"],
                    "rank": index
                }
                break

            gotRankedForSB = False
            gotProductForSB = {}  # {"asin": "", "page": number, "rank": number}

            for index, product in enumerate(sb_products_info, start=1):
                ASINS = product["asins"]  # list of ASINs in this SB ad
                
                found_asin = next((a for a in ASINS if a in target_asins), None)

                if not found_asin:
                    continue  # no matching target ASIN in this SB block

                gotRankedForSB = True
                gotProductForSB = {
                    "asin": found_asin,     # store the first matching ASIN
                    "page": product["page"],
                    "rank": index
                }
                break
            
            # Format results for output
            organic_result = "-"
            if gotRankedForOrganic:
                if gotProductForOrganic["page"] == 1:
                    organic_result = str(gotProductForOrganic["rank"])
                else:
                    organic_result = "2ページ目"

            sponsored_result = "-"
            if gotRankedForSponsored:
                if gotProductForSponsored["page"] == 1:
                    sponsored_result = str(gotProductForSponsored["rank"])
                else:
                    sponsored_result = "2ページ目"

            sb_result = "-"
            if gotRankedForSB:
                if gotProductForSB["page"] == 1:
                    sb_result = str(gotProductForSB["rank"])
                else:
                    sb_result = "2ページ目"

            print(f"キーワード： {keyword}, 自然検索: {organic_result}, SP: {sponsored_result}, SB: {sb_result}")

            result.append({
                "keyword": keyword,
                "自然検索": organic_result,
                "SP": sponsored_result,
                "SB": sb_result
            })

        await browser.close()
        # --- finish time ---
        print("スクレイピングが完了しました。")
        finish_time = datetime.now(jst)
        print(f"終了時刻: {finish_time.strftime('%Y-%m-%d, %H:%M:%S')}")

        # --- execution time ---
        execution_time = finish_time - start_time
        print(f"実行時間: {execution_time}", flush=True)

        # 結果をJSONとして出力（app.pyが読み取るため）
        if result:
            print(f"RESULT_DATA:{json.dumps(result, ensure_ascii=False)}", flush=True)
        
        return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--spreadsheet-id", help="Google Spreadsheet ID")
    parser.add_argument("--sheet", help="Sheet name")
    args = parser.parse_args()
    
    asyncio.run(scraping(args.spreadsheet_id, args.sheet))