import asyncio
from pathlib import Path
import pandas as pd
from playwright.async_api import async_playwright

# Change this if you want a different default save location
OUTPUT_DIR = Path.home() / "Documents"

async def scrape_ready_table(ws_endpoint="http://localhost:9222", output_file="pvalue_table.xlsx"):
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp(ws_endpoint)
        page = None
        # Grab the first open page with our table
        for ctx in browser.contexts:
            for pg in ctx.pages:
                html = (await pg.content()).lower()
                if "<table" in html or 'role="grid"' in html:
                    page = pg
                    break
            if page:
                break
        if not page:
            raise RuntimeError("No page with a table found. Make sure it's loaded in the debug browser.")

        # Try HTML table first
        if await page.locator("table").count() > 0:
            headers = await page.locator("table thead th").all_inner_texts()
            if not headers:
                headers = await page.locator("table tr th").all_inner_texts()
            rows = []
            trs = page.locator("table tbody tr")
            for i in range(await trs.count()):
                tds = await trs.nth(i).locator("td").all_inner_texts()
                rows.append([t.strip() for t in tds])
        else:
            # ARIA grid
            headers = await page.locator("[role='columnheader']").all_inner_texts()
            headers = [h.strip() for h in headers]
            rows = []
            rs = page.locator("[role='row']")
            for i in range(await rs.count()):
                r = rs.nth(i)
                if await r.locator("[role='columnheader']").count() > 0:
                    continue
                cells = await r.locator("[role='gridcell'],[role='cell']").all_inner_texts()
                rows.append([c.strip() for c in cells])

        # Normalize
        width = max(len(r) for r in rows) if rows else 0
        rows = [r + [""]*(width-len(r)) for r in rows]
        if not headers or len(headers) != width:
            headers = [f"col_{i+1}" for i in range(width)]
        df = pd.DataFrame(rows, columns=headers)

        out_path = OUTPUT_DIR / output_file
        df.to_excel(out_path, index=False)
        print(f"Saved {len(df):,} rows to: {out_path}")

if __name__ == "__main__":
    asyncio.run(scrape_ready_table())
