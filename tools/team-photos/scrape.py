"""Open every team page in targets.json in a headless browser and save the photo cards to raw.json.

    python3 tools/team-photos/scrape.py                 # all teams
    python3 tools/team-photos/scrape.py "Mars Rover"    # one or more teams; the rest of raw.json is kept

Needs Playwright:  pip install playwright && playwright install chromium
A target is a URL, or "URL|Link text" for single-page sites where the team view opens from a nav link.
"""
import json, os, sys, asyncio
from playwright.async_api import async_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
TARGETS = os.path.join(HERE, 'targets.json')
RAW = os.path.join(HERE, 'raw.json')
JS = open(os.path.join(HERE, 'extract.js')).read()
UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'


async def scrape_team(ctx, sem, team, specs, result):
    async with sem:
        page = await ctx.new_page()
        pages = {}
        for spec in specs:
            url, _, click = spec.partition('|')
            try:
                try:
                    await page.goto(url, wait_until='networkidle', timeout=30000)
                except Exception:
                    await page.wait_for_timeout(4000)
                if click:
                    await page.get_by_text(click, exact=True).first.click(timeout=8000)
                    await page.wait_for_timeout(2500)
                for _ in range(30):  # scroll so lazy images load
                    await page.mouse.wheel(0, 1200)
                    await page.wait_for_timeout(120)
                await page.wait_for_timeout(1200)
                pages[spec] = {'url': page.url, 'items': await page.evaluate(JS)}
            except Exception as e:
                pages[spec] = {'error': str(e)[:200]}
        result[team] = pages
        await page.close()
        print(f"{team}: " + ', '.join(f"{len(v['items'])} photos" if 'items' in v else 'error: ' + v['error'][:60]
                                     for v in pages.values()), flush=True)


async def main(only):
    targets = json.load(open(TARGETS))
    result = json.load(open(RAW)) if os.path.exists(RAW) else {}
    async with async_playwright() as pw:
        browser = await pw.chromium.launch()
        ctx = await browser.new_context(viewport={'width': 1400, 'height': 1000}, user_agent=UA)
        sem = asyncio.Semaphore(6)
        await asyncio.gather(*[scrape_team(ctx, sem, t, s, result)
                               for t, s in targets.items() if s and (not only or t in only)])
        await browser.close()
    with open(RAW, 'w') as f:
        json.dump(result, f, indent=1, ensure_ascii=False)
    print('wrote', os.path.relpath(RAW))


if __name__ == '__main__':
    asyncio.run(main(sys.argv[1:]))
