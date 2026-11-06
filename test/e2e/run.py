"""Browser test of the real webview (headless Chromium via Python Playwright).

Starts test/e2e/serve.js with a sample deck and checks: slides are drawn, thumbnails fill in, notes show,
search finds slides by text and notes, hostile text stays plain text, the SmartArt notice shows,
the size limit asks first, a 120-slide deck opens, and there are no CSP violations.
With --shots DIR it also saves the README screenshots.

Run: python3 test/e2e/run.py [--shots images] [--chromium PATH]
"""
import argparse
import glob
import os
import pathlib
import subprocess
import sys
import time

from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parents[2]
FIX = ROOT / "test" / "fixtures"
results = []
PROCS = []


def wait_js(page, expr, timeout=10000):
    """Poll with evaluate (CDP, not subject to the page CSP). wait_for_function needs unsafe-eval, which the viewer CSP blocks."""
    end = time.time() + timeout / 1000
    while time.time() < end:
        if page.evaluate(expr):
            return True
        time.sleep(0.05)
    raise TimeoutError("timed out waiting for: " + expr)


def check(name, ok, detail=""):
    print(("PASS " if ok else "FAIL ") + name + (f"  ({detail})" if detail else ""))
    results.append(ok)


def find_chromium(arg):
    if arg:
        return arg
    hits = sorted(glob.glob(os.path.expanduser("~/Library/Caches/ms-playwright/chromium_headless_shell-*/*/chrome-headless-shell")))
    return hits[-1] if hits else None


class Page:
    def __init__(self, browser, deck, theme="dark", max_mb="150", size=(1280, 800)):
        self.proc = subprocess.Popen(["node", str(ROOT / "test/e2e/serve.js"), str(deck), theme, max_mb], stdout=subprocess.PIPE, text=True)
        PROCS.append(self.proc)
        port = int(self.proc.stdout.readline().strip())
        self.page = browser.new_page(viewport={"width": size[0], "height": size[1]}, device_scale_factor=1)
        self.errors = []
        self.page.on("console", lambda m: self.errors.append(m.text) if m.type == "error" else None)
        self.page.on("pageerror", lambda e: self.errors.append(str(e)))
        self.page.goto(f"http://127.0.0.1:{port}/harness.html")

    def close(self):
        self.page.close()
        self.proc.terminate()
        self.proc.wait()

    def sent(self, kind):
        return self.page.evaluate(f"window.__sent.filter(m => m.type === {kind!r})")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--shots")
    ap.add_argument("--chromium")
    args = ap.parse_args()
    with sync_playwright() as p:
        exe = find_chromium(args.chromium)
        browser = p.chromium.launch(headless=True, **({"executable_path": exe} if exe else {}))
        try:
            run_all(browser, args.shots)
        finally:
            browser.close()
            for proc in PROCS:
                if proc.poll() is None:
                    proc.kill()
    failed = results.count(False)
    print(f"\n{len(results) - failed} passed, {failed} failed")
    sys.exit(1 if failed else 0)


def run_all(browser, shots):
    # 1. Main deck, dark theme
    t = Page(browser, FIX / "quarterly-review.pptx", "dark")
    pg = t.page
    pg.wait_for_selector("#outline .outlineItem", state="attached", timeout=10000)
    pg.wait_for_selector("#stage .slideFrame", timeout=20000)
    rendered = t.sent("rendered")
    check("6 slides drawn by pptx-preview", rendered and rendered[0]["slides"] == 6, str(rendered))
    check("slide 1 title is visible on the stage", "Q3 Platform Review" in pg.inner_text("#stage"))
    check("notes for slide 1 shown", "Keep the intro under two minutes" in pg.inner_text("#notes"))
    wait_js(pg, "() => document.querySelectorAll('#thumbs .thumb[data-filled]').length >= 6", 10000)
    check("thumbnails filled", pg.eval_on_selector_all("#thumbs .thumb[data-filled]", "e => e.length") == 6)
    pg.keyboard.press("ArrowRight")
    pg.wait_for_timeout(150)
    check("Right arrow goes to slide 2", pg.input_value("#slideInput") == "2")
    check("slide 2 bullets on stage", "p95 latency fell" in pg.inner_text("#stage"))
    pg.click('#thumbs .thumb[data-index="5"]')
    pg.wait_for_timeout(150)
    notice = pg.inner_text("#notice") if pg.is_visible("#notice") else ""
    check("chart slide shows a chart notice", "chart" in notice, notice[:120])
    pg.fill("#searchInput", "latency")
    wait_js(pg, "() => document.getElementById('searchCount').textContent.includes('slides')", 5000)
    hits = pg.eval_on_selector_all("#outline .outlineItem .num", "e => e.map(x => x.textContent)")
    check("search 'latency' finds slides 1, 2, 3, 6", hits == ["1", "2", "3", "6"], str(hits))
    marks = pg.eval_on_selector_all("#outline mark", "e => e.length")
    check("matches are highlighted in the outline", marks >= 4, str(marks))
    pg.press("#searchInput", "Enter")
    pg.press("#searchInput", "Enter")
    pg.wait_for_timeout(150)
    check("Enter jumps to the 2nd matching slide", pg.input_value("#slideInput") == "2", pg.input_value("#slideInput"))
    pg.fill("#searchInput", "owners")
    wait_js(pg, "() => document.querySelectorAll('#outline .outlineItem').length === 1", 5000)
    snip = pg.inner_text("#outline")
    check("search finds text that is only in the notes", "Notes:" in snip and "6" in snip, snip[:120])
    pg.click("#mdBtn")
    pg.wait_for_timeout(100)
    md = t.sent("copySlideMarkdown")
    check("Markdown button sends copySlideMarkdown for the current slide", md and md[-1] == {"type": "copySlideMarkdown", "slide": int(pg.input_value("#slideInput"))}, str(md))
    pg.click("#notesBtn")
    check("Notes button hides the notes", not pg.is_visible("#notes"))
    pg.click("#notesBtn")
    csp = [e for e in t.errors if "Content Security Policy" in e or "Refused" in e]
    check("no CSP violations or page errors", not t.errors, "; ".join(t.errors)[:300])

    if shots:
        out = ROOT / shots
        out.mkdir(exist_ok=True)
        pg.fill("#searchInput", "")
        pg.press("#searchInput", "Escape")
        pg.click('#thumbs .thumb[data-index="4"]')
        pg.wait_for_timeout(300)
        pg.screenshot(path=str(out / "screenshot-slides.png"))
        pg.fill("#searchInput", "latency")
        wait_js(pg, "() => document.getElementById('searchCount').textContent.includes('slides')", 5000)
        pg.press("#searchInput", "Enter")
        pg.press("#searchInput", "Enter")
        pg.press("#searchInput", "Enter")
        pg.wait_for_timeout(300)
        pg.screenshot(path=str(out / "screenshot-search.png"))
    t.close()

    # 2. Light theme, table slide (screenshot only, plus a draw check)
    t = Page(browser, FIX / "quarterly-review.pptx", "light")
    t.page.wait_for_selector("#stage .slideFrame", timeout=20000)
    t.page.click('#thumbs .thumb[data-index="3"]')
    t.page.wait_for_timeout(300)
    check("[light] table slide drawn", "ap-south" in t.page.inner_text("#stage"))
    if shots:
        t.page.screenshot(path=str(ROOT / shots / "screenshot-light.png"))
    t.close()

    # 3. Hostile text is shown as text and never runs
    t = Page(browser, FIX / "hostile.pptx", "dark")
    pg = t.page
    pg.wait_for_selector("#stage .slideFrame", timeout=20000)
    pg.wait_for_timeout(500)
    check("hostile: no script ran (title unchanged)", pg.title() == "PPTX Slide Viewer", pg.title())
    check("hostile: tag text shown literally", "<img src=x" in pg.inner_text("#stage") and "<b>not bold</b>" in pg.inner_text("#notes"))
    check("hostile: no <img>/<script> elements created from text", pg.eval_on_selector_all("#stage img, #stage script, #notes b, #outline img", "e => e.length") == 0)
    pg.fill("#searchInput", "script")
    pg.wait_for_timeout(300)
    check("hostile: outline search result stays text", pg.eval_on_selector_all("#outline script", "e => e.length") == 0 and "<script>" in pg.inner_text("#outline"))
    t.close()

    # 4. SmartArt + video notice
    t = Page(browser, FIX / "smartart.pptx", "dark")
    t.page.wait_for_selector("#stage .slideFrame, #stage .textSlide", timeout=20000)
    t.page.wait_for_timeout(300)
    n = t.page.inner_text("#notice") if t.page.is_visible("#notice") else ""
    check("SmartArt and video notice shown", "SmartArt" in n and "video" in n and "cannot draw" in n, n[:160])
    t.close()

    # 5. Size limit asks first, then opens
    t = Page(browser, FIX / "quarterly-review.pptx", "dark", max_mb="0.01")
    t.page.wait_for_selector("#notice.warn button", timeout=10000)
    check("over the size limit: shows a message and Open anyway", "limit" in t.page.inner_text("#notice"))
    t.page.click("#notice button")
    t.page.wait_for_selector("#stage .slideFrame", timeout=20000)
    check("Open anyway loads the deck", True)
    t.close()

    # 6. 120 slides: text outline first, drawing time
    t0 = time.time()
    t = Page(browser, FIX / "many-slides.pptx", "dark")
    t.page.wait_for_selector("#outline .outlineItem", state="attached", timeout=10000)
    first_text = time.time() - t0
    t.page.wait_for_selector("#stage .slideFrame", timeout=60000)
    r = t.sent("rendered")
    check("120-slide deck: outline ready, then all slides drawn", r and r[0]["slides"] == 120, f"outline {first_text:.2f}s, draw {r[0]['ms'] if r else '?'} ms")
    t.page.keyboard.press("End")
    t.page.wait_for_timeout(200)
    check("End goes to slide 120", t.page.input_value("#slideInput") == "120")
    t.close()

    # 7. ppsx and potx open the same way
    for f in ("quarterly-review.ppsx", "template.potx"):
        t = Page(browser, FIX / f, "dark")
        t.page.wait_for_selector("#stage .slideFrame", timeout=20000)
        check(f"{f} opens and draws", "Q3 Platform Review" in t.page.inner_text("#stage"))
        t.close()


if __name__ == "__main__":
    main()
