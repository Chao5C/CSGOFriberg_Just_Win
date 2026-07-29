"""
分析脚本 v2：点击开始游戏，分析输入框和反馈DOM结构
"""
import asyncio
import json
from playwright.async_api import async_playwright


async def inspect_game():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})

        # 监听XHR/fetch请求
        api_calls = []
        page.on("request", lambda req: api_calls.append(f"[REQ] {req.method} {req.url}") if '/api/' in req.url or '/game' in req.url or '/single' in req.url else None)
        page.on("response", lambda resp: api_calls.append(f"[RESP] {resp.status} {resp.url}") if '/api/' in resp.url or '/game' in resp.url or '/single' in resp.url else None)

        print("[1] 打开单人模式...")
        await page.goto("https://shnlfriberg.online/single", wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(2000)

        # 先关闭公告
        try:
            notice_btn = await page.query_selector('button:has-text("我已知晓")')
            if notice_btn:
                await notice_btn.click()
                await page.wait_for_timeout(1000)
                print("  关闭了公告")
        except:
            pass

        # 点击开始游戏（默认入门版）
        print("[2] 点击开始游戏...")
        start_btn = await page.query_selector('button:has-text("开始游戏")')
        if start_btn:
            await start_btn.click()
            await page.wait_for_timeout(3000)
            print("  已点击开始")
        else:
            print("  [ERROR] 找不到开始按钮")
            await browser.close()
            return

        await page.screenshot(path="data/inspect_game_started.png", full_page=True)
        html = await page.content()
        with open("data/inspect_game_started.html", "w", encoding="utf-8") as f:
            f.write(html)

        # 分析游戏界面
        print("[3] 分析游戏界面...")

        # 输入框
        inputs = await page.query_selector_all('input')
        print(f"\n=== 输入框 ({len(inputs)}) ===")
        for i, inp in enumerate(inputs):
            info = await inp.evaluate("""el => ({
                placeholder: el.placeholder,
                type: el.type,
                className: el.className,
                id: el.id,
                value: el.value,
                autocomplete: el.autocomplete,
                role: el.getAttribute('role'),
                outerHTML: el.outerHTML.substring(0, 500)
            })""")
            print(f"  [{i}] {json.dumps(info, ensure_ascii=False)}")

        # 所有按钮
        buttons = await page.query_selector_all('button')
        print(f"\n=== 按钮 ({len(buttons)}) ===")
        for i, btn in enumerate(buttons):
            text = (await btn.inner_text()).strip()[:80]
            cls = await btn.evaluate("el => el.className")
            print(f"  [{i}] text='{text}' class='{cls}'")

        # 所有包含class关键字的元素
        for tag in ['div', 'span', 'li', 'td', 'tr']:
            for keyword in ['guess', 'row', 'result', 'board', 'evaluation', 
                           'player', 'feedback', 'hint', 'column', 'col',
                           'green', 'yellow', 'gray', 'correct', 'arrow',
                           'celadon', 'up', 'down', 'match', 'answer',
                           'nickname', 'nation', 'team', 'age', 'role',
                           'champion', 'appearance', 'active', 'region']:
                els = await page.query_selector_all(f'{tag}[class*="{keyword}" i]')
                if els:
                    for e in els[:3]:
                        try:
                            cls = await e.evaluate("el => el.className")
                            text = (await e.inner_text()).strip()[:100]
                            print(f"  <{tag} class='{cls}'>: '{text}'")
                        except:
                            pass

        # 尝试输入一个选手名
        print("\n[4] 尝试输入选手...")
        guess_input = await page.query_selector('input[autocomplete="off"], input[type="text"], input:not([type="hidden"])')
        if guess_input:
            print("  找到输入框，输入 ZywOo...")
            await guess_input.click()
            await page.wait_for_timeout(500)
            # 逐个输入
            await guess_input.fill("ZywOo")
            await page.wait_for_timeout(2000)
            await page.screenshot(path="data/inspect_autocomplete.png", full_page=True)

            # 查找下拉选项
            options = await page.query_selector_all('[role="option"], [role="listbox"] li, ul[role="listbox"] > *, .autocomplete-item, [class*="suggest"]')
            print(f"  下拉选项: {len(options)} 个")
            for i, opt in enumerate(options[:10]):
                text = await opt.inner_text()
                cls = await opt.evaluate("el => el.className")
                print(f"    [{i}] text='{text.strip()}' class='{cls}'")

            # 按回车提交
            print("  按回车提交...")
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(3000)
            await page.screenshot(path="data/inspect_after_guess.png", full_page=True)
            
            html = await page.content()
            with open("data/inspect_after_guess.html", "w", encoding="utf-8") as f:
                f.write(html)

            # 分析反馈!!!
            print("\n[5] 分析反馈DOM...")
            # 搜索所有可能的反馈相关元素
            all_divs = await page.query_selector_all('div')
            feedback_info = []
            for div in all_divs:
                try:
                    cls = await div.evaluate("el => el.className") or ""
                    if any(kw in cls.lower() for kw in ['green', 'yellow', 'gray', 'correct', 'close', 'match', 'arrow', 'celadon', 'row', 'guess', 'evaluation', 'result', 'hint']):
                        text = (await div.inner_text()).strip()[:200]
                        html_snip = await div.evaluate("el => el.outerHTML.substring(0, 500)")
                        feedback_info.append({
                            'class': cls,
                            'text': text,
                            'html': html_snip
                        })
                except:
                    pass
            for fi in feedback_info[:30]:
                print(f"  class='{fi['class']}'")
                print(f"    text: {fi['text']}")
                print(f"    html: {fi['html']}")

        # API 调用
        if api_calls:
            print(f"\n=== API调用 ({len(api_calls)}条) ===")
            for c in api_calls[:20]:
                print(f"  {c}")

        # 页面可见文字
        text = await page.inner_text('body')
        print(f"\n=== 页面可见文字 ===\n{text[:3000]}")

        print("\n[DONE]")
        await browser.close()


if __name__ == "__main__":
    asyncio.run(inspect_game())
