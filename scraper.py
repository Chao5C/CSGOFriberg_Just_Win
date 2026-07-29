"""
CS2 选手数据爬虫
- 从 HLTV.org 抓取选手数据
- 支持 Playwright 浏览器自动化 + 截图 OCR 两种方式
"""
import re
import os
import asyncio
from playwright.async_api import async_playwright

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
SCREENSHOT_DIR = os.path.join(DATA_DIR, "screenshots")


def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)


# ============ 地区映射 ============
COUNTRY_REGION = {
    "Denmark": "Europe", "Sweden": "Europe", "Norway": "Europe",
    "Finland": "Europe", "France": "Europe", "Germany": "Europe",
    "United Kingdom": "Europe", "UK": "Europe", "Netherlands": "Europe",
    "Poland": "Europe", "Spain": "Europe", "Portugal": "Europe",
    "Belgium": "Europe", "Switzerland": "Europe", "Austria": "Europe",
    "Estonia": "Europe", "Latvia": "Europe", "Lithuania": "Europe",
    "Czech Republic": "Europe", "Slovakia": "Europe", "Hungary": "Europe",
    "Romania": "Europe", "Bulgaria": "Europe", "Serbia": "Europe",
    "Croatia": "Europe", "Slovenia": "Europe", "Bosnia and Herzegovina": "Europe",
    "Bosnia": "Europe", "Montenegro": "Europe", "North Macedonia": "Europe",
    "Kosovo": "Europe", "Turkey": "Europe", "Israel": "Europe",
    "Italy": "Europe", "Greece": "Europe", "Iceland": "Europe",
    "Ireland": "Europe", "Malta": "Europe", "Luxembourg": "Europe",
    "Russia": "CIS", "Ukraine": "CIS", "Belarus": "CIS",
    "Kazakhstan": "CIS", "Uzbekistan": "CIS", "Armenia": "CIS",
    "Azerbaijan": "CIS", "Georgia": "CIS", "Moldova": "CIS", "Kyrgyzstan": "CIS",
    "United States": "Americas", "USA": "Americas", "Canada": "Americas",
    "Brazil": "Americas", "Argentina": "Americas", "Chile": "Americas",
    "Mexico": "Americas", "Colombia": "Americas", "Peru": "Americas",
    "Uruguay": "Americas", "Venezuela": "Americas", "Ecuador": "Americas",
    "Guatemala": "Americas", "Dominican Republic": "Americas",
    "Puerto Rico": "Americas",
    "China": "Asia", "Mongolia": "Asia", "South Korea": "Asia",
    "Japan": "Asia", "Taiwan": "Asia", "Vietnam": "Asia",
    "Thailand": "Asia", "Indonesia": "Asia", "Malaysia": "Asia",
    "Philippines": "Asia", "India": "Asia", "Singapore": "Asia",
    "Hong Kong": "Asia", "Macau": "Asia",
    "Australia": "Oceania", "New Zealand": "Oceania",
    "South Africa": "Africa", "Egypt": "Africa", "Morocco": "Africa",
    "Tunisia": "Africa", "Algeria": "Africa",
}

# 状态关键词
STATUS_ACTIVE = ["active", "playing", "benched", "stand-in", "substitute"]
STATUS_INACTIVE = ["inactive", "away", "on leave"]
STATUS_RETIRED = ["retired", "coach"]


def get_region(country: str) -> str:
    return COUNTRY_REGION.get(country, "Unknown")


def detect_status(text: str) -> str:
    """从文本检测选手状态"""
    t = text.lower()
    for kw in STATUS_RETIRED:
        if kw in t:
            return "Retired"
    for kw in STATUS_INACTIVE:
        if kw in t:
            return "Inactive"
    return "Active"


def clean_nickname(name: str) -> str:
    """清理选手昵称（去除多余空格和特殊字符）"""
    return name.strip()


# ============ 方式一：Playwright 浏览器抓取 ============

async def scrape_hltv_players(max_players: int = 100) -> list:
    """
    从 HLTV.org 选手统计页面抓取数据
    https://www.hltv.org/stats/players
    """
    ensure_dirs()
    players = []
    seen_names = set()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            viewport={"width": 1920, "height": 1080}
        )
        page = await context.new_page()

        try:
            url = "https://www.hltv.org/stats/players"
            print(f"[爬虫] 访问 {url}")
            await page.goto(url, wait_until="networkidle", timeout=30000)

            # 等待表格
            await page.wait_for_selector("table.stats-table", timeout=15000)

            page_num = 1
            while len(players) < max_players:
                rows = await page.query_selector_all("table.stats-table tbody tr")
                print(f"[爬虫] 第{page_num}页, 读取{len(rows)}行")

                for row in rows:
                    if len(players) >= max_players:
                        break
                    try:
                        pdata = await _parse_player_row(page, row)
                        if pdata and pdata["nickname"] not in seen_names:
                            seen_names.add(pdata["nickname"])
                            players.append(pdata)
                    except Exception as e:
                        continue

                # 翻页
                next_btn = await page.query_selector("a.pagination-next:not(.disabled)")
                if not next_btn:
                    break
                await next_btn.click()
                await page.wait_for_timeout(2000)
                page_num += 1

        except Exception as e:
            print(f"[爬虫] 抓取出错: {e}")
        finally:
            await browser.close()

    print(f"[爬虫] 共获取 {len(players)} 位选手")
    return players


async def _parse_player_row(page, row) -> dict:
    """解析 HLTV 排行榜一行"""
    cells = await row.query_selector_all("td")
    if len(cells) < 8:
        return None

    # 第1列: 选手名
    name_el = await cells[0].query_selector("a")
    if not name_el:
        return None
    nickname = clean_nickname((await name_el.inner_text()).strip())
    profile_url = await name_el.get_attribute("href")

    # 第2列: 队伍
    team = "Unknown"
    team_el = await cells[1].query_selector("a")
    if team_el:
        team = (await team_el.inner_text()).strip()

    # 第8列: Rating
    rating = 0.0
    try:
        rating = float((await cells[7].inner_text()).strip())
    except Exception:
        pass

    # 进入详情页
    nationality = "Unknown"
    age = 0
    position = "Rifler"
    major_wins = 0
    major_appearances = 0
    status = "Active"
    image_url = ""

    if profile_url:
        detail = await _scrape_detail(page, profile_url)
        nationality = detail.get("nationality", "Unknown")
        age = detail.get("age", 0)
        position = detail.get("position", "Rifler")
        major_wins = detail.get("major_wins", 0)
        major_appearances = detail.get("major_appearances", 0)
        status = detail.get("status", "Active")
        image_url = detail.get("image_url", "")

    region = get_region(nationality)

    return {
        "nickname": nickname,
        "team": team,
        "country": nationality,
        "age": age if age > 0 else 22,
        "position": position,
        "major_wins": major_wins,
        "major_appearances": major_appearances,
        "status": status,
        "rating": rating,
        "image_url": image_url,
    }


async def _scrape_detail(page, profile_url: str) -> dict:
    """进入选手详情页抓取详细信息"""
    full_url = f"https://www.hltv.org{profile_url}"
    detail = {"nationality": "Unknown", "age": 0, "position": "Rifler",
              "major_wins": 0, "major_appearances": 0, "status": "Active",
              "image_url": ""}

    new_page = await page.context.new_page()
    try:
        await new_page.goto(full_url, wait_until="networkidle", timeout=20000)

        # 年龄
        try:
            age_el = await new_page.query_selector("div.playerAge span")
            if age_el:
                m = re.search(r"(\d+)", await age_el.inner_text())
                if m:
                    detail["age"] = int(m.group(1))
        except Exception:
            pass

        # 国籍 (从国旗 alt 属性)
        try:
            flag_el = await new_page.query_selector("img.flag[title]")
            if flag_el:
                detail["nationality"] = (await flag_el.get_attribute("title")).strip()
        except Exception:
            pass

        # 图片
        try:
            img_el = await new_page.query_selector("img.player-picture-img")
            if img_el:
                detail["image_url"] = await img_el.get_attribute("src") or ""
        except Exception:
            pass

        # 战队历史 → 推断 Major 出席次数和职业年数
        try:
            team_rows = await new_page.query_selector_all("div.player-team-row, div.team-history-row")
            years_set = set()
            for tr in team_rows:
                text = (await tr.inner_text()).strip()
                years = re.findall(r"20(\d{2})", text)
                for y in years:
                    years_set.add(int("20" + y))
            if years_set:
                earliest = min(years_set)
                detail["major_appearances"] = max(1, 2025 - earliest)
        except Exception:
            pass

        # 成就 → Major 冠军
        try:
            ach_els = await new_page.query_selector_all("div.achievement, div.trophy")
            major_count = 0
            for el in ach_els:
                text = (await el.inner_text()).strip().lower()
                if "major" in text:
                    m = re.search(r"(\d+)x?\s*major", text, re.IGNORECASE)
                    if m:
                        major_count = max(major_count, int(m.group(1)))
            detail["major_wins"] = major_count
        except Exception:
            pass

        # 状态
        try:
            body_text = (await new_page.inner_text("body")).lower()
            detail["status"] = detect_status(body_text)
        except Exception:
            pass

    except Exception as e:
        print(f"  [详情页错误] {profile_url}: {e}")
    finally:
        await new_page.close()

    return detail


async def scrape_hltv_team_players(max_players: int = 100) -> list:
    """
    备选方案：从 HLTV 排名页面抓取（更稳定）
    https://www.hltv.org/ranking/teams/2025
    """
    ensure_dirs()
    players = []
    seen_names = set()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            viewport={"width": 1920, "height": 1080}
        )
        page = await context.new_page()

        try:
            # 从排名页获取队伍
            await page.goto("https://www.hltv.org/ranking/teams/2025",
                           wait_until="networkidle", timeout=30000)
            await page.wait_for_selector("div.ranked-team", timeout=10000)

            team_links = await page.query_selector_all("a.moreLink[href*='/team/']")
            team_urls = []
            for link in team_links[:30]:
                href = await link.get_attribute("href")
                if href:
                    team_urls.append(href)

            for turl in team_urls:
                if len(players) >= max_players:
                    break
                try:
                    team_players = await _scrape_team_roster(page, turl)
                    for tp in team_players:
                        if tp["nickname"] not in seen_names:
                            seen_names.add(tp["nickname"])
                            players.append(tp)
                except Exception as e:
                    print(f"  [队伍错误] {turl}: {e}")
                    continue

        except Exception as e:
            print(f"[爬虫] 抓取出错: {e}")
        finally:
            await browser.close()

    print(f"[爬虫] 共获取 {len(players)} 位选手")
    return players


async def _scrape_team_roster(page, team_url: str) -> list:
    """抓取队伍阵容"""
    full_url = f"https://www.hltv.org{team_url}"
    roster = []

    await page.goto(full_url, wait_until="networkidle", timeout=20000)

    # 队伍名
    team_name = "Unknown"
    try:
        team_el = await page.query_selector("h1.profile-team-name")
        if team_el:
            team_name = (await team_el.inner_text()).strip()
    except Exception:
        pass

    # 队员列表
    player_els = await page.query_selector_all("div.player-container a[href*='/player/']")
    for pel in player_els:
        try:
            nickname = clean_nickname((await pel.inner_text()).strip())
            if not nickname or len(nickname) < 2:
                continue
            href = await pel.get_attribute("href")
            detail = await _scrape_detail(page, href)
            region = get_region(detail.get("nationality", "Unknown"))
            roster.append({
                "nickname": nickname,
                "team": team_name,
                "country": detail.get("nationality", "Unknown"),
                "age": detail.get("age", 22),
                "position": detail.get("position", "Rifler"),
                "major_wins": detail.get("major_wins", 0),
                "major_appearances": detail.get("major_appearances", 0),
                "status": detail.get("status", "Active"),
                "rating": 1.0,
                "image_url": detail.get("image_url", ""),
            })
        except Exception:
            continue

    return roster


# ============ 方式二：截图 + OCR 文字识别 ============

async def screenshot_and_ocr(url: str = "https://www.hltv.org/stats/players",
                              output: str = None) -> str:
    """
    截图网页并通过OCR提取文字（适用于动态加载页面）
    需要安装: pip install pytesseract Pillow
    """
    ensure_dirs()
    if output is None:
        output = os.path.join(SCREENSHOT_DIR, "page_screenshot.png")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})
        await page.goto(url, wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(3000)
        await page.screenshot(path=output, full_page=True)
        await browser.close()

    print(f"[截图] 已保存至 {output}")

    # 尝试 OCR
    text = ""
    try:
        from PIL import Image
        import pytesseract
        img = Image.open(output)
        text = pytesseract.image_to_string(img, lang="eng")
        print(f"[OCR] 提取到 {len(text)} 个字符")
    except ImportError:
        print("[OCR] pytesseract 未安装，仅保存截图。安装: pip install pytesseract Pillow")
        print("[OCR] 还需安装 Tesseract: https://github.com/tesseract-ocr/tesseract")

    return text


def parse_ocr_text(text: str) -> list:
    """从 OCR 文本中解析选手数据"""
    players = []
    lines = text.strip().split("\n")
    for line in lines:
        # HLTV 排行榜格式: # Name Team Maps Rounds K-D K/D Rating
        # 简化正则匹配
        match = re.match(
            r"(\d+)\s+([A-Za-z0-9_\-]+)\s+([A-Za-z0-9\s]+?)\s+(\d+)\s+.*?(\d+\.\d+)", line
        )
        if match:
            players.append({
                "nickname": match.group(2),
                "team": match.group(3).strip(),
                "country": "Unknown",
                "age": 0,
                "position": "Rifler",
                "major_wins": 0,
                "major_appearances": 0,
                "status": "Active",
                "rating": float(match.group(5)),
                "image_url": "",
            })
    return players


# ============ 统一入口 ============

def scrape(method: str = "hltv", max_players: int = 100) -> list:
    """
    统一抓取入口
    method: "hltv" | "team" | "screenshot"
    """
    ensure_dirs()

    if method == "screenshot":
        text = asyncio.run(screenshot_and_ocr())
        return parse_ocr_text(text)
    elif method == "team":
        return asyncio.run(scrape_hltv_team_players(max_players))
    else:
        return asyncio.run(scrape_hltv_players(max_players))


if __name__ == "__main__":
    import sys
    method = sys.argv[1] if len(sys.argv) > 1 else "hltv"
    count = int(sys.argv[2]) if len(sys.argv) > 2 else 50

    print(f"=" * 50)
    print(f"  CS2 选手数据爬虫")
    print(f"  方式: {method} | 目标: {count} 位")
    print(f"=" * 50)

    result = scrape(method=method, max_players=count)
    for i, p in enumerate(result, 1):
        print(f"  {i:3d}. {p['nickname']:20s} | {p['team']:15s} | {p['country']:15s} | "
              f"Age:{p['age']:2d} | {p['position']:10s} | "
              f"Major:{p['major_wins']} | App:{p['major_appearances']} | {p['status']}")

    print(f"\n  共获取 {len(result)} 位选手")
