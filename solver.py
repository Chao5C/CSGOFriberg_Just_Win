"""
CS2 Guessr 求解器
- 加载数据库
- 根据每次猜测的反馈（绿/黄/灰 + 箭头方向）过滤候选选手
- 排序并推荐下一轮最可能猜测的选手
"""
import os
import sys
import sqlite3
from typing import Optional

# 确保能导入 build_database
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
DB_PATH = os.path.join(DATA_DIR, "players.db")

# ============ 地区映射 (用于国家黄色匹配) ============
COUNTRY_REGION_MAP = {
    # Europe
    "Denmark": "Europe", "Sweden": "Europe", "Norway": "Europe",
    "Finland": "Europe", "France": "Europe", "Germany": "Europe",
    "UK": "Europe", "United Kingdom": "Europe", "Netherlands": "Europe",
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
    "Scotland": "Europe", "Wales": "Europe", "England": "Europe",
    # CIS
    "Russia": "CIS", "Ukraine": "CIS", "Belarus": "CIS",
    "Kazakhstan": "CIS", "Uzbekistan": "CIS", "Armenia": "CIS",
    "Azerbaijan": "CIS", "Georgia": "CIS", "Moldova": "CIS",
    "Kyrgyzstan": "CIS",
    # Americas
    "USA": "Americas", "United States": "Americas", "Canada": "Americas",
    "Brazil": "Americas", "Argentina": "Americas", "Chile": "Americas",
    "Mexico": "Americas", "Colombia": "Americas", "Peru": "Americas",
    "Uruguay": "Americas", "Venezuela": "Americas", "Ecuador": "Americas",
    "Guatemala": "Americas", "Dominican Republic": "Americas",
    "Puerto Rico": "Americas",
    # Asia
    "China": "Asia", "Mongolia": "Asia", "South Korea": "Asia",
    "Japan": "Asia", "Taiwan": "Asia", "Vietnam": "Asia",
    "Thailand": "Asia", "Indonesia": "Asia", "Malaysia": "Asia",
    "Philippines": "Asia", "India": "Asia", "Singapore": "Asia",
    "Hong Kong": "Asia", "Macau": "Asia",
    # Oceania
    "Australia": "Oceania", "New Zealand": "Oceania",
    # Africa
    "South Africa": "Africa", "Egypt": "Africa", "Morocco": "Africa",
    "Tunisia": "Africa", "Algeria": "Africa",
}

POSITION_GROUPS = {
    "AWPer": ["AWPer", "Sniper"],
    "IGL": ["IGL", "Captain", "In-Game Leader"],
    "Rifler": ["Rifler", "Entry", "Entry Fragger", "Lurker", "Anchor", "Fragger"],
    "Support": ["Support", "Supportive"],
    "Coach": ["Coach"],
}


def get_region(country: str) -> str:
    return COUNTRY_REGION_MAP.get(country, "Other")


def get_position_group(pos: str) -> str:
    pl = pos.lower()
    for group, variants in POSITION_GROUPS.items():
        for v in variants:
            if v.lower() in pl:
                return group
    return "Rifler"


# ============ 数据库加载 ============

def load_players() -> list:
    """从 SQLite 加载所有选手"""
    if not os.path.exists(DB_PATH):
        print(f"\n[错误] 数据库不存在: {DB_PATH}")
        print("请先运行: python build_database.py\n")
        return []

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM players ORDER BY nickname").fetchall()
    conn.close()

    players = []
    for r in rows:
        players.append(dict(r))
    return players


# ============ 反馈解析 ============

def parse_feedback(prompt: str, attr_name: str, is_numeric: bool = False) -> dict:
    """
    解析用户输入的单属性反馈
    返回: {"color": "green"|"yellow"|"gray", "direction": "higher"|"lower"|None}
    """
    result = {"color": "gray", "direction": None}
    raw = input(prompt).strip().lower()

    if not raw:
        # 空输入默认 gray
        result["color"] = "gray"
        return result

    # 识别颜色
    if raw[0] in ("g", "绿", "1"):
        result["color"] = "green"
    elif raw[0] in ("y", "黄", "2"):
        result["color"] = "yellow"
    elif raw[0] in ("r", "灰", "红", "3"):
        result["color"] = "gray"
    else:
        # 尝试解析完整输入如 "g", "y+", "r-"
        pass

    # 识别方向 (仅数值属性)
    if is_numeric:
        if "+" in raw or "↑" in raw or "高" in raw or "上" in raw:
            result["direction"] = "higher"
        elif "-" in raw or "↓" in raw or "低" in raw or "下" in raw:
            result["direction"] = "lower"

    return result


# ============ 筛选逻辑 ============

def filter_candidates(candidates: list, guess_player: dict, feedbacks: dict) -> list:
    """
    根据一次猜测的反馈，从候选列表中筛选符合的选手
    feedbacks = {
        "team": {"color": "gray"},
        "country": {"color": "yellow"},
        "age": {"color": "gray", "direction": "higher"},
        "position": {"color": "green"},
        "major_wins": {"color": "gray", "direction": "lower"},
        "major_appearances": {"color": "yellow"},
        "status": {"color": "green"},
    }
    """
    filtered = []

    for candidate in candidates:
        if candidate["id"] == guess_player["id"]:
            # 猜过的选手不再出现在候选中
            continue

        match = True

        # ---- 队伍 ----
        fb = feedbacks.get("team", {})
        if fb["color"] == "green":
            if candidate["team"].lower() != guess_player["team"].lower():
                match = False
        elif fb["color"] == "gray":
            if candidate["team"].lower() == guess_player["team"].lower():
                match = False
        # yellow 对队伍无意义，等同 gray

        if not match: continue

        # ---- 国家 ----
        fb = feedbacks.get("country", {})
        c_guess_country = guess_player["country"]
        c_cand_country = candidate["country"]

        if fb["color"] == "green":
            if c_cand_country.lower() != c_guess_country.lower():
                match = False
        elif fb["color"] == "yellow":
            # 同地区
            if get_region(c_cand_country) != get_region(c_guess_country):
                match = False
        elif fb["color"] == "gray":
            # 不同国家且不同地区
            if c_cand_country.lower() == c_guess_country.lower():
                match = False
            elif get_region(c_cand_country) == get_region(c_guess_country):
                match = False

        if not match: continue

        # ---- 年龄 ----
        fb = feedbacks.get("age", {})
        guess_age = guess_player["age"]
        cand_age = candidate["age"]

        if fb["color"] == "green":
            if cand_age != guess_age:
                match = False
        elif fb["color"] == "yellow":
            if abs(cand_age - guess_age) > 2:
                match = False
        elif fb["color"] == "gray":
            # 不接近：差 > 2
            if abs(cand_age - guess_age) <= 2:
                match = False

        # 方向校验
        if match and fb.get("direction"):
            if fb["direction"] == "higher":
                if cand_age >= guess_age:
                    match = False
            elif fb["direction"] == "lower":
                if cand_age <= guess_age:
                    match = False

        if not match: continue

        # ---- 位置 ----
        fb = feedbacks.get("position", {})
        if fb["color"] == "green":
            if candidate["position"].lower() != guess_player["position"].lower():
                match = False
        elif fb["color"] == "yellow":
            if get_position_group(candidate["position"]) != get_position_group(guess_player["position"]):
                match = False
        elif fb["color"] == "gray":
            if get_position_group(candidate["position"]) == get_position_group(guess_player["position"]):
                match = False

        if not match: continue

        # ---- Major 冠军 ----
        fb = feedbacks.get("major_wins", {})
        g_wins = guess_player["major_wins"]
        c_wins = candidate["major_wins"]

        if fb["color"] == "green":
            if c_wins != g_wins:
                match = False
        elif fb["color"] == "yellow":
            if abs(c_wins - g_wins) > 1:
                match = False
        elif fb["color"] == "gray":
            if abs(c_wins - g_wins) <= 1:
                match = False

        if match and fb.get("direction"):
            if fb["direction"] == "higher":
                if c_wins >= g_wins:
                    match = False
            elif fb["direction"] == "lower":
                if c_wins <= g_wins:
                    match = False

        if not match: continue

        # ---- Major 次数 ----
        fb = feedbacks.get("major_appearances", {})
        g_apps = guess_player["major_appearances"]
        c_apps = candidate["major_appearances"]

        if fb["color"] == "green":
            if c_apps != g_apps:
                match = False
        elif fb["color"] == "yellow":
            if abs(c_apps - g_apps) > 2:
                match = False
        elif fb["color"] == "gray":
            if abs(c_apps - g_apps) <= 2:
                match = False

        if match and fb.get("direction"):
            if fb["direction"] == "higher":
                if c_apps >= g_apps:
                    match = False
            elif fb["direction"] == "lower":
                if c_apps <= g_apps:
                    match = False

        if not match: continue

        # ---- 状态 ----
        fb = feedbacks.get("status", {})
        if fb["color"] == "green":
            if candidate["status"].lower() != guess_player["status"].lower():
                match = False
        elif fb["color"] == "gray":
            if candidate["status"].lower() == guess_player["status"].lower():
                match = False
        # yellow 对状态无意义

        if match:
            filtered.append(candidate)

    return filtered


# ============ 评分排序 ============

def score_candidates(candidates: list, all_guesses: list) -> list:
    """
    对候选选手打分排序
    评分逻辑：基于属性的信息量 + 数据库中的热度（rating）
    """
    scored = []
    for c in candidates:
        score = 0
        # 基础分：HLTV Rating 越高越可能是热门选手
        score += c.get("rating", 0) * 10
        # Major 冠军加分
        score += c.get("major_wins", 0) * 5
        # 活跃选手加分
        if c.get("status") == "Active":
            score += 10

        scored.append((c, score))

    scored.sort(key=lambda x: x[1], reverse=True)
    return scored


def recommend_guess(candidates: list, all_guesses: list) -> Optional[dict]:
    """
    推荐下一轮最优猜测
    策略：选择能最大化信息增益的选手（简化版：选Rating最高且属性分布差异大的）
    """
    if not candidates:
        return None

    if len(candidates) == 1:
        return candidates[0]

    # 简化为选 Rating 最高的
    return max(candidates, key=lambda c: c.get("rating", 0) or 0)


# ============ 交互式主循环 ============

def run_solver():
    """交互式求解器主程序"""
    players = load_players()
    if not players:
        return

    print("\n" + "=" * 60)
    print("  [CS2 Guessr 求解器]")
    print("  根据游戏反馈，逐步缩小范围，找出答案选手")
    print("=" * 60)
    print(f"\n  数据库已有 {len(players)} 位选手\n")

    # 初始候选 = 全部选手
    candidates = list(players)
    all_guesses = []  # 记录所有猜测过的选手昵称
    round_num = 0
    max_rounds = 8

    while round_num < max_rounds:
        round_num += 1
        print("-" * 60)
        print(f"  第 {round_num} 轮 | 候选: {len(candidates)} 位 | 剩余: {max_rounds - round_num} 次")
        print("-" * 60)

        if len(candidates) <= 10:
            print("\n  当前候选选手:")
            for i, c in enumerate(candidates, 1):
                print(f"    {i:2d}. {c['nickname']:20s} | {c['team']:15s} | "
                      f"{c['country']:12s} | Age:{c['age']:2d} | "
                      f"{c['position']:8s} | MajorW:{c['major_wins']} "
                      f"| 出席:{c['major_appearances']} | {c['status']}")

        # 推荐下一轮猜测
        recommendation = recommend_guess(candidates, all_guesses)
        if recommendation:
            print(f"\n  >> 推荐猜测: {recommendation['nickname']} "
                  f"({recommendation['team']}, {recommendation['country']}, "
                  f"{recommendation['position']}, Rating {recommendation.get('rating', 0)})")

        # 询问用户输入
        print(f"\n  请输入你在游戏中猜测的选手昵称:")
        guess_name = input("  > ").strip()

        if guess_name.lower() in ("quit", "exit", "q"):
            print("  退出求解器")
            break
        if guess_name.lower() == "list":
            # 显示全部候选
            print(f"\n  全部 {len(candidates)} 位候选:")
            for i, c in enumerate(candidates, 1):
                print(f"    {i:3d}. {c['nickname']:20s} | {c['team']:15s} | "
                      f"{c['country']:12s} | Age:{c['age']:2d} | "
                      f"{c['position']:8s} | Major:{c['major_wins']} "
                      f"| Apps:{c['major_appearances']} | {c['status']}")
            round_num -= 1
            continue
        if guess_name.lower() == "skip":
            round_num -= 1
            continue

        # 在数据库中查找选手
        guess_player = None
        for p in players:
            if p["nickname"].lower() == guess_name.lower():
                guess_player = p
                break

        if not guess_player:
            print(f"  [错误] 数据库中找不到选手 '{guess_name}'")
            print(f"  请检查拼写，或使用 'list' 查看候选列表")
            round_num -= 1
            continue

        print(f"\n  [{guess_player['nickname']}]: {guess_player['team']} | "
              f"{guess_player['country']} | Age:{guess_player['age']} | "
              f"{guess_player['position']} | "
              f"MajorW:{guess_player['major_wins']} | "
              f"Apps:{guess_player['major_appearances']} | "
              f"{guess_player['status']}")

        # 收集反馈
        print(f"\n  请输入游戏给出的反馈颜色:")
        print(f"     g=绿色(正确)  y=黄色(接近)  r=灰色(不匹配)")
        print(f"     数值属性可用 + 偏高 / - 偏低 (如: r- 表示灰色且偏高)\n")

        fb_team = parse_feedback(f"    队伍       [{guess_player['team']}]: ", "team")
        fb_country = parse_feedback(f"    国家       [{guess_player['country']}]: ", "country")
        fb_age = parse_feedback(f"    年龄       [{guess_player['age']}] [g/y/r +/-]: ", "age", is_numeric=True)
        fb_position = parse_feedback(f"    位置       [{guess_player['position']}]: ", "position")
        fb_major_wins = parse_feedback(f"    Major冠军  [{guess_player['major_wins']}] [g/y/r +/-]: ", "major_wins", is_numeric=True)
        fb_major_apps = parse_feedback(f"    Major次数  [{guess_player['major_appearances']}] [g/y/r +/-]: ", "major_appearances", is_numeric=True)
        fb_status = parse_feedback(f"    状态       [{guess_player['status']}]: ", "status")

        feedbacks = {
            "team": fb_team,
            "country": fb_country,
            "age": fb_age,
            "position": fb_position,
            "major_wins": fb_major_wins,
            "major_appearances": fb_major_apps,
            "status": fb_status,
        }

        # 检查是否全部绿色 = 猜对了
        all_green = all(fb["color"] == "green" for fb in feedbacks.values())
        if all_green:
            print(f"\n  *** 恭喜！答案就是 {guess_player['nickname']}！***")
            print(f"     用了 {round_num} 次猜测")
            break

        # 过滤候选
        candidates_before = len(candidates)
        candidates = filter_candidates(candidates, guess_player, feedbacks)
        all_guesses.append(guess_player["nickname"])

        removed = candidates_before - len(candidates)
        print(f"\n  本轮排除 {removed} 位，剩余 {len(candidates)} 位候选")

        if len(candidates) == 0:
            print("\n  [警告] 没有剩余候选！可能是反馈输入有误。")
            print("  请检查：")
            print("    1. 颜色选择是否正确 (g/y/r)")
            print("    2. 数值方向的箭头 (偏高选+, 偏低选-)")
            print("    3. 数据库中是否有该选手（可能需要补充数据）")
            # 允许重试本轮
            print("\n  按 'r' 重新输入本轮反馈，或其他键继续:")
            retry = input("  > ").strip().lower()
            if retry == "r":
                candidates = [p for p in players if p["nickname"] not in all_guesses]
                all_guesses.pop()  # 移除刚才的猜测
                round_num -= 1
                continue
            else:
                break

        if len(candidates) == 1:
            print(f"\n  !!! 只剩唯一候选！")
            print(f"     下一轮直接猜: {candidates[0]['nickname']}")

    else:
        # 8 次用完
        print(f"\n  ⏰ 8 次机会用完！")
        if candidates:
            print(f"\n  剩余 {len(candidates)} 位候选:")
            for i, c in enumerate(candidates[:15], 1):
                print(f"    {i:2d}. {c['nickname']:20s} | {c['team']:15s} | "
                      f"{c['country']:12s} | Age:{c['age']:2d} | "
                      f"{c['position']:8s} | Major:{c['major_wins']} | {c['status']}")

    print("\n" + "=" * 60)
    print("  求解结束")
    print("=" * 60 + "\n")


# ============ 批量测试模式 ============

def test_solver(answer_nickname: str):
    """
    测试模式：模拟用求解器找答案
    输入正确答案昵称，自动模拟游戏反馈
    """
    players = load_players()
    if not players:
        return

    answer = None
    for p in players:
        if p["nickname"].lower() == answer_nickname.lower():
            answer = p
            break

    if not answer:
        print(f"找不到选手: {answer_nickname}")
        return

    print(f"\n答案: {answer['nickname']} ({answer['team']})")
    print("-" * 50)

    candidates = list(players)
    all_guesses = []

    for round_num in range(1, 9):
        if len(candidates) == 1 and candidates[0]["id"] == answer["id"]:
            print(f"第{round_num}轮: *** 找到! {answer['nickname']}")
            break

        # 选推荐作为猜测
        guess = recommend_guess(candidates, all_guesses)
        if not guess:
            print("无可用推荐")
            break

        # 模拟反馈
        feedbacks = simulate_feedback(guess, answer)
        print(f"第{round_num}轮: 猜 {guess['nickname']:20s}", end=" ")

        # 打印颜色
        colors = []
        for attr in ["team", "country", "age", "position", "major_wins", "major_appearances", "status"]:
            fb = feedbacks[attr]
            c = {"green": "[G]", "yellow": "[Y]", "gray": "[R]"}[fb["color"]]
            if fb.get("direction") == "higher":
                c += "^"
            elif fb.get("direction") == "lower":
                c += "v"
            colors.append(c)
        print(" ".join(colors))

        all_guesses.append(guess["nickname"])
        candidates = filter_candidates(candidates, guess, feedbacks)

        if guess["id"] == answer["id"]:
            print(f"           *** 猜对了!")
            break

    else:
        print("8次用完")

    print(f"剩余候选: {len(candidates)}")
    return


def simulate_feedback(guess: dict, answer: dict) -> dict:
    """模拟游戏反馈"""
    def _feedback(attr, is_numeric=False, numeric_range=2):
        g_val = guess[attr]
        a_val = answer[attr]

        if isinstance(g_val, str) and isinstance(a_val, str):
            if g_val.lower() == a_val.lower():
                return {"color": "green", "direction": None}
        elif is_numeric:
            diff = g_val - a_val
            direction = "higher" if diff > 0 else ("lower" if diff < 0 else None)
            if diff == 0:
                return {"color": "green", "direction": None}
            elif abs(diff) <= numeric_range:
                return {"color": "yellow", "direction": direction}
            else:
                return {"color": "gray", "direction": direction}

        # 字符串属性不匹配处理
        if attr == "country":
            if get_region(g_val) == get_region(a_val):
                return {"color": "yellow", "direction": None}
        elif attr == "position":
            if get_position_group(g_val) == get_position_group(a_val):
                return {"color": "yellow", "direction": None}
        return {"color": "gray", "direction": None}

    return {
        "team": _feedback("team"),
        "country": _feedback("country"),
        "age": _feedback("age", is_numeric=True, numeric_range=2),
        "position": _feedback("position"),
        "major_wins": _feedback("major_wins", is_numeric=True, numeric_range=1),
        "major_appearances": _feedback("major_appearances", is_numeric=True, numeric_range=2),
        "status": _feedback("status"),
    }


# ============ 入口 ============

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        test_nickname = sys.argv[2] if len(sys.argv) > 2 else "ZywOo"
        test_solver(test_nickname)
    else:
        run_solver()
