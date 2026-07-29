# **V0.1 2026.7.29**
## 实现最基本的爬取网页和建立数据库，在搜索和推理方面还是较弱，下一版本将针对推理方法进一步优化。
## 未来在建立完善的数据库后引入多个agent协同，加速推理过程，目标推理时间在2s内。
# v0.2
1. 引入 random 模块（第 20 行）
diff
+ import random
2. recommend_guess → select_guess（第 380-386 行）
python
def select_guess(candidates, guessed_names):
    """在候选集中随机挑选一个（纯随机策略，不用 rating 评分）"""
    available = [c for c in candidates
                 if c["nickname"].lower() not in guessed_names]
    if not available:
        return None
    return random.choice(available)
删除了 (rating*10 + major_wins*5) 的评分排序逻辑
改为 random.choice 纯随机挑选
3. filter_candidates 新增 relax_yellow 参数（第 254 行）
relax_yellow=False（默认）：严格模式，黄色要求就近匹配（同地区/±2年龄等）
relax_yellow=True：降级模式，黄色仅排除精确值，不限制就近范围
4. 降级过滤策略（第 661-676 行）
code
严格过滤 → 0 结果
  → [降级L1] 放宽黄色约束，重新过滤
  → 0 结果
  → [降级L2] 重置候选池（排除已猜选手）
5. 删除未使用函数
移除 get_close_values
移除 filter_relaxed
6. 日志输出改进
"候选: X 位" → "候选范围: X 位"
"排除 X 位，剩余 X 位" → "严格过滤: 排除 X 位，剩余 X 位"
新增 [降级L1] / [降级L2] 标签
整体流程：每轮猜测
code
第1轮: 从全库随机抽1个
第2-N轮:
  1. 上一轮反馈 → 严格过滤候选集
  2. 候选>0 → 随机挑1个
  3. 候选=0 → L1放宽黄色 → 随机挑1个
  4. 仍为0 → L2重置候选池 → 随机挑1个
