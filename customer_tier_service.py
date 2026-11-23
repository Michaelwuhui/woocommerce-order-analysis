"""统一的客户等级评定、不良退货/拒收降级熔断与风控建议服务。

核心原则：
1. 全站（订单列表、首页概览、客户管理列表、客户详情弹窗）共享同一套等级评定标准。
2. 任何存在「问题退货」（调包/少件/损坏/扣留货损）或高频拒收未送达的风控客户，严禁被评为「VIP」或「优质」。
3. 对累犯问题退货、已被拉入手机号黑名单或高频拒收的客户直接判定为「劣质」。
"""

from __future__ import annotations

from datetime import datetime
from typing import Any


TIER_ORDER = ("bad", "new", "normal", "good", "vip")

TIER_CODE_TO_CN = {
    "vip": "VIP",
    "good": "优质",
    "normal": "普通",
    "new": "新客",
    "bad": "劣质",
}

TIER_CN_TO_CODE = {v: k for k, v in TIER_CODE_TO_CN.items()}

ORDER_BADGE_META = {
    "vip": {"label": "VIP", "class": "text-warning", "icon": "star-fill"},
    "good": {"label": "优质", "class": "text-success", "icon": "gem"},
    "normal": {"label": "普通", "class": "text-primary", "icon": "person-check"},
    "new": {"label": "新客", "class": "text-info", "icon": "stars"},
    "bad": {"label": "劣质", "class": "text-danger", "icon": "x-circle"},
}

DETAIL_BADGE_META = {
    "vip": {"level": "VIP", "color": "#f59e0b", "icon": "star-fill"},
    "good": {"level": "优质", "color": "#10b981", "icon": "gem"},
    "normal": {"level": "普通", "color": "#3b82f6", "icon": "person-check"},
    "new": {"level": "新客", "color": "#0dcaf0", "icon": "stars"},
    "bad": {"level": "劣质", "color": "#ef4444", "icon": "x-circle"},
}


def compute_base_quality_score(
    successful_orders: int,
    total_spending: float,
    avg_days_between: float,
) -> int:
    """计算不含退货惩罚的基础消费活跃分（0~100）。"""
    succ = max(0, int(successful_orders or 0))
    spend = max(0.0, float(total_spending or 0.0))
    avg_days = float(avg_days_between or 0.0)

    score = 0.0
    score += min(succ * 10, 30)
    score += min(spend / 100.0, 40.0)
    if avg_days > 0:
        if avg_days < 60:
            score += 30
        elif avg_days < 120:
            score += 15
    return int(round(min(score, 100.0)))


def compute_avg_days_between(
    first_order_date: str | None,
    last_order_date: str | None,
    order_count: int,
) -> float:
    """根据首末单时间与订单数计算平均复购间隔天数。"""
    count = int(order_count or 0)
    if count <= 1 or not first_order_date or not last_order_date:
        return 0.0
    try:
        first_dt = datetime.fromisoformat(str(first_order_date)[:19])
        last_dt = datetime.fromisoformat(str(last_order_date)[:19])
        days_span = max(0, (last_dt - first_dt).days)
        return days_span / (count - 1)
    except (TypeError, ValueError):
        return 0.0


def calculate_customer_tier(
    successful_orders: int,
    total_spending: float,
    avg_days_between: float,
    bad_orders: int = 0,
    meaningful_orders: int = 0,
    *,
    problem_return_orders: int = 0,
    undelivered_orders: int = 0,
    is_blocked: bool = False,
) -> str:
    """综合成功消费表现与退货/拒收/拉黑风控记录评定客户等级代码（vip/good/normal/new/bad）。"""
    if is_blocked:
        return "bad"

    succ_cnt = max(0, int(successful_orders or 0))
    try:
        prob_cnt = max(0, int(problem_return_orders or 0))
    except (TypeError, ValueError):
        prob_cnt = 0
    try:
        undeliv_cnt = max(0, int(undelivered_orders or 0))
    except (TypeError, ValueError):
        undeliv_cnt = 0
    try:
        raw_bad = max(0, int(bad_orders or 0))
    except (TypeError, ValueError):
        raw_bad = 0

    bad_cnt = max(raw_bad, prob_cnt + undeliv_cnt)
    try:
        raw_meaningful = max(0, int(meaningful_orders or 0))
    except (TypeError, ValueError):
        raw_meaningful = 0
    denom = max(raw_meaningful, bad_cnt + succ_cnt)
    bad_rate = (bad_cnt / denom) if denom > 0 else 0.0

    # 一级风控熔断：凡存在问题退货（调包/少件/损坏/扣留货损，属高风险）或惯性拒收，直接判定为劣质客户（bad）
    if prob_cnt >= 1:
        return "bad"
    if bad_cnt >= 3 and bad_rate >= 0.35:
        return "bad"
    if bad_cnt >= 2 and bad_rate >= 0.50:
        return "bad"

    # 基础分档
    quality_score = compute_base_quality_score(
        succ_cnt, total_spending, avg_days_between
    )
    if quality_score >= 80 and succ_cnt >= 2:
        tier = "vip"
    elif quality_score >= 60 and succ_cnt >= 1:
        tier = "good"
    elif quality_score >= 40 and succ_cnt >= 1:
        tier = "normal"
    else:
        tier = "new"

    rank = ["new", "normal", "good", "vip"]

    def _cap(current: str, ceiling: str) -> str:
        return current if rank.index(current) <= rank.index(ceiling) else ceiling

    # 二级风控封顶：任何问题退货（高风险）、2次及以上拒收/退回或不良率≥20%，严禁评为 VIP 或 优质
    if prob_cnt >= 1 or bad_cnt >= 2 or bad_rate >= 0.20:
        tier = _cap(tier, "normal")
    elif bad_cnt >= 1:
        # 偶发 1 次普通未取件且无问题退货的老客：最高不超过「优质」，不给 VIP
        tier = _cap(tier, "good")

    return tier


def normalize_manual_tier(value: str | None) -> str:
    """标准化手动设置的等级编码；未知或空值统一归为 auto。"""
    clean = str(value or "").strip().lower()
    if clean in {"vip", "good", "normal", "new", "bad"}:
        return clean
    return "auto"


def resolve_final_tier(
    *,
    manual_tier: str | None,
    successful_orders: int,
    total_spending: float,
    avg_days_between: float,
    bad_orders: int = 0,
    meaningful_orders: int = 0,
    problem_return_orders: int = 0,
    undelivered_orders: int = 0,
    is_blocked: bool = False,
) -> tuple[str, bool]:
    """返回 (最终等级代码, 是否手动指定)。已被拉黑拦截的客户强制为 bad。"""
    if is_blocked:
        return "bad", False
    norm_manual = normalize_manual_tier(manual_tier)
    if norm_manual != "auto":
        return norm_manual, True
    auto_code = calculate_customer_tier(
        successful_orders,
        total_spending,
        avg_days_between,
        bad_orders=bad_orders,
        meaningful_orders=meaningful_orders,
        problem_return_orders=problem_return_orders,
        undelivered_orders=undelivered_orders,
        is_blocked=is_blocked,
    )
    return auto_code, False


def build_smart_actions(
    *,
    tier_cn: str,
    successful_orders: int,
    problem_return_orders: int,
    undelivered_orders: int,
    is_blocked: bool,
    first_order_date: str | None,
    last_order_date: str | None,
    thirty_days_ago: str,
    ninety_days_ago: str,
) -> list[dict[str, str]]:
    """根据客户真实等级与风控画像生成客户管理列表的运营建议标签。"""
    actions: list[dict[str, str]] = []
    if is_blocked:
        actions.append({"type": "danger", "icon": "shield-fill-x", "text": "已拉黑拦截"})
        return actions

    if tier_cn == "劣质":
        actions.append({"type": "danger", "icon": "shield-exclamation", "text": "建议加入拦截"})
        return actions

    if int(problem_return_orders or 0) > 0:
        actions.append({"type": "danger", "icon": "exclamation-octagon", "text": "问题退货·严审COD"})
        return actions

    if int(undelivered_orders or 0) >= 2:
        actions.append({"type": "warning", "icon": "exclamation-triangle", "text": "多次未取·发货核实"})
        return actions

    if tier_cn == "VIP":
        actions.append({"type": "success", "icon": "gift", "text": "专属礼遇"})
        actions.append({"type": "primary", "icon": "people", "text": "邀请入群"})
    elif last_order_date and last_order_date < ninety_days_ago and int(undelivered_orders or 0) == 0:
        actions.append({"type": "warning", "icon": "ticket-perforated", "text": "召回优惠券"})
    elif int(successful_orders or 0) == 1 and first_order_date and first_order_date >= thirty_days_ago:
        actions.append({"type": "info", "icon": "book", "text": "欢迎指南"})
        actions.append({"type": "info", "icon": "bag-plus", "text": "关联推荐"})
    elif int(successful_orders or 0) > 3 and tier_cn in {"优质", "普通"}:
        actions.append({"type": "primary", "icon": "arrow-repeat", "text": "复购跟进"})
    return actions
