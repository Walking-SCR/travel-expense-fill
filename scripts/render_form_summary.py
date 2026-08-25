"""
render_form_summary.py — 生成网页报销填写清单。

只输出网页必填字段的 Markdown 清单，不操作网页、不上传、不提交。
"""
from datetime import date as date_cls


DEFAULT_FORM_CONFIG = {
    "department": "技术服务群/智能数据技术服务部/AI应用支持二中心",
    "title": "差旅报销",
    "reimbursement_entity": "悦智人工智能（深圳）有限责任公司",
    "currency": "人民币",
    "original_loan": 0,
    "business_code": "无",
}


def normalize_form_config(config):
    merged = dict(DEFAULT_FORM_CONFIG)
    if config:
        merged.update({k: v for k, v in config.items() if v is not None})
    return merged


def _amount(value):
    try:
        return round(float(value or 0), 2)
    except (TypeError, ValueError):
        return 0.0


def _ride_amount(entry):
    if isinstance(entry, list):
        return sum(_ride_amount(item) for item in entry)
    if isinstance(entry, dict):
        return _amount(entry.get("amount"))
    return _amount(entry)


def calculate_trip_total(days_data, train_fares=None, didi_trip=None):
    """按写入差旅费 Excel 的每日口径计算总额。"""
    train_fares = train_fares or {}
    didi_trip = didi_trip or {}
    total = 0.0
    for day, info in (days_data or {}).items():
        total += _amount(info.get("plane"))
        total += _amount(train_fares.get(day))
        total += _amount(info.get("bus"))
        total += _ride_amount(didi_trip.get(day))
        total += _amount(info.get("hotel"))
        total += _amount(info.get("subsidy"))
    return round(total, 2)


def calculate_base_total(base_amounts):
    """按 Base 地交通费 Excel 的每笔行程口径计算总额。"""
    total = 0.0
    for value in (base_amounts or {}).values():
        total += _ride_amount(value)
    return round(total, 2)


def _money(value):
    return f"{_amount(value):.2f}"


def build_form_data(
    *,
    form_config,
    year,
    month,
    out_trip,
    out_base=None,
    days_data=None,
    train_fares=None,
    didi_trip=None,
    base_amounts=None,
    report_date=None,
):
    config = normalize_form_config(form_config)
    report_date = report_date or date_cls.today()
    trip_total = calculate_trip_total(days_data, train_fares, didi_trip)
    base_total = calculate_base_total(base_amounts)
    reimbursement_total = round(trip_total + base_total, 2)
    original_loan = _amount(config.get("original_loan"))
    payable_balance = round(reimbursement_total - original_loan, 2)

    attachments = [str(out_trip)] if out_trip else []
    if out_base:
        attachments.append(str(out_base))

    return {
        "application": {
            "department": config["department"],
            "title": config["title"],
            "reimbursement_entity": config["reimbursement_entity"],
            "report_date": report_date.isoformat(),
        },
        "expense": {
            "summary": f"{int(month)}月差旅报销",
            "reimbursement_total": reimbursement_total,
            "currency": config["currency"],
            "original_loan": original_loan,
            "payable_balance": payable_balance,
            "business_code": config["business_code"],
            "attachments": attachments,
        },
        "totals": {
            "trip_total": trip_total,
            "base_total": base_total,
            "reimbursement_total": reimbursement_total,
        },
        "period": {
            "year": int(year),
            "month": int(month),
        },
    }


def render_markdown(form_data):
    app = form_data["application"]
    exp = form_data["expense"]
    attachments = exp.get("attachments") or []
    attachment_lines = "\n".join(f"- `{path}`" for path in attachments) or "- 无"

    return "\n".join([
        "# 报销网页填写清单",
        "",
        "> 仅包含网页截图中带 `*` 的必填字段；未标星字段无需填写。",
        "",
        "## 申请详情",
        "",
        "| 字段 | 填写内容 |",
        "|---|---|",
        f"| 所属部门 | {app['department']} |",
        f"| 标题 | {app['title']} |",
        f"| 报销主体 | {app['reimbursement_entity']} |",
        f"| 报销日期 | {app['report_date']} |",
        "",
        "## 报销列表",
        "",
        "| 字段 | 填写内容 |",
        "|---|---|",
        f"| 摘要 | {exp['summary']} |",
        f"| 报销合计 | {_money(exp['reimbursement_total'])} |",
        f"| 币种 | {exp['currency']} |",
        f"| 原借款 | {_money(exp['original_loan'])} |",
        f"| 应付余额 | {_money(exp['payable_balance'])} |",
        f"| 业务编号 | {exp['business_code']} |",
        "",
        "## 附件",
        "",
        attachment_lines,
        "",
    ])


def render_form_summary(**kwargs):
    form_data = build_form_data(**kwargs)
    return render_markdown(form_data), form_data
