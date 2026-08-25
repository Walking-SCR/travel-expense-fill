"""
gen_trip_reports.py — 差旅报销单生成工具（CLI 可复用版）

自动发现出差目录中的携程/滴滴/大巴/高铁文件，生成差旅费和Base地交通费Excel。

用法:
    python3 gen_trip_reports.py --dir /path/to/出差目录
    python3 gen_trip_reports.py --dir /path/to/2026年04出差 --non-interactive
    python3 gen_trip_reports.py --init          # 仅初始化配置+检查依赖

配置文件: ../config.json
  - pdf_parser.tool: PDF解析工具（默认 pymupdf，auto_install 自动下载）
  - image_parser.mcp_available: 是否配置了图片MCP工具
  - image_parser.mcp_tool_name: MCP工具名称（如 minimax_understand_image）
"""
import sys, json, re, argparse, subprocess, importlib
from pathlib import Path
from datetime import date as date_cls

# ── PyMuPDF 自动安装 ──
try:
    import fitz
except ImportError:
    print('⚠️ PyMuPDF 未安装，尝试自动安装...', file=sys.stderr)
    try:
        subprocess.run(
            [sys.executable, '-m', 'pip', 'install', 'PyMuPDF'],
            check=True, timeout=120, capture_output=True
        )
        fitz = importlib.import_module('fitz')
        print('✅ PyMuPDF 自动安装成功')
    except Exception as e:
        print(f'❌ PyMuPDF 安装失败: {e}', file=sys.stderr)
        print(f'请运行: {sys.executable} -m pip install PyMuPDF', file=sys.stderr)
        sys.exit(1)

sys.path.insert(0, str(Path(__file__).parent))

import parse_trip, parse_ticket, parse_didi, calc_subsidy, fill_trip_xlsx, fill_base_xlsx, adjust_trip
from parse_didi import parse_all_rides, classify_didi
from render_form_summary import DEFAULT_FORM_CONFIG, render_form_summary

SKILL_DIR = Path(__file__).parent.parent
CONFIG_PATH = SKILL_DIR / 'config.json'
PERSONAL_INFO = SKILL_DIR / 'personal_info.json'
REIMBURSEMENT_FORM = SKILL_DIR / 'reimbursement_form.json'

MONTH_NAMES = ['', '1月', '2月', '3月', '4月', '5月', '6月',
               '7月', '8月', '9月', '10月', '11月', '12月']

DEFAULT_CONFIG = {
    "version": 1,
    "pdf_parser": {
        "tool": "pymupdf",
        "auto_install": True
    },
    "image_parser": {
        "mcp_available": False,
        "mcp_tool_name": None,
        "fallback_to_manual": True
    }
}


def _parse_dir_year_month(dir_path, year, month):
    """从目录名提取年月。优先显式参数 > 目录名正则 > 当前日期兜底。"""
    if year and month:
        return year, month
    m = re.search(r'(\d{4})年(\d{1,2})月?出差', str(dir_path))
    if m:
        return int(m.group(1)), int(m.group(2))
    if year:
        return year, month or date_cls.today().month
    today = date_cls.today()
    return today.year, today.month


def _parse_date_from_filename(stem, year):
    """从文件名尝试提取日期。返回 date 或 None。"""
    m = re.search(r'(\d{1,2})月(\d{1,2})日', stem)
    if m:
        return date_cls(year, int(m.group(1)), int(m.group(2)))
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', stem)
    if m:
        return date_cls(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r'(\d{4})(\d{2})(\d{2})', stem)
    if m:
        return date_cls(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r'(\d{2})(\d{2})高铁', stem)
    if m:
        mo, dy = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12 and 1 <= dy <= 31:
            return date_cls(year, mo, dy)
    return None


def _interactive_fare_input(images, target_dict, year, item_name, mcp_avail, mcp_tool, interactive):
    """交互式输入图片中的费用（金额+日期），支持大巴/高铁图片。"""
    for img in images:
        if not interactive:
            print(f"  ⚠️ {item_name}图片跳过（非交互模式）: {img.name}")
            continue
        default_d = _parse_date_from_filename(img.stem, year)
        hint = default_d.strftime('%Y年%m月%d日') if default_d else '日期未知'
        if mcp_avail and mcp_tool:
            print(f"  🔍 IMAGE_MCP|{img.name}|tool={mcp_tool}|hint={hint}")
            print(f"     正在调用 {mcp_tool} 解析图片，提取费用明细(金额+日期)...")
        print(f"  📷 {img.name} ({hint}) — 输入0或回车结束")
        try:
            entry_no = 1
            while True:
                print(f"    第{entry_no}笔 金额: ", end='')
                amt_str = input().strip()
                if not amt_str or amt_str == '0':
                    break
                amt = float(amt_str)
                if amt <= 0:
                    break
                print(f"      日期 (MM/DD): ", end='')
                date_str = input().strip()
                m = re.match(r'(\d{1,2})[/\-](\d{1,2})', date_str)
                if m:
                    d = date_cls(year, int(m.group(1)), int(m.group(2)))
                    target_dict[d] = round(amt, 2)
                    print(f"      → {d}  {amt}元")
                    entry_no += 1
                else:
                    print(f"      ⚠️ 日期格式无效，跳过")
        except (EOFError, KeyboardInterrupt):
            pass


def _ensure_config():
    """Ensure config.json exists with defaults. Returns config dict."""
    if not CONFIG_PATH.exists():
        with open(CONFIG_PATH, 'w') as f:
            json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
        print(f"📝 已创建默认配置文件: {CONFIG_PATH}")
    with open(CONFIG_PATH) as f:
        return json.load(f)


def _load_config():
    """Load config.json, merge with defaults for any missing keys."""
    config = DEFAULT_CONFIG.copy()
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH) as f:
            existing = json.load(f)
        # Deep merge
        for k, v in existing.items():
            if k in config and isinstance(config[k], dict) and isinstance(v, dict):
                config[k].update(v)
            else:
                config[k] = v
    return config


def _load_form_config():
    """Load reimbursement_form.json, merge with defaults for missing keys."""
    config = dict(DEFAULT_FORM_CONFIG)
    if REIMBURSEMENT_FORM.exists():
        with open(REIMBURSEMENT_FORM) as f:
            existing = json.load(f)
        config.update(existing)
    return config


def init(interactive=True):
    """Skill initialization: ensure config, check dependencies, detect MCP.

    Args:
        interactive: 是否允许交互式选择执行模式。

    Returns:
        dict: Loaded config (merged with defaults).
    """
    config = _ensure_config()

    # ── PDF parser check ──
    parser_tool = config.get('pdf_parser', {}).get('tool', 'pymupdf')
    print(f'🔧 PDF解析工具: {parser_tool}')
    try:
        import fitz
        print(f'   ✅ PyMuPDF {fitz.__version__} 可用')
        print(f'   📍 {fitz.__file__}')
    except ImportError:
        print('   ❌ PyMuPDF 不可用')
        if config.get('pdf_parser', {}).get('auto_install'):
            print('   ↻ auto_install=true，运行时将自动安装')

    # ── Image MCP check ──
    mcp = config.get('image_parser', {})
    if mcp.get('mcp_available') and mcp.get('mcp_tool_name'):
        print(f'🔧 图片MCP工具: {mcp["mcp_tool_name"]} ✅')
    else:
        print('ℹ️  图片MCP工具: 未配置')
        print('   检测到图片时将回退到手动输入')
        print('   如需配置，编辑 config.json 的 image_parser 字段')
        print('   或联系管理员配置 MCP server 后在 config 中启用')

    # ── 执行模式选择 ──
    current_mode = config.get('execution_mode', 'manual')
    if interactive:
        print(f'\n📋 当前执行模式: {current_mode}')
        print()
        print('请选择执行模式:')
        print('  1. CLI 模式    — 一条命令带全部参数，零交互，适合自动化/Claude Code')
        print('  2. Manual 模式 — 逐步调用脚本、逐步确认，适合首次使用/不确定参数')
        try:
            ans = input(f'请输入 (1/2, 默认{"1" if current_mode == "cli" else "2"}): ').strip()
        except (EOFError, KeyboardInterrupt):
            ans = ''
        if ans == '1':
            config['execution_mode'] = 'cli'
        elif ans == '2':
            config['execution_mode'] = 'manual'
        else:
            config['execution_mode'] = current_mode  # 保持原值

        if config['execution_mode'] != current_mode:
            with open(CONFIG_PATH, 'w') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            print(f'   ✅ 执行模式已更新: {config["execution_mode"]}')
    else:
        print(f'\n📋 执行模式: {current_mode}')

    print(f'\n📁 技能目录: {SKILL_DIR}')
    print(f'📄 配置文件: {CONFIG_PATH}')
    return config


def generate(work_dir, year=None, month=None, pi=None, base_city=None,
             interactive=True, config=None, extra_bus=None, extra_train=None,
             overtime_dates=None, project_info=None, standard=None):
    """自动发现并处理一个出差目录，生成差旅费和Base地交通费Excel。

    Args:
        work_dir: 出差目录路径
        year: 年份，None则从目录名提取
        month: 月份，None则从目录名提取
        pi: 个人信息dict，None则从personal_info.json读取
        base_city: Base城市，None则从pi读取
        interactive: 是否交互模式（非交互模式跳过所有input()）
        config: 配置dict，None则从config.json加载
        extra_bus: list of (date, amount) 通过CLI传入的大巴费用
        extra_train: list of (date, amount) 通过CLI传入的高铁费用
        standard: 1=公司订酒店，2=自行解决住宿；None 时交互询问或默认 2

    Returns:
        dict{out_trip, out_base, form_markdown, form_data, trips, trips_data} 或 None（失败时）
    """
    work_dir = Path(work_dir)
    if not work_dir.exists():
        print(f"❌ 目录不存在: {work_dir}")
        return None

    # ── Config ──
    if config is None:
        config = _load_config()
    mcp_avail = config.get('image_parser', {}).get('mcp_available', False)
    mcp_tool = config.get('image_parser', {}).get('mcp_tool_name')

    # ── 年月 ──
    year, month = _parse_dir_year_month(work_dir, year, month)
    month_str = MONTH_NAMES[month] if 1 <= month <= 12 else f'{month}月'
    print(f"📅 {year}年{month}月 → {work_dir}")

    # ── 个人信息 ──
    if pi is None:
        with open(PERSONAL_INFO) as f:
            pi = json.load(f)
    if base_city is None:
        base_city = pi.get('base_city', '广州')

    name_str = pi['name']

    # ══ 1. 携程行程单 ══
    trip_pdfs = sorted(work_dir.glob('*携程商旅*.pdf'))
    print(f'发现 {len(trip_pdfs)} 张携程PDF')
    if not trip_pdfs:
        print('❌ 未找到携程行程单PDF，终止')
        return None
    trips = parse_trip.parse_all_ctrip([str(p) for p in trip_pdfs], default_year=year)
    print(f'携程区间: {len(trips)} 个')
    for t in trips:
        print(f'  {t["start"]}~{t["end"]} {t["from_city"]}->{t["to_city"]}')

    # ══ 2. 打车行程单（滴滴 + 通用平台 + 通行费匹配）══
    ride_patterns = ['*滴滴*.pdf', '*出行*.pdf', '*打车*.pdf', '*行程单*.pdf']
    ride_pdfs = set()
    for pat in ride_patterns:
        ride_pdfs.update(work_dir.glob(pat))
    # 加上通行费 PDF（parse_all_rides 会自动识别并匹配到打车行程）
    ride_pdfs.update(work_dir.glob('EX_ESAA*.pdf'))
    # 排除已识别的携程文件
    ride_pdfs -= set(trip_pdfs)
    ride_pdfs = sorted(ride_pdfs)

    all_didi = parse_all_rides([str(p) for p in ride_pdfs], default_year=year)

    # ══ 3. 分类（Base地 vs 出差地）══
    base_classified = classify_didi(all_didi, trips, base_city)
    base_amounts = base_classified.get('base', {})
    didi_trip_data = base_classified.get('trip', {})
    print(f'\nBase地打车: {sum(len(v) if isinstance(v, list) else 1 for v in base_amounts.values())} 笔')
    for d, val in sorted(base_amounts.items()):
        entries = val if isinstance(val, list) else [val]
        for info in entries:
            print(f'  {d} {info["from"]}→{info["to"]} {info["amount"]}')
    print(f'出差地打车: {len(didi_trip_data)} 笔')
    for d, info in sorted(didi_trip_data.items()):
        print(f'  {d} {info["from"]}→{info["to"]} {info["amount"]}')

    # ══ 4. 大巴票（PDF）══
    bus_amounts = {}
    rich_bus_fares = {}
    bus_pdfs = sorted(work_dir.glob('*大巴*.pdf'))
    for pdf in bus_pdfs:
        try:
            result = parse_ticket.parse_bus_ticket(pdf)
            if result['date'] and result['amount'] > 0:
                bus_amounts[result['date']] = round(result['amount'], 2)
                rich_bus_fares[result['date']] = result
                print(f"  大巴PDF: {result['date']} {result['amount']}元")
        except Exception as e:
            print(f"  ⚠️ 大巴PDF解析失败 {pdf.name}: {e}")

    # ══ 4a. 大巴票（图片）══
    bus_imgs = sorted(work_dir.glob('*大巴*.png')) + sorted(work_dir.glob('*大巴*.jpg'))
    _interactive_fare_input(bus_imgs, bus_amounts, year, '大巴', mcp_avail, mcp_tool, interactive)
    for d, amt in bus_amounts.items():
        if d not in rich_bus_fares:
            rich_bus_fares[d] = amt

    # ══ 4b. 高铁票（文件名匹配 + 内容检测）══
    train_fares = {}
    rich_train_fares = {}
    train_pdfs = sorted(work_dir.glob('*火车*.pdf')) + sorted(work_dir.glob('*高铁*.pdf')) + sorted(work_dir.glob('*12306*.pdf'))

    # ── 内容检测：扫描尚未归类的 PDF，识别铁路电子客票 ──
    already_categorized = set(trip_pdfs) | set(ride_pdfs) | set(bus_pdfs) | set(train_pdfs)
    all_pdfs = set(work_dir.glob('*.pdf'))
    uncertain_pdfs = all_pdfs - already_categorized
    for pdf in sorted(uncertain_pdfs):
        try:
            doc = fitz.open(str(pdf))
            text = doc[0].get_text()
            doc.close()
            if any(kw in text for kw in ['铁路电子客票', '中国铁路', '电子发票（铁路', '火车票']):
                train_pdfs.append(pdf)
                print(f"  🔍 内容识别为火车票: {pdf.name}")
        except Exception:
            pass

    # 去重
    train_pdfs = sorted(list(set(train_pdfs)))
    for pdf in train_pdfs:
        try:
            result = parse_ticket.parse_train_ticket(pdf)
            if result['date'] and result['amount'] > 0:
                train_fares[result['date']] = train_fares.get(result['date'], 0.0) + round(result['amount'], 2)
                rich_train_fares[result['date']] = result
                print(f"  火车PDF: {result['date']} {result['amount']}元")
        except Exception as e:
            print(f"  ⚠️ 火车PDF解析失败 {pdf.name}: {e}")

    train_imgs = sorted(work_dir.glob('*高铁*.png')) + sorted(work_dir.glob('*高铁*.jpg'))
    _interactive_fare_input(train_imgs, train_fares, year, '高铁', mcp_avail, mcp_tool, interactive)
    for d, amt in train_fares.items():
        if d not in rich_train_fares:
            rich_train_fares[d] = amt

    # ══ 4c. 补贴标准选择 ══
    if standard not in (1, 2, None):
        raise ValueError('standard 必须是 1 或 2')
    selected_standard = standard or 2
    if interactive and standard is None:
        print(f'\n📋 出差区间:')
        for i, t in enumerate(trips):
            print(f'  {i+1}. {t["start"]}~{t["end"]} {t["from_city"]}->{t["to_city"]} | {t.get("reason", "出差")}')
        print('\n差旅补贴标准:')
        print('  A. 标准一：公司订统一酒店（一类60/二类50/三类40元/天）')
        print('  B. 标准二：自行解决住宿（一类220/二类180/三类145元/天）')
        ans = input('请选择 (A/B, 默认B): ').strip().lower()
        selected_standard = 1 if ans == 'a' else 2
    else:
        print(f'\n📋 补贴标准: 标准{selected_standard}')
    for t in trips:
        t['standard'] = selected_standard

    # ══ 4c2. CLI传入的大巴/高铁费用（在中断检测前合并）══
    for d, amt in (extra_bus or []):
        bus_amounts[d] = round(amt, 2)
        rich_bus_fares[d] = round(amt, 2)
        print(f'  → CLI大巴 {d} {amt}元')
    for d, amt in (extra_train or []):
        train_fares[d] = round(amt, 2)
        rich_train_fares[d] = round(amt, 2)
        print(f'  → CLI高铁 {d} {amt}元')

    # ══ 4d. 中断返程检测 ══
    trips, questions = adjust_trip.process_trip_interruptions(
        trips, rich_train_fares, rich_bus_fares, didi_trip_data, base_amounts
    )
    if questions:
        print(f'\n⚠️ 需确认 {len(questions)} 个中断返程场景:')
        for q in questions:
            trip_str = f'{q["trip"]["start"]}~{q["trip"]["end"]}'
            print(f'\n📋 区间 {trip_str}')
            print(f'   在 {q["last_return"]} 有高铁/大巴费用，之后无更多费用。')
            if interactive:
                ans = input('   出差是否提前结束？(y/n, 默认y): ').strip().lower()
                extra = adjust_trip.apply_answer([q], [ans != 'n'])
                trips.extend(extra)
                print(f'   → 已调整，共 {len(trips)} 个子区间')
            else:
                # 非交互模式：默认提前结束（True），不保留延伸段
                answers = []
                for q in questions:
                    answers.append(True)  # True → apply_answer 提前结束（不生成延伸段）
                extra = adjust_trip.apply_answer(questions, answers)
                trips.extend(extra)
                print(f"   → 已处理非交互提前结束，共 {len(trips)} 个子区间")
    print(f'\n调整后区间数: {len(trips)}')
    for t in trips:
        print(f'  {t["start"]}~{t["end"]} {t["from_city"]}->{t["to_city"]}')

    # ══ 4e. 周末加班检测（仅当出差区间包含周六/日时才询问）══
    overtime_set = None
    if overtime_dates:
        overtime_set = set()
        for d_str in overtime_dates.split(','):
            d_str = d_str.strip()
            parts = d_str.split('-')
            overtime_set.add(date_cls(int(parts[0]), int(parts[1]), int(parts[2])))
        print(f'📅 加班日期: {sorted(overtime_set)}')
    elif calc_subsidy.has_weekends_in_trips(trips):
        candidate_dates = calc_subsidy.collect_overtime_dates(trips)
        if candidate_dates:
            print('ℹ️ 以下日期若有加班会影响补贴，如有请用 --overtime-dates 指定:')
            for d in sorted(candidate_dates):
                label = '六' if d.weekday() == 5 else '日' if d.weekday() == 6 else '工作日'
                print(f'   {d} ({label})')
            if interactive:
                reply = input('请输入加班日期（如 12/9,12/10；无则回车）: ').strip()
                if reply:
                    overtime_set = calc_subsidy.parse_overtime_reply(reply, year=year)
                    print(f'📅 加班日期: {sorted(overtime_set)}')

    # ══ 5. 补贴计算 ══
    trips_data_list = calc_subsidy.compute_daily_subsidy(trips, overtime_dates=overtime_set, year=year)
    trips_data = dict(trips_data_list)
    total_subsidy = sum(v['subsidy'] for _, v in trips_data_list)
    print(f'\n补贴合计: {total_subsidy}元')
    print(f'补贴明细 ({len(trips_data)} 天):')
    for d, info in sorted(trips_data.items()):
        print(f'  {d} {info["reason"][:20]}... {info["subsidy"]}元')

    # ══ P60：合并大巴到 days_data（fill_trip_xlsx 只读 days_data 中的 bus 字段）══
    for d, amt in bus_amounts.items():
        if d in trips_data:
            trips_data[d]['bus'] = amt
            print(f'  → 合并大巴 {d} {amt}元')

    # ══ 6. 生成文件 ══
    out_trip = work_dir / f'《{name_str}-{month_str}-个人费用报告》 - 差旅费.xlsx'
    out_base = work_dir / f'《{name_str}-{month_str}-个人费用报告》 - Base地交通费.xlsx'

    fill_trip_xlsx.fill_trip_xlsx(
        out_path=str(out_trip),
        days_data=trips_data,
        train_fares=train_fares,
        didi_trip=didi_trip_data,
        pi=pi,
        merged_trips=trips,
        project_info=project_info,
    )

    if base_amounts:
        fill_base_xlsx.fill_base_xlsx(
            out_path=str(out_base),
            base_amounts=base_amounts,
            pi=pi,
            project_info=project_info,
        )
        out_base_str = str(out_base)
        print(f'\n✅ 生成完成:')
        print(f'  {out_trip}')
        print(f'  {out_base}')
    else:
        out_base_str = None
        print(f'\n✅ 生成完成:')
        print(f'  {out_trip}')
        print(f'  ℹ️ 无Base地滴滴行程，跳过生成Base地交通费表格')

    form_markdown_path = work_dir / '报销网页填写清单.md'
    form_markdown, form_data = render_form_summary(
        form_config=_load_form_config(),
        year=year,
        month=month,
        out_trip=str(out_trip),
        out_base=out_base_str,
        days_data=trips_data,
        train_fares=train_fares,
        didi_trip=didi_trip_data,
        base_amounts=base_amounts,
    )
    form_markdown_path.write_text(form_markdown, encoding='utf-8')
    print(f'  {form_markdown_path}')

    return {
        'out_trip': str(out_trip),
        'out_base': out_base_str,
        'form_markdown': str(form_markdown_path),
        'form_data': form_data,
        'trips': trips,
        'trips_data': trips_data,
    }


def main():
    parser = argparse.ArgumentParser(description='差旅报销单生成工具')
    parser.add_argument('--dir', help='出差目录路径')
    parser.add_argument('--year', type=int, default=None, help='年份（默认从目录名提取）')
    parser.add_argument('--month', type=int, default=None, help='月份（默认从目录名提取）')
    parser.add_argument('--non-interactive', action='store_true', help='非交互模式（跳过所有input()）')
    parser.add_argument('--init', action='store_true', help='仅初始化配置+检查依赖，不生成报销单')
    parser.add_argument('--bus', action='append', help='大巴费 "YYYY-MM-DD,金额" 可重复', default=[])
    parser.add_argument('--train', action='append', help='高铁费 "YYYY-MM-DD,金额" 可重复', default=[])
    parser.add_argument('--overtime-dates', help='加班日期 "YYYY-MM-DD,YYYY-MM-DD"（多个用逗号分隔）')
    parser.add_argument('--standard', type=int, choices=[1, 2], help='补贴标准：1=公司订酒店，2=自行解决住宿')
    parser.add_argument('--project-code', help='项目编号/商机编号')
    parser.add_argument('--project-name', help='项目名称/商机名称')
    parser.add_argument('--project-status', help='项目状态')
    parser.add_argument('--project-manager', help='项目经理')
    args = parser.parse_args()

    if args.init:
        init(interactive=not args.non_interactive)
        return 0

    if not args.dir:
        parser.error('请指定 --dir（出差目录）或 --init（仅初始化）')

    def _parse_extra(v):
        parts = v.split(',')
        return (date_cls(*map(int, parts[0].split('-'))), float(parts[1]))

    extra_bus = [_parse_extra(v) for v in args.bus] if args.bus else None
    extra_train = [_parse_extra(v) for v in args.train] if args.train else None

    config = _load_config()
    # 个人信息
    with open(PERSONAL_INFO) as f:
        pi = json.load(f)
    # 项目信息（多级优先级：CLI 参数 > last_trip.json > personal_info.json）
    project_info = {}
    # 1. CLI 参数
    for k in ('project_code', 'project_name', 'project_status', 'project_manager'):
        val = getattr(args, k, None)
        if val:
            project_info[k] = val
    # 2. last_trip.json
    last_trip_path = SKILL_DIR / 'last_trip.json'
    if last_trip_path.exists() and not project_info.get('project_manager'):
        lt = json.loads(last_trip_path.read_text())
        key_map = {'project_id': 'project_code', 'project_name': 'project_name',
                   'project_status': 'project_status', 'project_manager': 'project_manager'}
        for old_k, new_k in key_map.items():
            val = lt.get(old_k) or lt.get(new_k)
            if val and new_k not in project_info:
                project_info[new_k] = val
    # 3. personal_info.json（兜底）
    for k in ('project_code', 'project_name', 'project_status', 'project_manager'):
        if k not in project_info:
            val = pi.get(k)
            if val:
                project_info[k] = val
    project_info = project_info or None

    result = generate(
        work_dir=args.dir,
        year=args.year,
        month=args.month,
        interactive=not args.non_interactive,
        config=config,
        extra_bus=extra_bus,
        extra_train=extra_train,
        overtime_dates=args.overtime_dates,
        project_info=project_info,
        standard=args.standard,
    )
    return 0 if result else 1


if __name__ == '__main__':
    sys.exit(main())
