---
name: travel-expense-fill-v3
description: 根据携程商旅 PDF、打车行程单、火车票/大巴票生成差旅报销 Excel。触发词：填写差旅报销、travel-expense-fill-v3、报销填写、生成差旅费报告、个人费用报告。
trigger: "携程PDF+报销|填写差旅报销|travel-expense-fill-v3|生成差旅费报告|差旅费报销|报销填写|个人费用报告"
validator:
  - 必须处理至少一个携程PDF
  - 必须输出两个Excel（差旅费+Base地交通费，除非Base地无数据）
  - 必须使用scripts/目录模块，不得手写解析逻辑
  - manual 模式不得跳步或手写 openpyxl
  - cli 模式 agentic 环境下须用 --non-interactive + --standard + 预收集数据
  - 禁止 restore_template_styles()、row_dimensions.hidden
  - 合计行必须 visible，多余行用 delete_rows 物理删除
---

# travel-expense-fill-v3

## 使用原则

- 必须至少有一个携程商旅 PDF。
- 优先使用 `scripts/` 模块和模板文件，不要手写解析逻辑或直接改 Excel。
- 代理环境优先走 CLI 非交互模式；先向用户收集项目、补贴标准、加班日期、图片票据金额等无法从 PDF 稳定判断的信息。
- `config.json` 是运行配置和操作提示；`execution_mode` 不会自动切换脚本行为。
- 图片票据没有可用 MCP 时，非交互 CLI 会跳过，需要提前把金额通过参数传入或在手动流程中向用户确认。

## 脚本说明

| 脚本 | 作用 |
|------|------|
| `scripts/parse_trip.py` | 携程 PDF 解析 + 多区间合并 |
| `scripts/parse_didi.py` | 打车行程单解析（Y 坐标分组法） |
| `scripts/parse_ticket.py` | 火车票/大巴票 PDF 解析 |
| `scripts/calc_subsidy.py` | 补贴计算 + 加班日期解析 |
| `scripts/adjust_trip.py` | 出差区间中断返程检测与拆分 |
| `scripts/fill_trip_xlsx.py` | 差旅费 Excel 填写 |
| `scripts/fill_base_xlsx.py` | Base地交通费 Excel 填写 |
| `scripts/render_form_summary.py` | 网页报销必填字段 Markdown 清单 |
| `scripts/gen_trip_reports.py` | CLI 一键生成入口 |

## 补贴规则速查

| | 标准一（公司订酒店） | 标准二（自行解决住宿） |
|---|---|---|
| **一类城市** | 60元/天 | 220元/天 |
| **二类城市** | 50元/天 | 180元/天 |
| **三类城市** | 40元/天 | 145元/天 |
| **出发/结束日** | 标准 × 50% | 标准 × 50% |
| **中间周末（无加班）** | 标准 × 50% | 标准 × 50% |
| **中间周末（有加班）** | 标准 × 100% | 标准 × 100% |

> **一类**：北京/上海/广州/深圳/香港
> **二类**：澳门/厦门/珠海/天津/重庆/青岛/苏州/武汉/成都/杭州/南京/西安/郑州/长沙/济南/福州/合肥/昆明/沈阳/哈尔滨/长春/南昌/太原/石家庄/贵阳/南宁/兰州/海口/银川/西宁/拉萨/乌鲁木齐/呼和浩特
> 其余 = 三类

## 快速流程

### Step 0 初始化（首次使用）

1. 运行 `gen_trip_reports.py --init` 检查依赖
2. 选择执行模式（init 会交互询问）：

| 模式 | 优点 | 缺点 | 适用场景 |
|------|------|------|----------|
| **CLI 模式** | 一键生成、零交互、适合自动化 | 需提前收集所有参数 | Claude Code、重复报销 |
| **Manual 模式** | 逐步确认、灵活调整 | 需要多步交互 | 首次使用、不确定参数 |

→ 选择后自动写入 `config.json` 的 `execution_mode`。

3. 确认 `personal_info.json` 包含 `name, emp_id, manager, base_city`
4. 确认 `last_trip.json` 包含 `project_code, project_name, project_status, project_manager`
   > 项目信息可填 `"无"`，不影响报销单生成。
5. 确认 `reimbursement_form.json` 包含网页报销固定必填项：所属部门、标题、报销主体、币种、原借款、业务编号。

### Step 1 预收集信息

每次生成前先确认：

- 是否延用 `last_trip.json` 的项目信息；项目变化时只更新 `last_trip.json`，不要写入 `personal_info.json`。
- 补贴标准：`1` = 公司统一订酒店；`2` = 自行解决住宿。
- 加班日期：仅当出差区间包含周六/日时才需要确认；周末有加班按 100% 补贴，格式用 `YYYY-MM-DD,YYYY-MM-DD`。
- 图片票据或解析失败票据的日期和金额。

### Step 2 生成报销表和网页填写清单

CLI 参数优先，适合代理非交互执行：

```bash
python3 scripts/gen_trip_reports.py \
  --dir /path/to/出差目录 \
  --year 2025 --month 12 \
  --non-interactive \
  --standard 1 \
  --overtime-dates 2025-12-09,2025-12-10 \
  --project-code "无" --project-name "无" --project-status "售前" --project-manager "无"
```

可选票据补录：

```bash
--bus "2025-12-24,45.00" --train "2025-12-24,74.50"
```

生成完成后会同时输出：

- 差旅费 Excel
- Base地交通费 Excel（无 Base 地数据时跳过）
- `报销网页填写清单.md`：只包含网页截图中带 `*` 的必填字段，附件只列生成的 Excel 文件。

### Step 3 手动流程（需要逐步确认时）

按以下顺序调用脚本模块，不要跳步：

1. 携程解析  
`parse_trip.parse_all_ctrip(pdf_paths: list[str], default_year=int) → list[dict]`

2. 火车票/大巴票  
`parse_ticket.parse_train_ticket(path)`、`parse_ticket.parse_bus_ticket(path)`；仅保留日期在出差区间内的票据。

3. 补贴标准  
把用户确认的 `standard`（1 或 2）写回每个 merged 区间；没有明确答复时默认标准二。

4. 打车解析与分类  
`parse_didi.parse_all_rides(pdf_paths: list[str], default_year=int) → list[dict]`
`result = parse_didi.classify_didi(didi_list, merged, base_city) → {trip, base, unknown}`

> 支持滴滴、阳光出行、呼我出行（百度打车）等平台。高速通行费电子行程单自动匹配到对应打车行程。

5. 中断返程检测  
`adjust_trip.process_trip_interruptions(merged, train_fares, bus_fares, didi_trip, base_didi) → (adjusted, questions)`；拆分后各子区间独立计算补贴。若用户确认提前结束，保留 `q['sub_trips_built']`；若否，补上 `last_return + 1` 到原结束日的区间。

6. 加班确认与补贴计算  
`calc_subsidy.has_weekends_in_trips(merged)` 判断是否涉及周末；仅当有周末时调用 `collect_overtime_dates(merged)` 收集候选日期。用 `parse_overtime_reply()` 解析用户回复，再调用 `compute_daily_subsidy()`。

7. 填写 Excel  
`fill_trip_xlsx.fill_trip_xlsx(...)` 生成差旅费表；仅当 `base_amounts` 有数据时调用 `fill_base_xlsx.fill_base_xlsx(...)`。

8. 生成网页报销填写清单  
调用 `render_form_summary.render_form_summary(...)` 生成 `报销网页填写清单.md`；金额口径为差旅费总额 + Base地交通费总额，不从 Excel 公式反读。

## 验证检查项

### 差旅费
- [ ] C2=姓名 J2=工号 J3=汇报对象 C4=项目编号 C5=项目名称
- [ ] J4=项目状态（默认"售前"） J5=项目经理（有则填姓名，无则"无"）
- [ ] B/C=出发地 D=出差地 E=出差原因（非写死） M=补贴标准备注
- [ ] 合计行 SUM 正确 | Row8-9 表头不变 | 多余行 delete_rows 物理删除

### Base地交通费
- [ ] A=日期 B=出发地 C=目的地 D=事由 F=费用 | 合计行 F=SUM

### 通用
- [ ] 确认目标 sheet 和模板行列没有错位
- [ ] 合并格按脚本当前策略处理，写值只写左上角锚点

### 网页填写清单
- [ ] `报销网页填写清单.md` 已生成
- [ ] 只包含带 `*` 的必填字段：所属部门、标题、报销主体、报销日期、摘要、报销合计、币种、原借款、应付余额、业务编号、附件
- [ ] 报销合计 = 差旅费总额 + Base地交通费总额；原借款默认 0；应付余额等于报销合计
- [ ] 附件只列生成的 Excel 文件

## 注意事项

1. **打车解析**：优先用滴滴 Y 坐标分组法，失败降级通用解析器；通行费按时间自动匹配
2. **Base地交通费**：同日多笔行程会独立成行（如早晚各一趟往返高铁站）
3. **火车票/大巴票**：只填入出差区间内的票据
4. **多区间合并**：间隔 ≤1 天视为连续区间
5. **日期写入**：用 `datetime` 对象（不要用 `date` 对象）
6. **表头不修改**：Row 1-9 保持模板原样

## 参考资料

| 文件 | 内容 |
|------|------|
| `references/column-mappings.md` | 差旅费/Base地模板列映射 |
| `references/iteration-log.md` | v1迭代记录、删除的无效代码、流程变更 |
| `references/benchmark-iteration-1.md` | v1方向A评测方法 + Iteration 1实测结果 |
| `references/bugfix_list.md` | v3已知bug修复记录 |
