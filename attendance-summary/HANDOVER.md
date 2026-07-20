# 考勤汇总脚本交接指南

## 1. 项目概述

本项目是一个基于 Python + FastAPI 的微服务，用于把门禁/打卡的原始 `.xls` 记录转换成两份 Excel：

1. **`汇总.xlsx`** —— 按员工统计当月工时、天数、异常情况的汇总表。
2. **`分发.xlsx`** —— 保留每个人每天上下班详细时间的明细表。

服务以 Docker 容器形式运行在 `8009` 端口，飞书网关通过 HTTP 调用上传原始记录并下载生成结果。

2026-07-10 起，8009 服务也支持车间考勤原始流水。HTTP 入口会读取文件内容识别类型：

- 车间流水：表头包含 `人员ID / 姓名 / 设备名称 / 方向 / 识别方式 / 时间`，走 `src/workshop_attendance.py`，输出 `N月奢匠车间考勤.xlsx`。
- 办公室原始记录：继续走 `src/generate_attendance.py`，输出办公室汇总表和分发表。

识别逻辑不依赖 `.xls` / `.xlsx` 扩展名。

---

## 2. 目录结构

```
D:\Services\robot-services\attendance-summary
├── src/
│   ├── attendance_summary_http.py      # HTTP 服务入口（FastAPI）
│   ├── generate_attendance.py          # 核心生成逻辑
│   ├── workshop_attendance.py          # 车间考勤生成逻辑
│   ├── parse_input.py                  # 解析原始 .xls 输入
│   ├── verify_input.py                 # 输入校验（可选）
│   └── verify_summary.py               # 汇总结果校验（可选）
├── src/templates/
│   ├── base/汇总.xlsx                   # 通用 31 天汇总模板
│   ├── base/分发.xlsx                   # 通用 31 天分发模板
│   ├── 3月/汇总.xlsx                    # 3 月份专用汇总模板
│   ├── 3月/分发.xlsx                    # 3 月份专用分发模板
│   └── 6月/...                         # 其他月份模板
├── deploy/docker/
│   ├── Dockerfile
│   └── docker-compose.yml              # 生产部署配置
├── config/requirements.txt             # Python 依赖
├── data/output/                        # 容器内输出目录
├── logs/                               # 运行日志
└── HANDOVER.md                         # 本文件
```

车间默认模板和样本人工调整配置：

```
src/templates_workshop/
├── workshop_template.xlsx
└── sample_adjustments_2026_01_06.json
```

---

## 3. 运行方式

### 3.1 本地命令行生成（调试/测试）

```bash
cd D:\Services\robot-services\attendance-summary

# 用 Docker 运行，避免本机缺少 Python
docker run --rm \
  -v "/d/wechat/.../2026年3月办公室原始记录.xls:/data/input.xls:ro" \
  -v "/d/output:/data/output" \
  attendance-summary-attendance-summary \
  python src/generate_attendance.py /data/input.xls /data/output
```

输出：
- `data/output/3月职能部门办公室考勤数据汇总(1).xlsx`
- `data/output/3月办公室考勤数据分发(1).xlsx`

### 3.2 HTTP 服务调用

```bash
curl -X POST "http://localhost:8009/process" \
  -F "file=@2026年3月办公室原始记录.xls" \
  -F "name=3月办公室考勤"
```

返回 JSON 中带有 `summary_url` 和 `distribution_url` 两个下载链接。

### 3.3 启动 / 重启服务

```bash
cd D:\Services\robot-services\attendance-summary\deploy\docker
docker compose up -d --build      # 首次或依赖变更后
docker compose restart            # 仅代码变更后重启
docker compose logs -f            # 查看日志
```

代码目录 `src/` 以只读方式挂载进容器，因此修改 `src/generate_attendance.py` 后只需 `restart`，不需要重建镜像。

---

## 4. 核心逻辑说明

### 4.1 输入解析

`parse_input.py` 读取原始 `.xls`：
- 跳过姓名、工号等表头行。
- 按姓名汇总每个人每天的全部刷卡时间。
- 对同一天多次刷卡去重并排序。

### 4.2 模板选择

`generate_attendance.py` 根据输入文件名中的月份关键字自动选模板：

| 文件名包含 | 使用模板 |
|---|---|
| `3月`、`03月` | `src/templates/3月/` |
| `6月`、`06月` | `src/templates/6月/` |
| 其他月份 | `src/templates/base/` |

### 4.3 两种生成策略

#### 策略 A：base 模板 —— 完全重建

没有月份专用模板时，从空白模板重新创建所有员工行，统一填 `五楼`、姓名、工号、每天时间。

#### 策略 B：月份专用模板 —— 保守更新

保留模板里原有的：
- 员工姓名、工号、部门/楼层
- 部门、行高、边框、字体等格式
- 已有的手动填写列（如汇总表 `S~W` 列）
- 原有的 VLOOKUP、汇总公式

只更新时间相关的单元格，并追加输入中有但模板里没有的新员工。

---

## 5. 关键：天数统计规则（3 月参考表规则）

这是 2026-07-07 根据 `3月考勤汇总.xlsx` 修正后的规则，只影响**汇总表**里的 `部门分发` sheet。

### 5.1 规则

| 当天刷卡次数 | 汇总表表现 | 主行天数 | 续行天数 |
|---|---|---|---|
| 1 次 | 一个单元格放时间，对侧为空 | 0.5 | — |
| 2 次 | 上班时间 / 下班时间 | 若 `上班 ≤ 10:00` 且 `下班 ≥ 14:00` 则为 1 天，否则 0.5 天 | — |
| 3 次 | 拆成两行：主行取前两次，续行取第三次 | 0.5 | 0.5 |
| 4 次及以上 | 拆成两行：主行取第 1、2 次，续行取第 3、4 次 | 0.5 | 0.5 |

### 5.2 代码位置

相关函数集中在 `src/generate_attendance.py`：

- `_compute_summary_row_days(all_times, is_extra)` —— 计算单行天数
- `_times_span_noon(times)` —— 判断是否跨中午（已不再使用，保留备用）
- `build_summary_distribution_sheet(...)` —— base 模板重建
- `update_summary_distribution_sheet(..., split_3plus_days)` —— 月份模板保守更新
- `append_summary_distribution_rows(..., split_3plus_days)` —— 追加新员工
- `add_summary_distribution_totals(...)` / `update_summary_formulas(...)` —— 3 月汇总表合计列与公式修复

`split_3plus_days=True` 仅在月份为 `3月` 时启用，其他月份保持原有保守更新行为。

### 5.3 时间单元格颜色标记

从 2026-07-14 起，汇总表和分发表的每日时间单元格会按以下规则上色：

| 颜色 | 含义 | 判断条件 |
|---|---|---|
| 红色 | 迟到 | 上班时间（当天第 1 个时间）> 08:30 |
| 黄色 | 早退 | 下班时间（当天最后 1 个时间）< 季节阈值（5-10 月 < 18:00，其它 < 17:40） |
| 绿色 | 仅 1 次刷卡 | 当天只有 1 个刷卡记录 |

说明：
- 汇总表 `部门分发` 中上班/下班是分开的单元格，可分别标红/黄。
- 分发表每天只有 1 个单元格，按整体状态标一种颜色；若同时迟到和早退，红色优先。
- 代码位置：`LATE_FILL` / `EARLY_FILL` / `SINGLE_FILL` 常量，以及 `_time_cell_fill` / `_dist_day_cell_fill` 函数。

---

## 6. 常见问题与调试

### 6.1 生成后打开汇总表显示 `#VALUE!`

- 检查 `部门分发` sheet 最后两列（DZ / EA）是否已有合计列。
- 检查 `汇总` sheet 的 H/I/J/R 列是否指向正确的合计列。
- 代码中 `add_summary_distribution_totals` 与 `update_summary_formulas` 负责这件事，只对 3 月生效。

### 6.2 总天数与参考表对不上

1. 确认输入数据是否一致（参考表可能是旧数据，当前输入人数可能更多）。
2. 用 `verify_summary.py` 或写脚本对比 `部门分发` sheet 的 `days` 列合计。
3. 检查该员工是否有 3 次以上刷卡但模板里没有续行。

### 6.3 新员工没有追加到正确位置

- 新员工统一追加在 `部门分发` sheet 末尾，楼层填 `五楼`。
- 若需要按部门分组，需手动调整或修改 `append_summary_distribution_rows` 逻辑。

### 6.4 新增月份模板

1. 在 `src/templates/` 下新建目录，如 `9月/`。
2. 放入 `汇总.xlsx` 和 `分发.xlsx` 两个模板。
3. 在 `generate_attendance.py` 的 `month_label_from_filename` 中增加识别规则（已有 `3月`、`6月` 示例）。
4. 若该月份需要特殊拆分规则，参考 `split_3plus_days` 的写法新增分支。

---

## 7. 修改历史

| 时间 | 修改内容 |
|---|---|
| 2026-07-07 | 修复 3 月汇总表天数统计：1 次 0.5 天；2 次按上班时间 ≤10:00 且下班时间 ≥14:00 判定 1 天；3 次及以上拆成两行各 0.5 天。6 月输出保持不变。 |
| 2026-07-07 | 修复 3 月汇总表 `#VALUE!` 错误：在 `部门分发` 增加 DZ/EA 辅助合计列，汇总表 H/I/J/R 列改用 SUMIF 引用。 |
| 2026-07-07 | 增加 3 月分发专用模板支持；修复分发 sheet 字体、对齐、自动换行。 |
| 2026-07-14 | 给汇总表、分发表的每日时间单元格增加颜色标记：迟到红色、早退黄色、仅 1 次刷卡绿色；并移除原来的「N日异常」文本列，状态完全用颜色表示。6 月天数仍沿用模板值。 |
| 2026-07-14 | 过滤并隐藏整月无打卡数据的员工：从汇总表、汇总部门分发、分发表中删除无数据行，并删除汇总表末尾遗留的空行。3 月汇总表公式在删除后重新写入 SUMIF，避免行号错位。 |

---

## 8. 关键命令速查

```bash
# 健康检查
curl http://localhost:8009/health

# 重启服务
cd D:\Services\robot-services\attendance-summary\deploy\docker
docker compose restart

# 查看日志
docker compose logs -f

# 进容器调试
docker exec -it attendance-summary-8009 bash
```

---

## 9. 联系与备注

- 当前服务端口：`8009`
- 容器名：`attendance-summary-8009`
- 主要维护文件：`src/generate_attendance.py`、`src/attendance_summary_http.py`
- 月份模板目录：`src/templates/`

如需进一步调整月份规则，优先在 `_compute_summary_row_days` 和 `update_summary_distribution_sheet` 中修改。
