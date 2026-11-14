# 8003 销售部业绩核对 — 平台接入就绪说明（未切流）

> 日期：2026-11-09
> 结论：平台侧 `dealer-sales.v1`（127.0.0.1:18090）**已通过影子对比，随时可切**，
> 但按业务方要求「不急用、脚本还在改进」**暂不切流**，生产 8003 继续由本机容器承担。

## 一、现状

| 项 | 生产 | 平台 |
|---|---|---|
| 入口 | `dealer-sales-8003`（may-sales 镜像）→ 127.0.0.1:8003 | `robot-platform-dsales-api-1` → 127.0.0.1:18090 |
| 代码 | `may-sales/src/*`（2026-08-13 最后改动） | `platform/services/dealer_sales/`，**md5 与生产逐字一致** |
| 运行时 | python 3.13.14 + node v20.20.2 + exceljs 4.4.0 | 同源（node_modules 直接取自制生产镜像） |
| 销量累计表 | `data/sales_state/销量累计_2026.xlsx` 73,413 B（9-10 12:02 写） | 持久卷 `robot-platform_dsales_artifacts:/data/sales_state/`，已导入同一份（sha256 `41046c0c…`） |
| 切流开关 | — | `cutover_service.ps1` 8003 行 `Enabled=$false`（不启适配器、不开隧道） |

## 二、影子测试（2026-11-09 10:15–10:17）

输入：平台 `samples/dealer-sales/real_5yue/` 真实三件套
（`综合查询5月(1).xls` 328 KB、`联思系统5月(1).xlsx` 90 KB、`奢匠下单统计5月.xlsx` 5.1 MB）

链路：临时适配器（`PLATFORM_MULTI=bundle`）→ 打 zip 单任务 → 平台 capability → 3 个产物回传

结果：

| 项 | 结果 |
|---|---|
| 任务 | success=True，25.8 s（生产同批 25–45 s） |
| 产物 | 3 个：`2026年5月销售部业绩核对表.xlsx`(23,766 B)、`2026年奢匠各经销商销量情况（5月）.xlsx`(331,055 B)、`汇总表_提取结果_5月.xlsx`(184,789 B) |
| 主表逐格 | 两个 sheet、各 116 行，**0 个单元格差异**（与生产 5 月基线 `ref_main.xlsx` 对照） |
| warning 透传 | 一致：`⚠️ 奢匠明细文件缺少「直营店+电商下单表」Sheet…` |

产物与对比 JSON 留档：`C:\Users\Administrator\WorkBuddy\2026-11-06-17-13-31\tmp_8003_ab\`

## 三、切流步骤

工具：`C:\Users\Administrator\WorkBuddy\2026-11-06-17-13-31\dealer_sales_cutover.ps1`

```powershell
# 1) 业务脚本改完后先同步（避免平台版本落后）
pwsh -File dealer_sales_cutover.ps1 -SyncCode

# 2) 一口气切：停生产容器 → 导最新累积表 → worker 开持久累积 → 适配器接管 8003
pwsh -File dealer_sales_cutover.ps1 -Enable

# 3) 抽查
pwsh -File dealer_sales_cutover.ps1 -Status
```

回滚（一条命令，10 秒级）：

```powershell
pwsh -File dealer_sales_cutover.ps1 -Disable
```

## 四、四个必须知道的坑

1. **累积表会被写回**。平台处于持久模式（`DEALER_SALES_PERSIST_STATE=1`）时，任何一次任务都会按当月数据回写
   累积表 —— 实测拿 5 月数据跑了一次，73,413 B 变成 72,830 B。所以：
   - 非切流期一律保持 `0`（隔离试点），跑测试不碰真值；
   - 每次 `-Enable` 会自动重新导入生产最新累积表；
   - 切流期间平台写过的表要作废，回滚后下次切之前必须重新 `-SeedState`。
2. **端口顺序**。适配器要监听 8003，生产容器必须先让出端口，否则适配器 bind 失败
   （实测 `WinError 10013`），而计划任务里它只是静默退出。`-Enable` 已把「停容器」做在「重启适配器」之前。
3. **网关不用改**。`CHAT_ROUTES` 里 8003 保持不变，网关看到的始终是 `host.docker.internal:8003`，
   切流只是把 8003 后面的实现换成适配器 → 平台。这也意味着网关层不需要发版。
4. **平台只有一个 role**。`dealer-sales.v1` 的输入是单 role `files`，三个文件必须打成一个 zip
   提交（成员保留原名），否则 capability 拿到单文件会走生产「需要至少 3 个文件」的错误路径。
   适配器的 `PLATFORM_MULTI=bundle` 就是干这个的。

## 五、待办

- [ ] 业务脚本稳定后：`-SyncCode` → 真实月份跑一次 A/B → `-Enable`
- [ ] 切流后观察首个真实月结（含累积表跨月递增是否符合生产口径）
- [ ] 平台 `docs/dealer-sales-迁移清单.md` 补一节「持久累积与切流」（平台侧 NOTICE 已同步）

---

## 六、2026-11-09 10:30 已切流（生产容器已停）

业务确认「可以切过去了，以后改完代码直接同步服务器」，因此不再等脚本稳定，直接切流。

**执行结果**

| 项 | 值 |
|---|---|
| 生产容器 | `dealer-sales-8003` 已 `docker stop`（Exited 137，`docker start` 可回滚） |
| 适配器 | `cutover_service.ps1` 中 8003 行 `Enabled=$true`，计划任务已重启 |
| 端口归属 | `curl 127.0.0.1:8003/health` → `{"workflow":"dealer-sales.v1","platform":{...ready}}` |
| 隧道 | 18192 → 平台 127.0.0.1:18090 |
| worker | `DEALER_SALES_PERSIST_STATE=1`（持久累积，卷内表 73,413 B / sha256 41046c0c…） |
| 代码 | 生产与平台 6 个业务文件 md5 全一致，未发生版本漂移 |

**端到端验证（走正式链路 8003 → 18192 → 18090）**

- 输入：真实 5 月三件套（综合查询 328 KB / 联思 88 KB / 奢匠 5.0 MB）
- 结果：`success=True`，26.5 s，3 个产物（23,766 / 331,055 / 184,789 B，与上轮平台产物字节大小完全一致）
- warning 与生产一致：`⚠️ 奢匠明细文件缺少「直营店+电商下单表」Sheet…`
- 主表逐格：`Sheet1 (2)` 116 行 0 差异、`Sheet1` 116 行 0 差异，**总差异 0**
- 测试后已把累积表恢复为生产真值（`73413 B / 41046c0c…`），避免 5 月数据污染 9 月状态

**此后改代码的流程（一条命令）**

```powershell
pwsh -File dealer_sales_cutover.ps1 -SyncCode     # 推业务脚本 + 模板，自动重启 worker
pwsh -File dealer_sales_cutover.ps1 -Status       # 看两边 md5 是否对齐
```

注意：`-SyncCode` 只同步业务文件（`dealer_sales_http.py`、两个 js、`package.json`、两个模板），
**不会碰** 平台侧 `capability.py` / `registry.json`，所以平台改造不会被覆盖。

**本轮踩到并修掉的脚本问题**

1. `-Enable`/`-Disable` 里停/起容器写成了 `ssh sudo docker …`，但生产 8003 容器其实在**本机**，
   结果永远查不到状态、也不会真停 → 已改为走本机 `docker.exe`（`Invoke-LocalDocker`）。
2. PowerShell 宿主里 `$env:HOME` 为空 → `ssh/scp` 找不到 `~/.ssh`，BatchMode 直接失败，
   而脚本只看到 "连不上" → 已在脚本顶部兜底 `HOME=$env:USERPROFILE`。
3. 隧道端口冲突：上一轮临时隧道残留进程占着 18192，新 ssh 建不上该转发（`Address already in use`），
   但 `ExitOnForwardFailure` 下 ssh 进程并未退出，表现为「适配器起来了、8003 却 502」。
   重启用 `Stop-ScheduledTask` + 杀掉 `-L` 的 ssh 进程再 `Start-ScheduledTask` 即可。

## 七、回滚

```powershell
pwsh -File dealer_sales_cutover.ps1 -Disable
```

约 10 秒回到本机容器（`service=may-sales`）。回滚后平台累积表作废，下次切之前必须重新 `-SeedState`。

## 八、待办（更新）

- [x] 切流（2026-11-09 完成）
- [ ] 观察首个真实月结：确认累积表跨月递增、三件套缺项文案、耗时（生产 5–10 min 的月份请留意）
- [ ] 若月结耗时明显变长，检查平台 worker 的 node 内存限制（`--max-old-space-size`）
