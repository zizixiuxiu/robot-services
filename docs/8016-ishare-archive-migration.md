# 8016 iShare 归档同步 —— 迁移到服务器

> 日期：2026-11-09　结论：**已迁完，下车间单（8006）直接写共享盘，本机 8016 进程可停**

## 一、为什么要迁

8006 已于 2026-11-07 切到平台，但平台侧 `capability.py` 当时明确「不做 ishare 归档、不调 8016」，
导致 11-07 14:04 之后所有下车间单**只回传、不进共享盘**，本机备份同步停摆。

本轮把归档这一步直接放进平台能力里，由服务器自己写共享盘，本机不再参与。

## 二、最终链路

```
飞书群 → 网关容器（本机） → 适配器 8006（本机） → ssh 隧道 18191
      → 平台 worker-workshop-order（服务器）
          ├─ 转换 + 改动说明 → 回传网关
          └─ 归档 → /mnt/ishare/{年}/{月}/{姓名}/原文件|处理后
                        ↑ 宿主机 systemd 挂载 //192.168.8.135/Ishare录单临时文件
      → ishare-backup.timer（服务器，每 10 分钟）
          → /home/ops/ishare_backup/{年}/…（替代原 D:\Ishare_2026_backup）
```

## 三、服务器侧改动

| 项 | 内容 |
|---|---|
| 凭据 | `/etc/cifs-ishare.cred`（root 0600，`username=LD`） |
| 挂载单元 | `/etc/systemd/system/mnt-ishare.mount`，`vers=2.1`（3.0 报 95 不支持），已 `enable` |
| 挂载点 | `/mnt/ishare`（读写已实测） |
| worker 挂载 | `compose.workshop-order.yaml` 增加 `/mnt/ishare:/mnt/ishare:rslave`、`WORKSHOP_ARCHIVE_ROOT=/mnt/ishare` |
| 归档代码 | `services/workshop_order/capability.py` 新增 `_archive_pair()`，逐字对齐生产 `archive_workshop_pair`（含 `(1)` 重传后缀归一、同名旧副本清理） |
| 版本 | capability 0.2.0 → **0.3.0**，`side_effects` 改为 `writes_ishare_share` |
| 备份同步 | `/home/ops/ishare_backup/sync_backup.py` + `ishare-backup.timer`（10 分钟，只增量、从不删除） |
| 备份目录 | `/home/ops/ishare_backup/2026`（本机历史备份已整包导入） |

适配器侧同步改动：`cutover_generic.py` 把网关 payload 里的 `uploader_name` 透传给平台 params
（仅 `workshop-order` 生效，避免污染其他服务的 params）。

## 四、实测（2026-11-09 10:55）

用 11-07 真实单 `S2611-1505郝丹丹.xls`，`uploader_name=__smoke__`（不污染真实目录）：

| 项 | 结果 |
|---|---|
| 往返耗时 | 17.5 s（生产 8006 同单约 60–90 s） |
| 归档落盘 | `/mnt/ishare/2026/11/__smoke__/原文件/S2611-1505郝丹丹.xls`（371,200 B）<br>`…/处理后/S2611-1505郝丹丹下车间.xlsx`（295,723 B） |
| 逐格比对 | 与生产基线 12 个 sheet **全部 0 差异**（值/格式/字体/边框/对齐） |
| 测试清理 | `__smoke__` 目录已删除，共享盘未留痕 |

## 五、回滚

1. `sudo systemctl disable --now mnt-ishare.mount ishare-backup.timer`
2. `compose.workshop-order.yaml` 里删掉 worker 的 `/mnt/ishare` 挂载与 `WORKSHOP_ARCHIVE_ROOT`
3. `capability.py` 恢复 `capability.py.bak-ishare`（compose 备份 `compose.workshop-order.yaml.bak-ishare`）
4. 重启 wso 三件套；本机重新启用 service-hub 的「iShare归档同步」（`config/services.json` 该条 `enabled: true`）

## 六、注意

- `vers=3.0` 会 `mount error(95)`，必须用 **2.1**；改挂载参数时同步改 `mnt-ishare.mount`。
- 挂载失效时归档不会阻断回传：任务照常成功，`summary.json` 的 `archive.status` 为 `error/disabled`，
  `metrics.archive_status` 也能看到，便于告警。
- 共享盘基址下还有「已录」「未录」「明细表」等别人的目录，备份同步只按顶层年份夹走，不会误搬。
