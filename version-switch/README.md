# version-switch — 新旧版本一键切换

## 原理（一句话）

8 个微服务的 HTTP 层刚从手写 `http.server` 重构到公共骨架 `factory_common`。
compose 中各服务的 `src/` 目录是**只读挂载**进容器的，所以
**切换版本 = 替换宿主机上各服务 `src/` 下的 `_http.py` 文件 + 重启容器**，无需重新 build。

## 目录结构

```
version-switch/
├── README.md          # 本文件
├── old/               # 8 个旧版 _http.py（迁移前的手写 HTTP 层）
├── new/               # 8 个新版 _http.py（factory_common 骨架，当前 src/ 的保险副本）
├── backup/            # 每次切换前自动备份当前文件（带时间戳）
└── switch.ps1         # 切换脚本
```

## 常用命令

在 Git Bash 中（或直接 powershell）：

```bash
cd /d/Services/robot-services/version-switch

# 全部 8 个服务切回旧版
powershell.exe -File switch.ps1 -Service all -Target old

# 全部切回新版
powershell.exe -File switch.ps1 -Service all -Target new

# 单个服务切回旧版（服务名见下表）
powershell.exe -File switch.ps1 -Service csv-board -Target old

# 单个服务切回新版
powershell.exe -File switch.ps1 -Service csv-board -Target new
```

服务名清单：`hardware-summary`、`order-split`、`may-sales`、`csv-board`、
`pvc-classify`、`workshop-order`、`quote-maker`、`dealer-report`

脚本行为：参数校验 → 备份当前文件到 `backup/` → 复制目标版本到 src/ →
`docker restart` → 轮询 `http://localhost:<port>/health`（最多 60 秒）→ 输出结果表。

## 注意事项

1. **切换只影响 HTTP 层**，各服务的业务代码（生成逻辑）新旧完全一致，输出结果不受影响。
2. 新版镜像里挂载的 `factory_common` 目录对旧版代码**无害**（旧代码不 import 它），已验证兼容。
3. 切换后建议**跑一单真实业务**验证，确认出文件正常。
4. 目前只有 `src/<svc>_http.py` 有新旧之分。如果未来业务脚本（如 `generate_*.js`）也发生改动，
   回退时需要连同业务脚本一起恢复——**目前不需要**。
5. 每次切换会自动把被覆盖的文件备份到 `version-switch/backup/`，误操作可手动找回。
6. `may-sales` 的容器名是 `dealer-sales-8003`（历史遗留），脚本内部已处理映射。
