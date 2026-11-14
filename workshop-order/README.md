# ⚠️ 此目录不是部署源（已废弃）

`D:\Services\robot-services\workshop-order\src\make_workshop_order.py` 是**旧版单文件实现**，
2026-11 起线上已改为平台化结构，这个目录**改了对生产没有任何影响**。

## 线上真正在跑的位置

| 项 | 位置 |
|---|---|
| 源码（唯一源） | `GitHub → zizixiuxiu/robot-platform → services/workshop_order/workshop_core.py` |
| 服务器路径 | `/home/ops/robot-platform-tests/platform/services/workshop_order/` |
| 容器 | `robot-platform-worker-workshop-order-1` / `-2-1` / `-3-1` / `-4-1` |
| 挂载方式 | 源码 bind mount 到 `/app/services/workshop_order`，**改完必须重启 worker** |

## 改代码的正确流程

```bash
ssh -p 2222 -i ~/.ssh/id_ed25519 ops@192.168.6.245
cd /home/ops/robot-platform-tests/platform
# 改 services/workshop_order/workshop_core.py
git add -A && git commit -m "..." && git push
for c in $(sudo docker ps --format '{{.Names}}' | grep workshop-order); do sudo docker restart $c; done
```

## 为什么这里还留着

历史版本与变更记录（见仓库根 `SERVICES.md` 的变更记录表，2026-09 ~ 10 月多数条目出自这棵树）。
需要查旧实现细节时可以对照，但**不要在这里改功能**。
