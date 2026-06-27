# 剩余服务 Docker 迁移计划

> 创建时间：2026-06-27
> 当前状态：8001-8006 已 Docker 化并健康运行
> 计划执行时间：中午

---

## 一、剩余未迁移服务

| 序号 | 服务 | 端口/类型 | 当前路径 | 优先级 | 复杂度 |
|------|------|-----------|----------|--------|--------|
| 1 | simple-ims 平贴系统 | 8090 HTTP | `C:\Users\Administrator\Documents\Codex\2026-06-04\github\simple-ims` | 高 | 中 |
| 2 | tcp_proxy_8082 | 8082 TCP | `C:\tcp_proxy_8082.py` | 中 | 低 |
| 3 | feishu_bot_ws 飞书网关 | WS 网关 | `D:\1\feishu_bot_ws.py` | 中 | 高 |
| 4 | OrderFlowMonitor | 监控脚本 | `C:\Users\Administrator\Documents\Codex\2026-05-28\sqlserver\monitor_refresh_order_flow.py` | 低 | 低 |

---

## 二、迁移目标

1. 所有 HTTP/TCP 服务全部 Docker 化
2. 去除对 Windows 绝对路径和特定 Python 环境的依赖
3. 统一返回格式 `{filename: [{filename, file_content}]}`
4. 保留旧启动脚本兼容（改为指向 Docker）
5. 更新 `.kimi/skills/check-robot-services/SKILL.md`

---

## 三、分服务方案

### 3.1 8090 simple-ims 平贴系统

**当前状态**：uvicorn + FastAPI，路径 `C:\Users\Administrator\Documents\Codex\2026-06-04\github\simple-ims`

**迁移步骤**：
1. 在 `D:\Services\robot-services\simple-ims` 建立标准目录
2. 复制 `simple-ims` 源码到 `src/`
3. 检查依赖，写入 `config/requirements.txt`
4. 编写 Dockerfile（基于 python:3.12-slim）
5. 编写 docker-compose.yml，映射 `0.0.0.0:8090:8090`
6. 修改旧 VBS `D:\1\start_simple_ims_8090.vbs` 指向 Docker 启动脚本
7. 停止旧 Windows 进程（PID 5432），启动 Docker 容器
8. 测试 `/health` 和关键接口

**风险点**：
- 源码可能有 Windows 路径硬编码
- 可能有本地文件/数据库依赖

---

### 3.2 8082 tcp_proxy_8082

**当前状态**：单一 Python 文件 `C:\tcp_proxy_8082.py`

**迁移步骤**：
1. 在 `D:\Services\robot-services\tcp-proxy` 建立标准目录
2. 复制 `C:\tcp_proxy_8082.py` 到 `src/tcp_proxy_8082.py`
3. 去除 Windows 路径依赖
4. 编写 Dockerfile + docker-compose.yml，映射 `0.0.0.0:8082:8082`
5. 修改旧 VBS `D:\1\start_tcp_proxy_8082.vbs`
6. 停止旧进程（PID 1772），启动 Docker 容器
7. 测试端口连通性

**风险点**：
- 需要明确该代理转发到哪里（Dolibarr? ERP?）
- 可能涉及内网地址

---

### 3.3 feishu_bot_ws 飞书网关

**当前状态**：`D:\1\feishu_bot_ws.py`，所有 Docker 服务的入口

**迁移策略**：建议最后迁

**迁移步骤**：
1. 在 `D:\Services\robot-services\feishu-ws-gateway` 建立目录
2. 复制并改造 `feishu_bot_ws.py`：
   - 去除 Windows 路径
   - 服务地址改为 `localhost:8001-8006`
   - 日志输出到标准路径
3. 编写 Dockerfile + docker-compose.yml
4. 飞书应用后台配置 WebSocket 地址指向新端口/路径
5. 修改旧 VBS `D:\1\start_feishu_ws_hidden.vbs`
6. 停止旧网关进程，启动 Docker 容器
7. 完整测试：飞书机器人 → 网关 → 各后端服务

**风险点**：
- 迁移期间飞书机器人会中断
- 需要飞书后台配合改地址
- 如果网关不在 Docker 中也能稳定运行，可考虑不迁，只做标准化

---

### 3.4 OrderFlowMonitor

**当前状态**：定时监控脚本，`--interval 900`

**建议**：不强制 Docker 化

**原因**：
- 不是 HTTP 服务
- 需要连接 SQL Server
- Windows 计划任务运行即可

**可选优化**：
- 创建 `scripts/run_order_flow_monitor.ps1`
- 用 Windows 计划任务每 15 分钟执行一次
- 日志输出到统一目录

---

## 四、执行顺序

```
中午执行：
  1. 8090 simple-ims（优先级最高）
  2. 8082 tcp_proxy（确认用途后继续）
  3. feishu_ws_gateway（如果中午时间够且允许短暂中断）
  4. OrderFlowMonitor（可选，改计划任务）
```

---

## 五、回滚方案

每个服务迁移前：
1. 记录旧进程 PID
2. 保留旧文件和 VBS 启动脚本（不删除）
3. 如果 Docker 启动失败，立即停止容器，手动启动旧进程

---

## 六、验证清单

每个服务迁移后检查：
- [ ] `docker ps` 显示端口映射为 `0.0.0.0:<port>-><port>/tcp`
- [ ] `curl http://localhost:<port>/health` 返回 200
- [ ] 关键业务接口 `/process` 或对应接口测试成功
- [ ] 旧 Windows 进程未占用端口
- [ ] 旧 VBS 已改为指向 Docker 启动脚本
- [ ] GitHub 已推送
- [ ] `.kimi/skills/check-robot-services/SKILL.md` 已更新

---

## 七、预计耗时

| 服务 | 预计耗时 |
|------|----------|
| 8090 simple-ims | 30-45 分钟 |
| 8082 tcp_proxy | 15-20 分钟 |
| feishu_ws_gateway | 40-60 分钟 |
| OrderFlowMonitor | 10-15 分钟 |
| **总计** | **约 1.5-2.5 小时** |

---

## 八、注意事项

1. Docker Desktop WSL2 端口映射可能丢失，如果某个服务 healthy 但端口不通，用 `health_check_and_restart.ps1` 自动修复
2. 飞书网关迁移会导致机器人短暂不可用，建议避开业务高峰期
3. 所有新服务统一使用 `python:3.12-slim` 镜像
4. 输出格式统一为 `{filename: [{filename, file_content}]}`
