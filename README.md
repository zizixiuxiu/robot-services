# 飞书机器人 Docker 服务

克隆后装 Docker Desktop，填密钥，再按目录 `docker compose up -d` 即可。网关必须最后启动。

## 1. 克隆

```powershell
git clone git@github.com:zizixiuxiu/robot-services.git
cd robot-services
```

## 2. 只需改的配置

```powershell
copy feishu-ws-gateway\.env.example feishu-ws-gateway\.env
# 编辑 feishu-ws-gateway\.env：填 FEISHU_APP_SECRET

copy dealer-report\.env.example dealer-report\.env
# 经销商报表如果要用 AI，再填 AI_API_KEY
```

## 3. 一键启动当前在跑的服务

```powershell
powershell -ExecutionPolicy Bypass -File .\start-all.ps1
```

或按目录单独启动，例如木皮优化：

```powershell
cd door-skin-converter
docker compose up -d --build
```

## 服务目录

| 服务 | 端口 | compose 目录 |
|------|------|----------------|
| hardware-summary | 8001 | `hardware-summary/deploy/docker` |
| order-split | 8002 | `order-split/deploy/docker` |
| dealer-sales | 8003 | `may-sales/deploy/docker` |
| csv-board | 8004 | `csv-board/deploy/docker` |
| pvc-classify | 8005 | `pvc-classify/deploy/docker` |
| workshop-order | 8006 | `workshop-order/deploy/docker` |
| quote-maker | 8007 | `quote-maker/deploy/docker` |
| dealer-report | 8008 | `dealer-report/deploy/docker` |
| attendance-summary | 8009 | `attendance-summary/deploy/docker` |
| door-skin-converter | 8010 | `door-skin-converter` |
| order-summary | 8011 | `order-summary/deploy/docker` |
| pms-door-split | 8013 | `pms-door-split` |
| simple-ims | 8090 | `simple-ims/deploy/docker` |
| feishu-ws-gateway | — | `feishu-ws-gateway/deploy/docker`（最后启动） |

健康检查：`http://localhost:<端口>/health`

版本与群绑定见 `SERVICES.md`。
