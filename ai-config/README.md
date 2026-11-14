# AI 接入配置服务（8020）

一个只做配置的小服务：把「用哪家模型、key 是什么」从 compose/环境变量里搬出来，
改成可以随时用 HTTP 改，改完业务侧立刻生效，不用重启。

支持两种协议：

| provider | 端点 | 认证头 | 说明 |
|---|---|---|---|
| `openai` | `{base_url}/chat/completions` | `Authorization: Bearer <key>` | OpenAI、DeepSeek、通义、Moonshot 等兼容端点 |
| `anthropic` | `{base_url}/v1/messages` | `x-api-key` + `anthropic-version` | Claude 官方 API |

## 接口

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/health` | 存活 + 当前 provider、是否已配 key（不返回 key） |
| GET | `/ai-config` | 读配置，key 默认脱敏；`?reveal=1` + token 看明文 |
| PUT | `/ai-config` | 改配置，支持部分更新，写盘即生效 |
| POST | `/ai-config/test` | 用指定 provider 发一条最小请求，验证 key 能不能用 |

## 配置结构

```json
{
  "active": "openai",
  "providers": {
    "openai":    {"api_key": "sk-...", "base_url": "https://api.deepseek.com/v1", "model": "deepseek-v4-flash"},
    "anthropic": {"api_key": "sk-ant-...", "base_url": "https://api.anthropic.com", "model": "claude-sonnet-4-5-20250929"}
  },
  "options": {"temperature": 0.4, "max_tokens": 6000, "timeout": 120},
  "updated_at": "2026-11-07T17:45:00+08:00"
}
```

- `active` 决定业务默认用哪家，随时可切。
- 两家 key 可以同时存在，互不覆盖。
- 文件默认 `/data/ai_config.json`（容器内），宿主机是 `ai-config/config/ai_config.json`。
- 业务服务直接读这个文件，所以本服务挂了不影响业务，只是改不了配置。

## 用法

```bash
# 看当前配置（key 脱敏）
curl -s http://127.0.0.1:8020/ai-config

# 配 Anthropic 的 key（只传要改的字段）
curl -s -X PUT http://127.0.0.1:8020/ai-config \
  -H 'Content-Type: application/json' \
  -d '{"providers":{"anthropic":{"api_key":"sk-ant-xxx","model":"claude-sonnet-4-5-20250929"}}}'

# 切到 Anthropic
curl -s -X PUT http://127.0.0.1:8020/ai-config \
  -H 'Content-Type: application/json' -d '{"active":"anthropic"}'

# 也可以用简写，直接传 provider 名
curl -s -X PUT http://127.0.0.1:8020/ai-config \
  -H 'Content-Type: application/json' \
  -d '{"openai":{"api_key":"sk-xxx","base_url":"https://api.openai.com/v1","model":"gpt-4o-mini"}}'

# 试一下 key 通不通（override 可以先用没保存的 key 试）
curl -s -X POST http://127.0.0.1:8020/ai-config/test \
  -H 'Content-Type: application/json' \
  -d '{"provider":"anthropic","override":{"api_key":"sk-ant-xxx"}}'
```

`GET` 返回的是 `sk-ab***wxyz` 这种脱敏值，原样回传 `PUT` 时会被视为"不改 key"，不会把真 key 覆盖成星号。

## 鉴权

- 设了 `AI_CONFIG_TOKEN` 环境变量：`PUT`/`POST`/`reveal=1` 必须带 `X-Admin-Token: <token>`。
- 没设：写接口对所有来源开放（内网服务，启动日志会打 `auth=OPEN` 提示）。

## 命令行（推荐，省得手写 curl）

```bash
cd D:/Services/robot-services/ai-config/tools

./ai-config.sh get                                  # 看生产配置（key 脱敏）
./ai-config.sh put '{"active":"anthropic"}'         # 切模型
./ai-config.sh put '{"openai":{"api_key":"sk-xxx"}}'  # 换 key
./ai-config.sh test                                 # 验证当前 provider 通不通
./ai-config.sh get --local                          # 看本机 8020
```

默认操作**平台**上的服务（`ops@192.168.6.245:2222`，写的是 worker 真正读的那份配置），
加 `--local` 才动本机 8020。

## 部署

### 本机（容器，8020）

```bash
cd D:/Services/robot-services/ai-config/deploy/docker
docker compose up -d          # 纯标准库，无需 build
```

配置文件：`ai-config/config/ai_config.json`，只给本机容器/调试用。

### 平台（生产，systemd 常驻）

| 项 | 值 |
|---|---|
| 程序 | `/home/ops/ai-config/ai_config_server.py` |
| 服务 | `systemctl status ai-config`（`/etc/systemd/system/ai-config.service`） |
| 监听 | `127.0.0.1:8020`（只对本机，外部走 ssh） |
| 配置 | `/home/ops/robot-platform-tests/platform/services/dealer_report/ai_config.json` |
| 日志 | `/home/ops/ai-config/ai_config.log` |

**生产那份才是权威**：dealer-report worker 通过只读挂载直接读它，
改完下一次调用就生效，不用重启 worker。

```bash
# 平台侧排查三件套
sudo systemctl status ai-config
curl -s http://127.0.0.1:8020/health
sudo docker logs --tail 20 robot-platform-worker-dealer-report-1 | grep ai_client
```

## 错误码

| 代码 | 含义 |
|---|---|
| AICFG-E001 | 参数不合法（JSON、字段类型、base_url、模型名） |
| AICFG-E002 | 未知 provider（只认 openai / anthropic） |
| AICFG-E003 | 未授权（token 不对） |
| AICFG-E004 | 写配置文件失败 |
| AICFG-E101 | 该 provider 没配 key |
| AICFG-E102 | 上游返回非 2xx（key 无效、模型不存在等，消息里带上游原文） |
| AICFG-E103 | 连不上上游（网络/DNS/超时） |
| AICFG-E104 | 上游响应结构不符合该协议 |

## 谁在用

| 部署 | 服务 | 读取方式 |
|---|---|---|
| 平台（生产 8008） | `robot-platform-worker-dealer-report-1` | 读 `services/dealer_report/ai_config.json`（目录已只读挂进容器），改完即生效 |
| 本机 | `dealer-report-8008` | 读 `/app/config/ai_config.json`（挂 `ai-config/config`），改完即生效 |

新增服务接入：挂载配置文件或复制一份，用 `ai_client.py` 的 `chat()` 即可；
平台侧把 `ai_client.py` 放进服务目录（capability 已把目录加进 sys.path），配置放在同目录。
