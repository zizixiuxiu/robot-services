# 门扇清单转换服务

这是一个可在 Windows Docker Desktop 上部署的本地服务。上传门扇 `.xls`（也兼容 `.xlsx`）后，服务会按 26 列新表头下载转换完成的 GB18030 编码 CSV，不依赖外部模板文件。

## 一、Windows 部署

1. 安装并启动 Docker Desktop。
2. 解压本源码包，在文件夹空白处按住 `Shift` 并右键，选择“在终端中打开”。
3. 执行：

```powershell
docker compose up -d --build
```

4. 浏览器打开：<http://localhost:8080>

停止服务：

```powershell
docker compose down
```

更新源码后重新部署：

```powershell
docker compose up -d --build
```

## 二、接口调用

健康检查：

```text
GET /healthz
```

转换接口：

```text
POST /api/convert
Content-Type: multipart/form-data
字段名：file
```

Windows 命令行示例：

```powershell
curl.exe -f -X POST -F "file=@C:\订单\门扇.xls" http://localhost:8080/api/convert -o "C:\订单\门扇_转换.csv"
```

接口成功时会返回 CSV，并附带以下响应头：

- `X-Source-Rows`：源数据行数
- `X-Output-Rows`：输出数据行数
- `X-Quantity-Sum`：输出数量合计

## 三、已固化的转换规则

| 输出列 | 规则 |
|---|---|
| 订单号 | 源表“订单编号” |
| 板件名称 | 源表“工件名称” |
| 加工长度 | 源表“开料长” |
| 加工宽度 | 源表“开料宽” |
| 数量 | 每条输出行均为源数量 × 2 |
| 材料描述 | 最终厚度 + `mm` + 源表“颜色”，如 `5mmPY12-晚秋胡桃` |
| 纹理 | 固定为 `1` |
| 特殊工艺 | `材料 + 厚度 + mm + 特殊要求`，如 `门扇厚度50mm.不开锁孔` |
| 客户 | 源表“客户地址” |
| 品牌 | 源表“品牌” |
| 厚度 | 从“工艺”提取；`5+8mm` 拆成 5 和 8 两行 |
| 材料描述2 | 与“材料描述”完全相同 |
| 其余新表头列 | 留空 |

补充规则：

- `4+4mm` 会拆成两条厚度均为 4 的记录。
- “工艺”中没有可提取厚度时，最终厚度留空，材料描述不添加无效的 `mm`。
- “特殊要求”正文不改写，只在前面拼接材料和门扇厚度。
- CSV 使用与新模板相同的 GB18030 编码及 CRLF 换行。

## 四、本地命令行运行（不使用 Docker）

需要 Python 3.11 或更高版本：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m app.cli "C:\订单\门扇.xls" -o "C:\订单\门扇_转换.csv"
```

## 五、运行自动化测试

```powershell
python -m unittest discover -s tests -v
```

## 六、常见问题

- 提示缺少字段：请检查源文件第一张工作表的表头，字段名必须包含规则表所需列。
- 端口被占用：修改 `docker-compose.yml` 中左侧端口，例如改成 `8088:8080`，然后访问 `http://localhost:8088`。
- 外部电脑访问：把 `localhost` 换成部署电脑的局域网 IP，并在 Windows 防火墙中放行对应端口。
- 默认最大上传文件为 50MB，可修改 `docker-compose.yml` 中的 `MAX_UPLOAD_BYTES`。
