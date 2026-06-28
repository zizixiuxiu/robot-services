"""AI recognizers for order entry from images.

This module is pluggable: implement BaseRecognizer and set RECOGNIZER_CLASS
in the FastAPI app state to swap backends.
"""

from __future__ import annotations

import os
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any


@dataclass
class OrderItem:
    seq: int = 1
    item_code: str = ""
    name: str = ""
    spec: str = ""
    unit: str = "张"
    qty: Decimal = Decimal("0")
    remark: str = ""


@dataclass
class RecognizeResult:
    document_no: str = ""
    date: str = ""
    department: str = ""
    description: str = ""
    remark: str = ""
    items: list[OrderItem] = field(default_factory=list)
    raw_text: str = ""
    recognizer: str = ""
    confidence: str = "low"


class BaseRecognizer(ABC):
    @abstractmethod
    def recognize(self, image_path: str) -> RecognizeResult:
        raise NotImplementedError


class MockRecognizer(BaseRecognizer):
    """Demo recognizer that returns the known structure for the sample ERP screenshot.

    Use this to validate the API plumbing when no OCR/vision backend is available.
    """

    def recognize(self, image_path: str) -> RecognizeResult:
        # In a real implementation this would inspect the image bytes.
        # Here we produce the parsed structure for the sample procurement screenshot.
        return RecognizeResult(
            document_no="CJ202606000272",
            date="2026-06-22",
            department="",
            description="6.22特急规定，陈瑞利",
            remark="PVC",
            items=[
                OrderItem(
                    seq=1,
                    item_code="10102122439",
                    name="MD-09银色E0颗粒板",
                    spec="2800*1220*18",
                    unit="张",
                    qty=Decimal("3.00"),
                    remark="特急，（膜皮脱层）陈瑞利B2606-8059",
                )
            ],
            raw_text=(
                "单据编号 CJ202606000272 审核完成 编制时间 2026-06-22 "
                "部门 计划说明 6.22特急规定，陈瑞利 备注 PVC "
                "物品编码 10102122439 物品名称 MD-09银色E0颗粒板 "
                "规格型号 2800*1220*18 计量单位 张 计划采购数量 3.00张 "
                "备注 特急，（膜皮脱层）陈瑞利B2606-8059"
            ),
            recognizer="mock",
            confidence="demo",
        )


class KimiVisionRecognizer(BaseRecognizer):
    """Recognizer backed by the Moonshot (Kimi) vision model.

    Requires the environment variable MOONSHOT_API_KEY to be set.
    """

    def __init__(self, api_key: str | None = None, model: str = "moonshot-v1-32k-vision-preview") -> None:
        self.api_key = api_key or os.getenv("MOONSHOT_API_KEY")
        if not self.api_key:
            raise RuntimeError(
                "MOONSHOT_API_KEY is required for KimiVisionRecognizer. "
                "Set it as an environment variable or pass api_key."
            )
        self.model = model

    def _preprocess_image(self, image_path: str) -> str:
        """Upscale and enhance contrast to improve OCR accuracy."""
        from PIL import Image, ImageEnhance

        img = Image.open(image_path)
        if img.mode != "RGB":
            img = img.convert("RGB")

        # Upscale small screenshots so fine text is easier to read.
        width, height = img.size
        if width < 1600 or height < 900:
            scale = max(2, min(4, int(2000 / max(width, 1))))
            img = img.resize((width * scale, height * scale), Image.LANCZOS)

        # Slightly increase contrast to separate text from background.
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.4)

        suffix = os.path.splitext(image_path)[1] or ".png"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        img.save(tmp.name, quality=95)
        return tmp.name

    def recognize(self, image_path: str) -> RecognizeResult:
        import base64
        import json
        import urllib.request

        processed_path = self._preprocess_image(image_path)
        try:
            with open(processed_path, "rb") as f:
                image_bytes = f.read()
        finally:
            os.unlink(processed_path)

        b64_image = base64.b64encode(image_bytes).decode("utf-8")

        system_prompt = (
            "你是一个订单识别助手。请严格按图片中实际可见的内容提取采购/订单信息，"
            "特别注意物品编码、物品名称、规格型号等字段的每个字符，"
            "不要漏字、不要错字。以严格的 JSON 格式返回，不要包含任何额外解释。\n"
            "重要规则：\n"
            "1. 图片里有什么就提取什么；图片里没有或看不清的字段，必须返回空字符串，禁止编造、禁止推断、禁止用列标题或示例值填充。\n"
            "2. 物品名称中的每个汉字都要逐字核对，不要凭上下文猜测，避免因字形相近而误读（例如'琳琅'不要读成'拼银'）。\n"
            "3. 单据头部（表头区域）的备注字段，输出到根级 remark；明细表格中每一行的备注列，输出到对应 item.remark。\n"
            "4. 对于每个 item 的 remark，必须读取该行备注列单元格内的实际内容；"
            "如果单元格为空或只有空白，则 item.remark 必须填空字符串，"
            "绝对不要把列标题（如'计量单位'、'主计量单位'、'采购计价单位'、'备注'）当作单元格内容。\n"
            "JSON 结构如下：\n"
            "{\"document_no\": \"单据编号\", \"date\": \"YYYY-MM-DD\", "
            "\"department\": \"部门\", \"description\": \"计划说明\", "
            "\"remark\": \"单据头部备注\", "
            "\"items\": [{\"seq\": 1, \"item_code\": \"物品编码\", "
            "\"name\": \"物品名称\", \"spec\": \"规格型号\", "
            "\"unit\": \"计量单位\", \"qty\": \"数量(字符串)\", "
            "\"remark\": \"明细行备注\"}]}"
        )

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64_image}"}},
                        {
                            "type": "text",
                            "text": (
                                "请提取这张单据的信息并返回 JSON。\n"
                                "注意：\n"
                                "1. 严格按图片实际内容提取；没有的信息（包括备注、规格、数量等）必须留空，不要编造。\n"
                                "2. 物品名称中的每个汉字都要逐字核对，不要因字形相近而误读（例如'琳琅'不要读成'拼银'）。\n"
                                "3. 单据表头区域有一个'备注'字段，对应根级 remark。\n"
                                "4. 明细表格的列从左到右依次是：序号、物品编码、物品名称、备注、规格型号、是否赠品、计量单位、数量、采购计价单位。\n"
                                "每个 item 的 remark 必须取自'备注'列该行单元格的实际内容；"
                                "如果单元格为空或只有空白，则 item.remark 必须填''，"
                                "不要把'计量单位'、'主计量单位'、'采购计价单位'等列标题或其他列的值写进 remark。"
                            ),
                        },
                    ],
                },
            ],
            "temperature": 0.2,
            "response_format": {"type": "json_object"},
        }

        req = urllib.request.Request(
            "https://api.moonshot.cn/v1/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )

        with urllib.request.urlopen(req, timeout=120) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        content = data["choices"][0]["message"]["content"]
        parsed: dict[str, Any] = json.loads(content)

        def _clean_remark(value: str) -> str:
            """Drop values that are clearly column headers rather than cell content."""
            if not value:
                return ""
            normalized = value.strip().replace(" ", "").replace("\u3000", "")
            header_like = {"备注", "计量单位", "主计量单位", "采购计价单位", "单位"}
            if normalized in header_like:
                return ""
            return value.strip()

        items = []
        for raw_item in parsed.get("items", []):
            items.append(
                OrderItem(
                    seq=int(raw_item.get("seq", 1)),
                    item_code=str(raw_item.get("item_code", "")),
                    name=str(raw_item.get("name", "")),
                    spec=str(raw_item.get("spec", "")),
                    unit=str(raw_item.get("unit", "张")),
                    qty=Decimal(str(raw_item.get("qty", "0"))),
                    remark=_clean_remark(str(raw_item.get("remark", ""))),
                )
            )

        return RecognizeResult(
            document_no=str(parsed.get("document_no", "")),
            date=str(parsed.get("date", "")),
            department=str(parsed.get("department", "")),
            description=str(parsed.get("description", "")),
            remark=str(parsed.get("remark", "")),
            items=items,
            raw_text=content,
            recognizer="kimi-vision",
            confidence="high",
        )


class PaddleOCRRecognizer(BaseRecognizer):
    """Recognizer backed by the PaddleOCR cloud API.

    Requires the environment variable PADDLE_OCR_TOKEN to be set.
    """

    JOB_URL = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs"
    MODEL = "PaddleOCR-VL-1.6"

    def __init__(self, token: str | None = None) -> None:
        self.token = token or os.getenv("PADDLE_OCR_TOKEN")
        if not self.token:
            raise RuntimeError(
                "PADDLE_OCR_TOKEN is required for PaddleOCRRecognizer. "
                "Set it as an environment variable or pass token."
            )

    def _submit_job(self, image_path: str) -> str:
        import json as _json
        import time

        import requests

        headers = {"Authorization": f"bearer {self.token}"}
        data = {
            "model": self.MODEL,
            "optionalPayload": _json.dumps(
                {
                    "useDocOrientationClassify": False,
                    "useDocUnwarping": False,
                    "useChartRecognition": False,
                }
            ),
        }
        with open(image_path, "rb") as f:
            files = {"file": f}
            resp = requests.post(self.JOB_URL, headers=headers, data=data, files=files)

        if resp.status_code != 200:
            raise RuntimeError(f"PaddleOCR job submission failed: {resp.status_code} {resp.text}")

        return resp.json()["data"]["jobId"]

    def _poll_result(self, job_id: str) -> str:
        import time

        import requests

        headers = {"Authorization": f"bearer {self.token}"}
        for _ in range(60):
            resp = requests.get(f"{self.JOB_URL}/{job_id}", headers=headers)
            resp.raise_for_status()
            state = resp.json()["data"]["state"]
            if state == "done":
                return resp.json()["data"]["resultUrl"]["jsonUrl"]
            if state == "failed":
                error_msg = resp.json()["data"].get("errorMsg", "unknown error")
                raise RuntimeError(f"PaddleOCR job failed: {error_msg}")
            time.sleep(3)
        raise RuntimeError("PaddleOCR job polling timed out")

    def _extract_header_remark(self, text: str) -> str:
        """Best-effort extraction of the header-level remark from OCR text."""
        for line in text.splitlines():
            line = line.strip()
            if line.startswith("备注") and len(line) > 2:
                return line[2:].strip().strip(":：")
        return ""

    def _parse_table(self, table_html: str) -> list[OrderItem]:
        """Parse the HTML table returned by PaddleOCR into OrderItems."""
        from html.parser import HTMLParser

        class TableParser(HTMLParser):
            def __init__(self) -> None:
                super().__init__()
                self.rows: list[list[str]] = []
                self.current_row: list[str] = []
                self.current_cell: list[str] = []
                self.in_cell = False

            def handle_starttag(self, tag: str, attrs: list) -> None:
                if tag == "tr":
                    self.current_row = []
                elif tag in ("td", "th"):
                    self.in_cell = True
                    self.current_cell = []

            def handle_endtag(self, tag: str) -> None:
                if tag in ("td", "th"):
                    self.in_cell = False
                    self.current_row.append("".join(self.current_cell).strip())
                elif tag == "tr" and self.current_row:
                    self.rows.append(self.current_row)

            def handle_data(self, data: str) -> None:
                if self.in_cell:
                    self.current_cell.append(data)

        parser = TableParser()
        parser.feed(table_html)

        if not parser.rows:
            return []

        # Find the header row to map column names to indices.
        header_idx = 0
        for idx, row in enumerate(parser.rows):
            if "物品编码" in row or "物品名称" in row:
                header_idx = idx
                break

        header = parser.rows[header_idx]
        col_map = {name: i for i, name in enumerate(header)}

        def get(row: list[str], key: str, default: str = "") -> str:
            idx = col_map.get(key)
            if idx is None or idx >= len(row):
                return default
            return row[idx]

        items: list[OrderItem] = []
        for row in parser.rows[header_idx + 1 :]:
            if not row or not get(row, "物品编码") or get(row, "物品编码") == "合计":
                continue
            name = get(row, "物品名称")
            spec = get(row, "规格型号")
            unit = get(row, "计量单位") or "张"
            qty_text = get(row, "数量") or "0"
            remark = get(row, "备注")

            # Drop column-header-like values that leaked into cells.
            header_like = {"计量单位", "主计量单位", "采购计价单位", "备注", "单位"}
            if remark.replace(" ", "").replace("\u3000", "") in header_like:
                remark = ""
            if unit.replace(" ", "").replace("\u3000", "") in header_like:
                unit = "张"

            try:
                qty = Decimal(str(qty_text))
            except Exception:
                qty = Decimal("0")

            items.append(
                OrderItem(
                    seq=len(items) + 1,
                    item_code=get(row, "物品编码"),
                    name=name,
                    spec=spec,
                    unit=unit,
                    qty=qty,
                    remark=remark,
                )
            )

        return items

    def recognize(self, image_path: str) -> RecognizeResult:
        import json as _json
        import requests

        job_id = self._submit_job(image_path)
        jsonl_url = self._poll_result(job_id)
        jsonl_text = requests.get(jsonl_url).text

        # Combine all text blocks and tables from the result.
        text_parts: list[str] = []
        table_htmls: list[str] = []
        for line in jsonl_text.strip().splitlines():
            if not line.strip():
                continue
            data = _json.loads(line)
            result = data.get("result", {})
            for layout in result.get("layoutParsingResults", []):
                pruned = layout.get("prunedResult", {})
                for block in pruned.get("parsing_res_list", []):
                    label = block.get("block_label", "")
                    content = block.get("block_content", "")
                    if label == "table" and content:
                        table_htmls.append(content)
                    elif content:
                        text_parts.append(content)

        all_text = "\n".join(text_parts)
        items: list[OrderItem] = []
        for html in table_htmls:
            items.extend(self._parse_table(html))

        return RecognizeResult(
            document_no="",
            date="",
            department="",
            description="",
            remark=self._extract_header_remark(all_text),
            items=items,
            raw_text=jsonl_text,
            recognizer="paddleocr",
            confidence="high",
        )


def get_recognizer() -> BaseRecognizer:
    """Factory that returns the configured recognizer.

    Priority:
      1. PADDLE_OCR_TOKEN env var -> PaddleOCRRecognizer
      2. MOONSHOT_API_KEY env var -> KimiVisionRecognizer
      3. Otherwise -> MockRecognizer (for API plumbing tests)
    """
    paddle_token = os.getenv("PADDLE_OCR_TOKEN")
    if paddle_token:
        return PaddleOCRRecognizer(token=paddle_token)
    api_key = os.getenv("MOONSHOT_API_KEY")
    if api_key:
        return KimiVisionRecognizer(api_key=api_key)
    return MockRecognizer()
