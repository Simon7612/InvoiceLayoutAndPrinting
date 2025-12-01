import os
import io
import base64
import tempfile
import warnings
from typing import List, Optional, Tuple
from xml.etree import ElementTree as ET
from pypdf import PdfReader
warnings.filterwarnings("ignore", category=SyntaxWarning, module=r"^ofdparser(\.|$)")
warnings.filterwarnings("ignore", category=SyntaxWarning, module=r"^easyofd(\.|$)")
try:
    # 轻量纯 Python 库，支持 ofd -> pdf（返回 PDF bytes）
    from ofdparser import OfdParser  # type: ignore
    _HAS_OFDPARSER = True
except Exception:
    _HAS_OFDPARSER = False
try:
    # 备用库 easyofd（功能更全，但可能依赖 numpy/opencv）
    import easyofd  # type: ignore
    _HAS_EASYOFD = True
except Exception:
    _HAS_EASYOFD = False
"""
读取与收集发票/票据文件。

v0.0.2：支持 PDF；OFD/XML 直接转换（优先纯 Python 方案）。
"""
#TODO: 支持ofd和xml文件, 目前只支持pdf文件
def collect_pdfs(path: str) -> List[str]:
    """收集指定路径下的发票文件（PDF/OFD/XML）。"""
    def _accept(name: str) -> bool:
        n = name.lower()
        return n.endswith(".pdf") or n.endswith(".ofd") or n.endswith(".xml")
    if os.path.isdir(path):
        return sorted(
            [
                os.path.join(path, f)
                for f in os.listdir(path)
                if _accept(f)
            ]
        )
    return [path]

def read_pdf(input_path: str) -> PdfReader:
    """读取 PDF 文件并返回 PdfReader。"""
    if not os.path.exists(input_path):
        raise FileNotFoundError(input_path)
    ext = os.path.splitext(input_path)[1].lower()
    if ext != ".pdf":
        raise ValueError("暂不支持该文件类型，请使用 PDF")
    return PdfReader(input_path)

def _convert_ofd_to_pdf_bytes(ofd_path: str) -> bytes:
    """将 OFD 转换为 PDF，返回 PDF 二进制。

    优先使用纯 Python 库 `ofdparser`，其次尝试 `easyofd`。
    """
    if not os.path.exists(ofd_path):
        raise FileNotFoundError(ofd_path)
    with open(ofd_path, "rb") as f:
        ofd_bytes = f.read()

    # 先尝试 ofdparser（接口需要 base64 字符串）
    if _HAS_OFDPARSER:
        ofdb64 = base64.b64encode(ofd_bytes).decode("ascii")
        pdf_bytes = OfdParser(ofdb64).ofd2pdf()  # type: ignore
        if isinstance(pdf_bytes, (bytes, bytearray)) and len(pdf_bytes) > 10:
            return bytes(pdf_bytes)

    # 其次尝试 easyofd（若已安装）。API 在不同版本可能变化，这里尽量兼容常见用法。
    if _HAS_EASYOFD:
        # 常见用法：easyofd.ofd2pdf(in_bytes) -> bytes 或 easyofd.ofd2pdf(in_path, out_path)
        try:
            if hasattr(easyofd, "ofd2pdf"):
                res = easyofd.ofd2pdf(ofd_bytes)  # type: ignore
                if isinstance(res, (bytes, bytearray)):
                    return bytes(res)
        except Exception:
            # 尝试路径写出方式
            try:
                tmp_dir = tempfile.mkdtemp(prefix="easyofd_")
                out_pdf = os.path.join(tmp_dir, os.path.splitext(os.path.basename(ofd_path))[0] + ".pdf")
                easyofd.ofd2pdf(ofd_path, out_pdf)  # type: ignore
                with open(out_pdf, "rb") as pf:
                    return pf.read()
            except Exception:
                pass

    raise RuntimeError(
        "无法直接转换 OFD，请安装 ofdparser 或 easyofd（pip install ofdparser 或 easyofd）"
    )

def _decode_base64_safe(text: str) -> Optional[bytes]:
    """尝试将文本按 base64 解码，非法则返回 None。"""
    try:
        return base64.b64decode(text, validate=True)
    except Exception:
        return None

def _detect_payload_type(data: bytes) -> Optional[str]:
    """粗略检测数据类型：返回 'pdf' 或 'ofd' 或 None。"""
    if not data or len(data) < 4:
        return None
    if data.startswith(b"%PDF"):
        return "pdf"
    # OFD 是 ZIP 包，常见魔数为 PK\x03\x04
    if data.startswith(b"PK\x03\x04"):
        return "ofd"
    return None

def _extract_embedded_payload_from_xml(xml_path: str) -> Tuple[Optional[bytes], Optional[str]]:
    """从 XML 中提取内嵌的 PDF/OFD base64 数据。

    返回 (payload_bytes, payload_type)，若未找到则 (None, None)。
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # 遍历所有节点文本，尝试 base64 解码并识别
    for elem in root.iter():
        txt = (elem.text or "").strip()
        if not txt or len(txt) < 64:  # 过短的不可能是有效载荷
            continue
        blob = _decode_base64_safe(txt)
        if not blob:
            continue
        typ = _detect_payload_type(blob)
        if typ in {"pdf", "ofd"}:
            return blob, typ
    return None, None

def _convert_xml_to_pdf_bytes(xml_path: str) -> bytes:
    """将包含内嵌 PDF/OFD 的 XML 转换为 PDF bytes。

    限制：仅支持 XML 中内嵌的 base64 PDF/OFD 负载；
    若为 OFD 负载，则使用纯 Python 转换（ofdparser/easyofd）得到 PDF。
    """
    if not os.path.exists(xml_path):
        raise FileNotFoundError(xml_path)
    payload, typ = _extract_embedded_payload_from_xml(xml_path)
    if not payload or not typ:
        raise RuntimeError("未在 XML 中发现内嵌的 PDF/OFD base64 数据，无法直接转换")
    if typ == "pdf":
        return payload
    if typ == "ofd":
        # ofdparser 接口需要 base64 字符串
        if _HAS_OFDPARSER:
            ofdb64 = base64.b64encode(payload).decode("ascii")
            pdf_bytes = OfdParser(ofdb64).ofd2pdf()  # type: ignore
            return bytes(pdf_bytes)
        if _HAS_EASYOFD:
            try:
                if hasattr(easyofd, "ofd2pdf"):
                    res = easyofd.ofd2pdf(payload)  # type: ignore
                    if isinstance(res, (bytes, bytearray)):
                        return bytes(res)
            except Exception:
                pass
        raise RuntimeError("XML 含 OFD 负载，但未安装 ofdparser/easyofd，无法直接转换")
    raise RuntimeError("未知负载类型，无法转换")

def read_document(input_path: str) -> PdfReader:
    """读取 PDF/OFD/XML，并返回 PdfReader。

    - PDF：直接读取
    - OFD：使用纯 Python 转换为 PDF 后读取
    - XML：从内嵌 base64 中提取 PDF/OFD 负载并转换为 PDF 后读取
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(input_path)
    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".pdf":
        return PdfReader(input_path)
    if ext == ".ofd":
        pdf_bytes = _convert_ofd_to_pdf_bytes(input_path)
        return PdfReader(io.BytesIO(pdf_bytes))
    if ext == ".xml":
        pdf_bytes = _convert_xml_to_pdf_bytes(input_path)
        return PdfReader(io.BytesIO(pdf_bytes))
    raise ValueError(f"不支持的文件类型: {ext}")

# ---------------------------
# 车票识别辅助
# ---------------------------

_TICKET_KEYWORDS_CN = [
    "车次", "座位", "检票口", "中国铁路", "高铁", "动车", "出发", "到达",
    "发车", "到站", "站台", "票价", "检票", "一等座", "二等座", "商务座", "硬座", "软卧",
]
_TICKET_KEYWORDS_EN = [
    "train", "seat", "gate", "railway", "departure", "arrival", "platform", "fare",
]

def detect_ticket_document(reader: PdfReader, max_pages: int = 3) -> Tuple[bool, Optional[str]]:
    """基于关键词识别是否为车票，并返回推荐方向（portrait/landscape）。

    规则：
    - 提取前 `max_pages` 页文本，匹配常见中文/英文关键词。
    - 推荐方向依据第一页的页面尺寸：宽>=高→landscape，否则→portrait。
    """
    try:
        n = min(len(reader.pages), max_pages)
    except Exception:
        n = 0
    matched = False
    for i in range(n):
        try:
            txt = reader.pages[i].extract_text() or ""
        except Exception:
            txt = ""
        low = txt.lower()
        # 中文关键词直接 in 检测（不转小写）
        if any(k in txt for k in _TICKET_KEYWORDS_CN) or any(k in low for k in _TICKET_KEYWORDS_EN):
            matched = True
            break
    # 推荐方向
    orient: Optional[str] = None
    try:
        p0 = reader.pages[0]
        w = float(p0.mediabox.width)
        h = float(p0.mediabox.height)
        orient = "landscape" if w >= h else "portrait"
    except Exception:
        pass
    return matched, orient
