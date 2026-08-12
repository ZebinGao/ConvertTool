#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf2txt.py — 将 PDF 转换为 TXT 纯文本文件（文本型 / 扫描图片型均可）。

依赖：
    - PyMuPDF（fitz）            —— 必装，文本提取与页面渲染
    - rapidocr (>=3.9.0)         —— 可选，扫描/图片页 OCR（--ocr 时用）。默认即 PP-OCRv6 的
                                    ch 模型，单个模型同时识别简体/繁体中文，故扫描繁体书配合
                                    --t2s 无需另下繁体识别模型（识别后再繁→简即可）。首次会自动
                                    从魔搭（ModelScope）下载模型。
    - zhconv                     —— 可选，繁体转简体（--t2s 时用）

常见用法：
    python pdf2txt.py 输入.pdf                  # 文本型 PDF → 同目录同名 .txt
    python pdf2txt.py 版面型.pdf --no-unwrap      # 保留原始换行（默认会合并软换行）
    python pdf2txt.py 直排书.pdf                  # 竖排 PDF 自动按列重建（无需额外参数）
    python pdf2txt.py 直排书.pdf --layout v       # 强制按直排处理（自动检测不准时）
    python pdf2txt.py 扫描书.pdf --ocr             # 扫描/图片型 PDF，走 OCR（默认 300 DPI）
    python pdf2txt.py 繁体书.pdf --t2s             # 繁体转简体
    python pdf2txt.py 扫描繁体书.pdf --ocr --t2s   # OCR（v6 自动认繁体）+ 繁简转换
    python pdf2txt.py 资料/ -o out/                # 批量转换整个目录（递归）
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import fitz  # PyMuPDF

# Windows 控制台默认编码可能不是 UTF-8，统一一下，避免打印中文报错
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


# 句末标点：行尾出现这些字符时，认为该行是段落结束（硬换行）
_SENT_END = set("。！？!?.…;；：”」』)）")

# 直排（竖排）版面重建相关常量。直排 PDF 被 PyMuPDF 逐字提取（每字一行），
# 逐行启发式无法还原阅读顺序；改用位置：同一竖列的字 x 中心相同，
# 列从右到左、列内从上到下。
_COL_TOL = 8.0     # 列聚类：两行 x 中心差 ≤ 此值视为同一列（pt）
_GAP_FACTOR = 1.6  # 列内纵向空隙 > 中位步长 × 此系数 → 视为段落分隔

# 页码模式：整行仅为一个（可带装饰的）数字，或"第N页"
_PAGE_NUM_RE = re.compile(
    r"^\s*[\-—‑–_=*·.]*\s*\d+\s*[\-—‑–_=*·.]*\s*$"
    r"|^\s*第\s*\d+\s*[页頁]\s*$"
)


def _unwrap_text(text: str) -> str:
    """合并段落内的软换行（排版折行），保留真正的段落分隔。

    判断规则（启发式）：
    - 空行、行首缩进、或上一行以句末标点结尾 → 视为段落边界，保留换行；
    - 否则视为排版折行（软换行），与上一行合并 —— 中文直接拼接，
      中文与字母/数字交界处补一个空格，避免粘连。
    局限：页眉页脚、多栏排版等无标点的短行可能被误合并，需另外处理。
    """
    paras: list[str] = []
    cur = ""

    def flush():
        nonlocal cur
        if cur.strip():
            paras.append(cur.strip())
        cur = ""

    for ln in text.splitlines():
        s = ln.strip()
        if not s:                       # 空行 → 段落分隔
            flush()
            continue
        if ln != ln.lstrip():           # 行首缩进 → 新段落
            flush()
            cur = s
            continue
        if not cur:
            cur = s
            continue
        if cur[-1] in _SENT_END:        # 上一行句末 → 段落结束
            flush()
            cur = s
        else:                           # 软换行 → 合并
            need_space = (cur[-1].isalnum() != s[0].isalnum()) and (cur[-1].isascii() or s[0].isascii())
            cur += (" " if need_space else "") + s
    flush()
    return "\n\n".join(paras)


# ---------- 版面检测与直排（竖排）重建 ----------
@dataclass
class _Line:
    """dict 里的一个文本行，附带 bbox，用于布局判定与位置重建。"""
    text: str
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def xc(self) -> float:   # x 中心
        return (self.x0 + self.x1) / 2

    @property
    def w(self) -> float:
        return self.x1 - self.x0

    @property
    def h(self) -> float:
        return self.y1 - self.y0


def _extract_lines(page) -> list[_Line]:
    """取页面内所有文本行（含 bbox）；图片块跳过。"""
    lines: list[_Line] = []
    for b in page.get_text("dict")["blocks"]:
        if b.get("type", 0) != 0:          # 0=文本，1=图片
            continue
        for ln in b.get("lines", []):
            t = "".join(s.get("text", "") for s in ln.get("spans", []))
            if not t.strip():
                continue
            x0, y0, x1, y1 = ln["bbox"]
            lines.append(_Line(t, x0, y0, x1, y1))
    return lines


def _is_vertical_page(lines: list[_Line]) -> bool:
    """竖排页判定：'高>宽' 的行占比 > 50%。
    直排逐字 PDF 每行单字（高>宽）；横排每行宽扁（宽>高）。"""
    if not lines:
        return False
    tall = sum(1 for ln in lines if ln.h > ln.w)
    return tall / len(lines) > 0.5


def _norm_line(s: str) -> str:
    """规范化行文本用于跨页比对：去首尾空白，内部空白（含 NBSP）压成单空格。"""
    return re.sub(r"\s+", " ", s.replace("\xa0", " ")).strip()


@dataclass
class _HFItem:
    """页眉/页脚剥离规则。"""
    dropset: set[str]       # 高频复现的短文本（如书名"罗织经"）→ 丢弃
    strip_pagenum: bool     # 是否剥离整行页码（仅当多页都检出页码时才开启）


def _compute_hf(per_page_lines: list[list[_Line]], n_pages: int) -> _HFItem:
    """跨页统计，得出页眉/页脚剥离规则。

    - 短文本（≤10 字）在 ≥ max(3, 60%页) 上出现 → 加入 dropset；
    - 页码行（匹配 _PAGE_NUM_RE）在 ≥ max(3, 60%页) 上出现 → 开启 strip_pagenum。
    阈值保守，小文档基本不触发，避免误伤正文。
    """
    pages_with: dict[str, set[int]] = {}
    pn_pages = 0
    for pi, lines in enumerate(per_page_lines):
        seen: set[str] = set()
        has_pn = False
        for ln in lines:
            key = _norm_line(ln.text)
            if 2 <= len(key) <= 10:      # 单字行（直排正文逐字）不参与，避免误伤
                seen.add(key)
            if len(key) >= 2 and _PAGE_NUM_RE.match(ln.text):
                has_pn = True
        for key in seen:
            pages_with.setdefault(key, set()).add(pi)
        if has_pn:
            pn_pages += 1
    thr = max(3, int(n_pages * 0.6))
    dropset = {k for k, ps in pages_with.items() if len(ps) >= thr}
    return _HFItem(dropset=dropset, strip_pagenum=pn_pages >= thr)


def _is_hf(line_text: str, hf: _HFItem) -> bool:
    """该行是否为页眉/页脚（规范化文本命中 dropset，或命中页码正则）。
    单字行直接返回 False：直排 PDF 里单字行是正文，绝非页眉/页脚。"""
    key = _norm_line(line_text)
    if len(key) < 2:
        return False
    if key in hf.dropset:
        return True
    return hf.strip_pagenum and bool(_PAGE_NUM_RE.match(line_text))


def _strip_hf_lines(text: str, hf: _HFItem) -> str:
    """从纯文本按行剥离页眉/页脚（横排路径用）。"""
    if not hf.dropset and not hf.strip_pagenum:
        return text
    return "\n".join(ln for ln in text.splitlines() if not _is_hf(ln, hf))


def _join_zh(a: str, b: str) -> str:
    """拼接两段：中文直接连；中文与 ASCII 字母数字交界处补一个空格。"""
    if not a or not b:
        return a + b
    need_space = (a[-1].isalnum() != b[0].isalnum()) and (a[-1].isascii() or b[0].isascii())
    return a + (" " if need_space else "") + b


def _reconstruct_vertical(lines: list[_Line]) -> str:
    """直排重建：按 x 中心聚类成列（容差 _COL_TOL），列从右到左、列内从上到下，
    还原阅读顺序。

    - 列内纵向空隙 > 中位步长 × _GAP_FACTOR → 段落分隔（空行，忠实版面）；
    - 列与列交界 → 软换行（单换行）：保留原版面"每列一行"的折行，便于阅读，
      不把版面折行误当作段落边界。
    """
    if not lines:
        return ""
    # 1) 聚类成列：归入最近的已有列（x 中心差 ≤ _COL_TOL）
    cols: dict[float, list[_Line]] = {}
    for ln in lines:
        best_key, best_d = None, _COL_TOL
        for k in cols:
            d = abs(k - ln.xc)
            if d <= best_d:
                best_key, best_d = k, d
        key = best_key if best_key is not None else ln.xc
        cols.setdefault(key, []).append(ln)
    # 2) 排成阅读顺序：列右→左，列内上→下；记下所属列以便区分列边界
    ordered: list[tuple[float, _Line]] = []
    for ck in sorted(cols.keys(), reverse=True):
        for ln in sorted(cols[ck], key=lambda z: z.y0):
            ordered.append((ck, ln))
    # 3) 用"同列相邻步长"的中位数定段落断点阈值
    steps = sorted(b.y0 - a.y0 for (ckp, a), (ckc, b)
                   in zip(ordered, ordered[1:]) if ckp == ckc)
    if steps:
        gap_thr = steps[len(steps) // 2] * _GAP_FACTOR
    else:
        gap_thr = ordered[0][1].h * _GAP_FACTOR
    # 4) 顺序拼接：列边界→软换行，列内大空隙→段落分隔，否则直接连
    out = ordered[0][1].text
    for (ckp, prev), (ckc, cur) in zip(ordered, ordered[1:]):
        if ckp != ckc:
            out += "\n" + cur.text
        elif cur.y0 - prev.y0 > gap_thr:
            out += "\n\n" + cur.text
        else:
            out = _join_zh(out, cur.text)
    return out.strip()


# ---------- 横排多栏（扫描杂志 / 学术书）阅读顺序重建 ----------
def _lines_from_boxes(txts, boxes) -> list[_Line]:
    """从 rapidocr 的 (txts, boxes) 构建 _Line 列表。boxes 形如 [N,4,2] 的像素坐标。"""
    import numpy as np
    lines: list[_Line] = []
    for t, b in zip(txts, boxes):
        b = np.asarray(b)
        xs, ys = b[:, 0], b[:, 1]
        lines.append(_Line(str(t), float(xs.min()), float(ys.min()),
                           float(xs.max()), float(ys.max())))
    return lines


def _count_side_by_side(lines: list[_Line], page_w: float) -> int:
    """统计「有并排邻居」的行数：同一高度(y 区间重叠)且横向不交叠(一左一右)。

    这是区分单/双栏的本质判据——双栏页同高度有左右两段文字并排；单栏页所有行
    上下堆叠、无并排。比「x 中心间距」(单栏里的短行会被误判成单独一列)和
    「中段密度谷」(任何跨栏标题/水印横穿中段都会把谷填平)都稳得多。
    按 y0 排序后用扫描窗口，复杂度 ≈ O(n·窗口宽)。"""
    n = len(lines)
    if n < 6:
        return 0
    gap_min = 0.02 * page_w          # 横向至少隔 2% 页宽才算「不交叠」
    with_nbr = [False] * n
    order = sorted(range(n), key=lambda i: lines[i].y0)
    for a in range(n):
        i = order[a]
        li = lines[i]
        for b in range(a + 1, n):
            j = order[b]
            lj = lines[j]
            if lj.y0 > li.y1:        # j 已在 i 下方，其后更下，不再可能重叠
                break
            if min(li.y1, lj.y1) - max(li.y0, lj.y0) <= 0:
                continue
            if (li.x1 + gap_min <= lj.x0) or (lj.x1 + gap_min <= li.x0):
                with_nbr[i] = with_nbr[j] = True
    return sum(with_nbr)


def _column_gutter(lines: list[_Line], page_w: float) -> float:
    """估算左右栏分界 x：取所有「并排对」边界中点的中位数。

    用实际不交叠对的边界定位栏沟，比「行覆盖密度的最小值」稳——后者会被偏宽的
    OCR 检测框填满（很多框略越过栏沟，密度谷就不明显了）。"""
    import statistics
    gap_min = 0.02 * page_w
    mids: list[float] = []
    for i in range(len(lines)):
        li = lines[i]
        for j in range(i + 1, len(lines)):
            lj = lines[j]
            if min(li.y1, lj.y1) - max(li.y0, lj.y0) <= 0:
                continue
            if li.x1 + gap_min <= lj.x0:
                mids.append((li.x1 + lj.x0) / 2)
            elif lj.x1 + gap_min <= li.x0:
                mids.append((lj.x1 + li.x0) / 2)
    return statistics.median(mids) if mids else page_w / 2


def _order_horizontal(lines: list[_Line], page_w: float) -> list[_Line]:
    """横排版面阅读顺序：多栏 → 列从左到右、列内从上到下；单栏 → 按 y 从上到下。

    输出仍是「一行一条」(每个 _Line 一条)，段落合并交给下游 _unwrap_text。
    双栏页左列底→右列顶本是连续正文，_unwrap_text 的软换行合并正好接上，无需特殊处理。"""
    if not lines:
        return []
    nbr = _count_side_by_side(lines, page_w)
    if nbr < 5 or nbr < 0.2 * len(lines):     # 不足 5 行或不足 20% → 单栏
        return sorted(lines, key=lambda z: (z.y0, z.x0))
    gutter = _column_gutter(lines, page_w)
    left = sorted((ln for ln in lines if ln.xc <= gutter), key=lambda z: z.y0)
    right = sorted((ln for ln in lines if ln.xc > gutter), key=lambda z: z.y0)
    return left + right


def _ocr_text_from_result(txts, boxes, page_w: float) -> str:
    """把 rapidocr 单页结果按版面排成纯文本（每行一条），供下游清洗 / unwrap。

    竖排走 _reconstruct_vertical；横排走 _order_horizontal（含多栏检测）。
    两条 OCR 路径（CLI 的 _ocr_page、GUI worker）共用此函数，保证一致。"""
    if not txts:
        return ""
    lines = _lines_from_boxes(txts, boxes)
    if _is_vertical_page(lines):
        return _reconstruct_vertical(lines)
    return "\n".join(ln.text for ln in _order_horizontal(lines, page_w))


# ---------- OCR（扫描 / 图片页）----------
_OCR_ENGINE = None


def _get_ocr():
    """懒加载 RapidOCR 单例（仅 --ocr 时才初始化，避免无谓加载）。

    用新版 rapidocr（统一包，>=3.9.0）：默认即 PP-OCRv6 的 ch 模型，单个模型同时覆盖
    简体/繁体中文，故 --t2s（繁→简）无需另下繁体识别模型，识别后再做繁简转换即可。
    首次实例化会从魔搭（ModelScope）自动下载 v6 模型。"""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        try:
            from rapidocr import RapidOCR
        except ImportError:
            raise RuntimeError("未安装 rapidocr，请先 pip install rapidocr（>=3.9.0，默认 PP-OCRv6）")
        _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE


def _ocr_page(page, dpi: int) -> str:
    """把单页渲染成图片并 OCR，返回识别出的文本（已按版面排好阅读顺序）。"""
    import numpy as np
    pix = page.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:                      # RGBA → RGB
        img = img[:, :, :3]
    elif pix.n == 1:                    # 灰度 → 三通道
        img = np.stack([img[:, :, 0]] * 3, axis=2)
    # 新版 rapidocr 返回 RapidOCROutput（.txts + .boxes）；use_cls=True 保留方向分类器，竖排友好。
    # 关键：用 .boxes 按版面重排阅读顺序（横排多栏按列、竖排按列），否则双栏会被按行交错拼乱。
    result = _get_ocr()(img, use_cls=True)
    if result is None or result.txts is None:
        return ""
    return _ocr_text_from_result(result.txts, result.boxes, pix.width)


# ---------- 繁体转简体 ----------
_ZHCONV = None


def _to_simplified(text: str) -> str:
    """繁体转简体（zhconv，字级转换）。"""
    global _ZHCONV
    if _ZHCONV is None:
        try:
            import zhconv
        except ImportError:
            raise RuntimeError("未安装 zhconv，请先 pip install zhconv")
        _ZHCONV = zhconv
    return _ZHCONV.convert(text, "zh-cn")


@dataclass
class ConvertOptions:
    page_break: bool = False
    unwrap: bool = True
    ocr: bool = False
    ocr_dpi: int = 300          # OCR 标准分辨率，小字/密集古籍更准（200→300）
    ocr_threshold: int = 10
    t2s: bool = False
    layout: str = "auto"   # auto=自动检测横/直排，h=强制横排，v=强制直排


def convert_one(pdf_path: Path, out_path: Path, opts: ConvertOptions) -> tuple[int, int, int]:
    """转换单个 PDF 为 TXT，返回 (页数, 提取到的字符数, 走 OCR 的页数)。

    流程：
    - 先取每页 dict 文本行，做跨页页眉/页脚检测；
    - 逐页判定横排/直排：直排页按位置（竖列）重建阅读顺序并跳过软换行合并，
      横排页走 get_text() + _unwrap_text()；OCR 兜底返回横排文本，按横排处理；
    - 再做繁简转换、按需插页分隔。
    遇到需要密码的加密文档会抛 RuntimeError，由调用方报告并跳过。
    """
    with fitz.open(pdf_path) as doc:
        if doc.needs_pass and not doc.authenticate(""):
            raise RuntimeError("加密 PDF，需要密码")
        n_pages = doc.page_count
        pages = list(doc)
        # 前置：取每页文本行 + 跨页页眉/页脚规则
        per_page_lines = [_extract_lines(p) for p in pages]
        hf = _compute_hf(per_page_lines, n_pages)

        parts: list[str] = []
        total_chars = 0
        n_ocr = 0
        for i, page in enumerate(pages, start=1):
            text = page.get_text() or ""
            ocr_used = False
            if opts.ocr and len(text.strip()) < opts.ocr_threshold:
                try:
                    ocr_text = _ocr_page(page, opts.ocr_dpi)
                    if len(ocr_text) > len(text):
                        text = ocr_text
                        n_ocr += 1
                        ocr_used = True
                except Exception as e:
                    print(f"  [OCR 失败] 第 {i} 页：{e}", file=sys.stderr)

            if opts.layout == "v":
                vertical = True
            elif opts.layout == "h":
                vertical = False
            else:
                vertical = _is_vertical_page(per_page_lines[i - 1])

            if vertical and not ocr_used:
                kept = [ln for ln in per_page_lines[i - 1] if not _is_hf(ln.text, hf)]
                text = _reconstruct_vertical(kept)
                # 竖排重建按列输出（每列一行）；再过一遍软换行合并，把同一段落的多列
                # 拼成连贯段落（仅在大空隙/句末断段），消除“一列一行”的碎片感。
                if opts.unwrap:
                    text = _unwrap_text(text)
            else:
                text = _strip_hf_lines(text, hf)
                if opts.unwrap:
                    text = _unwrap_text(text)

            total_chars += len(text)
            if opts.t2s:
                text = _to_simplified(text)
            if opts.page_break:
                parts.append(f"\n\n===== 第 {i} 页 / 共 {n_pages} 页 =====\n")
            parts.append(text)
            if opts.ocr and (i % 10 == 0 or i == n_pages):
                print(f"  …已处理 {i}/{n_pages} 页", file=sys.stderr)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("".join(parts), encoding="utf-8", newline="\n")
    return n_pages, total_chars, n_ocr


def resolve_inputs(paths: list[str], recursive: bool) -> list[Path]:
    """把命令行传入的路径展开成待转换的 PDF 文件列表。"""
    files: list[Path] = []
    for p in paths:
        pp = Path(p)
        if pp.is_dir():
            glob = "**/*.pdf" if recursive else "*.pdf"
            files.extend(sorted(pp.glob(glob)))
        elif pp.is_file():
            if pp.suffix.lower() == ".pdf":
                files.append(pp)
            else:
                print(f"[跳过] {pp}：不是 .pdf 文件", file=sys.stderr)
        else:
            print(f"[跳过] {pp}：文件不存在", file=sys.stderr)
    return files


def resolve_output(pdf_path: Path, out_arg: str | None) -> Path:
    """根据 -o 参数（可能是文件、目录或 None）决定单个 PDF 的输出路径。

    -o 为空：输出到 PDF 同目录、同名 .txt。
    -o 指向已存在目录，或以路径分隔符结尾：视为目录，输出 该目录/同名.txt。
    -o 其它：当作完整的输出文件路径。
    """
    if out_arg is None:
        return pdf_path.with_suffix(".txt")
    out = Path(out_arg)
    if out.is_dir() or str(out_arg).endswith(("/", "\\")):
        return out / (pdf_path.stem + ".txt")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(
        description="将 PDF 转换为 TXT（文本型直接提取，扫描型可走 OCR）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：python pdf2txt.py 扫描繁体书.pdf --ocr --t2s -o out/",
    )
    ap.add_argument("paths", nargs="+", help="PDF 文件或目录（目录默认递归）")
    ap.add_argument("-o", "--output", help="输出文件，或输出目录（批量转换时用）")
    ap.add_argument("-p", "--page-break", action="store_true",
                    help="每页之间插入“===== 第 N 页 =====”分隔标记")
    ap.add_argument("--unwrap", action=argparse.BooleanOptionalAction, default=True,
                    help="合并段落内的软换行/折行（默认开启）；版面/表单型 PDF 用 --no-unwrap 保留原样")
    ap.add_argument("--layout", choices=["auto", "h", "v"], default="auto",
                    help="版面判定：auto 自动检测横排/直排（默认），h 强制横排，v 强制直排（竖排逐字 PDF 用）")
    ap.add_argument("--ocr", action="store_true",
                    help="对文本过少的页（疑似扫描/图片）启用 OCR 兜底（较慢；需 rapidocr>=3.9.0，默认 PP-OCRv6）")
    ap.add_argument("--ocr-dpi", type=int, default=300,
                    help="OCR 渲染 DPI，越高越准也越慢（默认 300；OCR 标准分辨率）")
    ap.add_argument("--ocr-threshold", type=int, default=10,
                    help="单页文本字符数低于此值才触发 OCR（默认 10）")
    ap.add_argument("--t2s", action="store_true",
                    help="繁体转简体（zhconv，字级转换；需 pip install zhconv）")
    ap.add_argument("--no-recurse", action="store_true",
                    help="传入目录时不递归，只处理顶层 PDF")
    args = ap.parse_args()

    files = resolve_inputs(args.paths, recursive=not args.no_recurse)
    if not files:
        print("没有找到可转换的 PDF。", file=sys.stderr)
        return 1

    # 批量场景下，-o 若指向单一文件名会让结果互相覆盖，提前提醒
    if (len(files) > 1 and args.output
            and not (Path(args.output).is_dir()
                     or str(args.output).endswith(("/", "\\")))):
        print(f"[警告] 共 {len(files)} 个文件，但 -o 指向单一文件名，"
              f"结果会互相覆盖。请用 -o 指定目录（如 -o out/）。", file=sys.stderr)

    opts = ConvertOptions(
        page_break=args.page_break,
        unwrap=args.unwrap,
        ocr=args.ocr,
        ocr_dpi=args.ocr_dpi,
        ocr_threshold=args.ocr_threshold,
        t2s=args.t2s,
        layout=args.layout,
    )

    # 开启 OCR 时，先初始化引擎，避免每个文件重复加载、也早暴露安装问题
    if opts.ocr:
        try:
            _get_ocr()
            print("OCR 引擎已就绪（PP-OCRv6，首次会自动下载模型且推理较慢）。", file=sys.stderr)
        except Exception as e:
            print(f"[错误] OCR 引擎不可用：{e}", file=sys.stderr)
            return 1

    n_ok = n_fail = 0
    for pdf in files:
        out = resolve_output(pdf, args.output)
        try:
            n_pages, n_chars, n_ocr = convert_one(pdf, out, opts)
        except Exception as e:
            print(f"[失败] {pdf}：{e}", file=sys.stderr)
            n_fail += 1
            continue
        ocr_info = f"  其中 {n_ocr} 页走 OCR" if n_ocr else ""
        hint = "  ⚠ 0 字符（可能是扫描件，加 --ocr 再试）" if n_chars == 0 else ""
        print(f"[OK] {pdf.name} -> {out}（{n_pages} 页，{n_chars} 字符）{ocr_info}{hint}")
        n_ok += 1

    print(f"\n完成：成功 {n_ok} 个，失败 {n_fail} 个。", file=sys.stderr)
    return 0 if n_fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
