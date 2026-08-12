#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_tool.py — PDF 转 TXT / EPUB 的转换引擎（GUI 后端）。

复用 pdf2txt.py 里打磨好的底层逻辑（横排/直排重建、页眉页脚剥离、软换行合并、
OCR、繁简转换），新增：
  - 按「页区间」提取（用来拆分多本拼合的小说）；
  - 逐页 OCR 缓存（扫描书 OCR 很慢，缓存让二次运行近乎免费）；
  - EPUB 生成（纯手写 zip，无 ebooklib 依赖，兼容 Python 3.14）；
  - 章节自动切分（按"第X章/回/节"等标题行）；
  - 基于页面留白的小说分界点检测（无需 OCR，几秒完成）。

对外主入口：
  extract_text(...)        按页区间提取纯文本（带 OCR 缓存）
  convert_range(...)       把一个页区间转成 TXT 或 EPUB
  detect_boundaries(...)   扫描 PDF 找出疑似分界页（封面/扉页）
  guess_title_by_ocr(...)  OCR 指定页拿书名
"""

from __future__ import annotations

import html
import os
import re
import sys
import uuid
import zipfile
from concurrent.futures import ProcessPoolExecutor, wait, FIRST_COMPLETED
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import fitz  # PyMuPDF

# 复用 pdf2txt.py 里现成的转换构件，避免重造轮子
import pdf2txt
from pdf2txt import ConvertOptions

# Windows 控制台统一 UTF-8
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

ProgressCB = Callable[[int, int, str], None]
StopCheck = Callable[[], bool]


# ---------------------------------------------------------------- OCR 缓存
def cache_dir_for(pdf_path: Path | str) -> Path:
    """OCR 逐页缓存目录：<pdf所在目录>/.<文件名去后缀>_ocrcache/"""
    p = Path(pdf_path)
    return p.parent / f".{p.stem}_ocrcache"


def _cache_file(cache_dir: Path, global_page_no: int, dpi: int) -> Path:
    return cache_dir / f"p{global_page_no:05d}_d{dpi}.txt"


def uncached_ocr_count(pdf_path: Path | str, start: int, end: int,
                       dpi: int, ocr_threshold: int = 10) -> int:
    """[start,end] 页里，需要 OCR 但还没缓存的有多少页。

    用来决定要不要启动进程池：若全是缓存命中，就不必启动（省几秒启动 + 几个空闲进程）。
    """
    cache_dir = cache_dir_for(pdf_path)
    n = 0
    with fitz.open(pdf_path) as doc:
        total = doc.page_count
        s = max(1, start); e = min(total, end)
        for p in range(s, e + 1):
            try:
                if len(doc[p - 1].get_text().strip()) >= ocr_threshold:
                    continue  # 有文本层，不会触发 OCR
            except Exception:
                pass
            if not _cache_file(cache_dir, p, dpi).exists():
                n += 1
    return n


# ----------------------------------------------- OCR 文本清洗（扫描书专用）
# 独立成行的页码（"160" / "第N页"）——OCR 返回时常独占一行，否则会被软换行
# 合并进正文，出现 "鹃化175眼" 这种页码夹词中间的粘连。
_PAGE_NUM_LINE = re.compile(r"^\s*\d{1,4}\s*$")
_PAGE_NUM_ZH_LINE = re.compile(r"^\s*第\s*\d+\s*[页頁]\s*$")
# 章节标题行：整行就是"第X章/回/节…"等（长度 ≤12，避免误伤正文里的引用）
_CHAP_LINE = re.compile(
    r"^(?:第[一二三四五六七八九十百千零〇两\d]+[章回节篇卷折幕]"
    r"|序言?|楔子|引子|前言|后记|尾声|跋|序章|终章|番外篇?)$"
)


def _clean_scanned_page(text: str) -> str:
    """单页 OCR 文本的预处理（unwrap 之前调用）：
      - 删掉独立成行的页码 "160" / "第N页"（否则会被软换行合并进正文）；
      - 让章节标题独占段落（前后补空行，否则 "第二十三章" 会和本章正文粘连）。
    """
    out = []
    for ln in text.split("\n"):
        s = ln.strip()
        if s and (_PAGE_NUM_LINE.match(s) or _PAGE_NUM_ZH_LINE.match(s)):
            continue  # 丢弃页码行
        if len(s) <= 12 and _CHAP_LINE.match(s):
            out.append("")      # 标题前空行 → 新段落
            out.append(s)
            out.append("")      # 标题后空行 → 与正文分段
        else:
            out.append(ln)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(out))


def _break_glued_chapters(text: str) -> str:
    """跨页拼接导致的残留粘连：上页结尾的字直接连着下页顶端的章名（如 "…不第二十四章…"）。
    在此类章名前补段落分隔。仅当章名紧跟在非空白字符后触发，已独立成行的标题不受影响。
    """
    return re.sub(
        r"(\S)(第[一二三四五六七八九十百千零〇两\d]{1,6}[章回节篇卷折幕])",
        r"\1\n\n\2",
        text,
    )


# --------------------------------------------------------- 进程池 OCR
# OCR 是 CPU 密集 + ONNX 内存池只涨不缩，单进程跑几百页必爆内存。
# 用进程池：每个 worker OCR 一页后返回文本，重对象（pixmap/numpy/ONNX arena）
# 全留在 worker 进程内，worker 处理 max_tasks_per_child 页后被回收，内存随进程释放。
_POOL_STATE: dict = {}


def _worker_init(pdf_path: str, threads_per_worker: int) -> None:
    """进程池 worker 启动时执行一次：打开 doc、**限制 ONNX 每进程线程数**。

    关键：RapidOCR 的 ONNX 默认 intra_op_num_threads = CPU 核数。若开 N 个进程、每个又用
    全核线程，就 N×核数 个线程抢核数个核（严重超订，实测比单进程还慢）。rapidocr 的 params
    无法设置该值（在三级配置路径下），所以这里 monkeypatch ONNX 的 SessionOptions，强制
    每个 worker 只用 threads_per_worker 个线程 → 总线程数 ≈ worker 数 × 该值 ≤ 物理核数。
    """
    for v in ("OMP_NUM_THREADS", "OPENBLAS_NUM_THREADS", "MKL_NUM_THREADS"):
        os.environ.setdefault(v, str(max(1, threads_per_worker)))
    try:
        import onnxruntime as _ort
        _orig = _ort.SessionOptions.__init__

        def _patched(self, *a, **kw):
            _orig(self, *a, **kw)
            try:
                self.intra_op_num_threads = threads_per_worker
                self.inter_op_num_threads = 1
            except Exception:
                pass

        _ort.SessionOptions.__init__ = _patched
    except Exception:
        pass
    import fitz as _fitz
    _POOL_STATE["doc"] = _fitz.open(pdf_path)
    _POOL_STATE["ocr"] = None  # 懒加载：第一次 OCR 时建


def _worker_get_ocr():
    ocr = _POOL_STATE.get("ocr")
    if ocr is None:
        from rapidocr import RapidOCR
        ocr = RapidOCR()
        _POOL_STATE["ocr"] = ocr
    return ocr


def _ocr_one_page_worker(global_page_no: int, dpi: int, cache_dir: str) -> tuple[int, str]:
    """worker 进程内：渲染并 OCR 单页，结果写缓存，返回 (页码, 文本)。先查缓存。"""
    import numpy as np
    cdir = Path(cache_dir)
    cf = _cache_file(cdir, global_page_no, dpi)
    if cf.exists():
        try:
            return global_page_no, cf.read_text(encoding="utf-8")
        except Exception:
            pass
    doc = _POOL_STATE["doc"]
    page = doc[global_page_no - 1]
    pix = page.get_pixmap(dpi=dpi)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n >= 4:
        arr = arr[:, :, :3]
    elif pix.n == 1:
        arr = np.stack([arr[:, :, 0]] * 3, axis=2)
    result = _worker_get_ocr()(arr, use_cls=True)
    text = "\n".join(result.txts) if (result and result.txts) else ""
    # 释放大数组
    del arr, pix
    try:
        cdir.mkdir(parents=True, exist_ok=True)
        cf.write_text(text, encoding="utf-8")
    except Exception:
        pass
    return global_page_no, text


def make_ocr_pool(pdf_path: Path | str, workers: int | None = None,
                  max_tasks_per_child: int = 15) -> ProcessPoolExecutor:
    """创建一个 OCR 进程池。workers 默认 cpu-2（8 核→6）。

    每个 worker 的 ONNX 线程数 = max(1, cpu // workers)，使「worker 数 × 每worker线程数 ≈ 核数」，
    既并行又不超订。max_tasks_per_child 控制单个 worker 处理多少页后回收（释放 ONNX arena
    内存，是防「扫描书跑几百页后内存爆掉」的关键）。
    """
    cpu = os.cpu_count() or 4
    if workers is None:
        workers = max(1, cpu // 2)   # 实测 4 worker×2 线程 优于 6×1 与 1×8
    workers = max(1, min(int(workers), 8))
    threads_per_worker = max(1, round(cpu / workers))
    kwargs = dict(max_workers=workers, initializer=_worker_init,
                  initargs=(str(pdf_path), threads_per_worker))
    try:
        return ProcessPoolExecutor(max_tasks_per_child=max_tasks_per_child, **kwargs)
    except TypeError:
        # 旧 Python 无 max_tasks_per_child
        return ProcessPoolExecutor(**kwargs)


# -------------------------------------------------------- 按页区间提取文本
def extract_text(
    pdf_path: Path | str,
    start: int,
    end: int,
    opts: ConvertOptions,
    pool: Optional[ProcessPoolExecutor] = None,
    workers: int | None = None,
    progress_cb: Optional[ProgressCB] = None,
    should_stop: Optional[StopCheck] = None,
) -> tuple[str, int, int]:
    """提取 [start, end] 页（1-based 闭区间）的纯文本，返回 (文本, 页数, OCR页数)。

    关键改进（相对 pdf2txt.convert_one）：
      - **并行 OCR**：需要 OCR 的页提交到进程池，重对象留在 worker 进程内、定时回收，
        彻底解决「扫描书跑几百页后内存爆掉」。缓存命中在主进程直接读，不进池。
      - **OCR 文本清洗**：剥离独立成行的页码（防 "鹃化175眼" 粘连）、
        章节标题独占段落（防 "第二十三章君多情人也" 粘连）。
      - 只处理指定页区间；页眉页脚在该区间内统计；可中途 should_stop() 中断。

    pool 为 None 时本函数自建临时池并负责关闭；若调用方要跨多本书复用（省模型加载），
    传入共享池，调用方自行关闭。
    """
    pdf_path = str(pdf_path)
    cache_dir = cache_dir_for(pdf_path)
    with fitz.open(pdf_path) as doc:
        n_total = doc.page_count
        start = max(1, start)
        end = min(n_total, end)
        if end < start:
            return "", 0, 0

        idx0 = list(range(start - 1, end))
        pages = [doc[i] for i in idx0]
        per_page_lines = [pdf2txt._extract_lines(p) for p in pages]
        hf = pdf2txt._compute_hf(per_page_lines, len(pages))
        raw_texts = [p.get_text() or "" for p in pages]

        # 1) 决定哪些页要 OCR
        need_ocr = [i for i, t in enumerate(raw_texts)
                    if opts.ocr and len(t.strip()) < opts.ocr_threshold]

        # 2) 并行 OCR（缓存命中在主进程读，未命中才进池；无未缓存页则不启动池）
        ocr_texts: dict[int, str] = {}
        failed_pages: list[int] = []
        if need_ocr:
            cached_hits = 0
            to_submit: list[tuple[int, int]] = []   # (i, global_no)
            for i in need_ocr:
                global_no = start + i
                cf = _cache_file(cache_dir, global_no, opts.ocr_dpi)
                if cf.exists():
                    try:
                        ocr_texts[i] = cf.read_text(encoding="utf-8")
                        cached_hits += 1
                        continue
                    except Exception:
                        pass
                to_submit.append((i, global_no))
            if cached_hits and progress_cb:
                progress_cb(cached_hits, len(need_ocr), f"命中缓存 {cached_hits} 页")

            if to_submit:
                owns_pool = pool is None
                if owns_pool:
                    pool = make_ocr_pool(pdf_path, workers)
                try:
                    futures: dict = {}
                    for i, global_no in to_submit:
                        try:
                            fut = pool.submit(_ocr_one_page_worker, global_no,
                                              opts.ocr_dpi, str(cache_dir))
                            futures[fut] = (i, global_no)
                        except Exception as e:
                            # 池已坏：剩余页全部标记失败，不挂死
                            ocr_texts[i] = ""
                            failed_pages.append(global_no)
                            if progress_cb:
                                progress_cb(len(ocr_texts), len(need_ocr),
                                            f"进程池不可用：第 {global_no} 页跳过（{e}）")
                    total_fut = len(futures)
                    done = 0
                    pending = set(futures)
                    # 用 wait(超时) 而非 as_completed：循环每 3 秒滴答一次，
                    # 既能响应 should_stop，又能在 worker 崩溃 / 池坏掉时收尾，不会死等。
                    while pending:
                        if should_stop and should_stop():
                            for f in pending:
                                f.cancel()
                            break
                        try:
                            finished, pending = wait(pending, timeout=3.0,
                                                     return_when=FIRST_COMPLETED)
                        except Exception:
                            break
                        if not finished:
                            continue  # 3 秒内无完成，回去重检 stop / 池健康
                        for fut in finished:
                            i, global_no = futures.pop(fut, (None, None))
                            if i is None:
                                continue
                            try:
                                _, text = fut.result()
                                ocr_texts[i] = text
                            except Exception as e:
                                # worker 崩溃 / 池坏 → 该页标记空，记录页码，继续不挂死
                                ocr_texts[i] = ""
                                failed_pages.append(global_no)
                                if progress_cb:
                                    progress_cb(done, total_fut,
                                                f"第 {global_no} 页 OCR 失败（{type(e).__name__}），跳过")
                            done += 1
                            if progress_cb and total_fut:
                                progress_cb(done, total_fut,
                                            f"OCR 第 {global_no}/{n_total} 页（{done}/{total_fut}）")
                finally:
                    if owns_pool:
                        pool.shutdown(wait=False, cancel_futures=True)
                if failed_pages and progress_cb:
                    progress_cb(len(ocr_texts), len(need_ocr),
                                f"⚠ {len(failed_pages)} 页 OCR 失败：{failed_pages[:8]}"
                                + ("…" if len(failed_pages) > 8 else "") + "，重跑可补回")

        # 3) 逐页后处理：组装文本（清洗/unwrap/hf），可中断
        parts: list[str] = []
        n_ocr = len(ocr_texts)
        total = len(pages)
        for i, page in enumerate(pages):
            if should_stop and should_stop():
                break
            global_no = start + i
            text = raw_texts[i]
            ocr_used = False
            if i in ocr_texts:
                ot = ocr_texts[i]
                if len(ot) > len(text):
                    text = ot
                    ocr_used = True

            if opts.layout == "v":
                vertical = True
            elif opts.layout == "h":
                vertical = False
            else:
                vertical = pdf2txt._is_vertical_page(per_page_lines[i])

            if vertical and not ocr_used:
                kept = [ln for ln in per_page_lines[i] if not pdf2txt._is_hf(ln.text, hf)]
                text = pdf2txt._reconstruct_vertical(kept)
                if opts.unwrap:
                    text = pdf2txt._unwrap_text(text)
            else:
                if ocr_used:
                    text = _clean_scanned_page(text)   # 剥页码 + 章名独占段落
                text = pdf2txt._strip_hf_lines(text, hf)
                if opts.unwrap:
                    text = pdf2txt._unwrap_text(text)

            if opts.t2s:
                text = pdf2txt._to_simplified(text)

            if opts.page_break:
                parts.append(f"\n\n===== 第 {global_no} 页 / 共 {n_total} 页 =====\n")
            parts.append(text)

            if progress_cb and not need_ocr:  # 纯文本页才用每页进度（OCR 进度上面已报）
                progress_cb(i + 1, total, f"处理第 {global_no}/{n_total} 页")

        joined = "".join(parts)
        joined = _break_glued_chapters(joined)   # 跨页粘连的章节标题补换行
        return joined, total, n_ocr


# ----------------------------------------------------------- 分界点检测
def _density_cache_path(pdf_path: Path | str) -> Path:
    p = Path(pdf_path)
    return p.parent / f".{p.stem}_density.json"


def _page_densities(
    pdf_path: Path | str,
    progress_cb: Optional[ProgressCB] = None,
    should_stop: Optional[StopCheck] = None,
    dpi: int = 30,
    use_cache: bool = True,
) -> list[float]:
    """逐页「非白像素比例」，返回每页一个值（0~1）。

    结果缓存到 <pdf 同目录>/.<stem>_density.json，二次调用近乎免费
    （JPX 扫描书渲染几百页较慢，首次约 1~2 分钟）。
    """
    import json
    import numpy as np

    p = Path(pdf_path)
    cache = _density_cache_path(p)
    if use_cache and cache.exists():
        try:
            data = json.loads(cache.read_text(encoding="utf-8"))
            if isinstance(data, list) and len(data) >= 1:
                return [float(x) for x in data]
        except Exception:
            pass

    out: list[float] = []
    with fitz.open(pdf_path) as doc:
        n = doc.page_count
        for i in range(n):
            if should_stop and should_stop():
                break
            pix = doc[i].get_pixmap(dpi=dpi)
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            chan = arr[:, :, :3] if pix.n >= 3 else np.stack([arr[:, :, 0]] * 3, axis=2)
            out.append(float((chan.mean(axis=2) < 230).mean()))
            if progress_cb and (i % 25 == 0 or i == n - 1):
                progress_cb(i + 1, n, f"扫描页面 {i + 1}/{n} 定位分界…")
    if use_cache:
        try:
            import json as _json
            cache.write_text(_json.dumps(out), encoding="utf-8")
        except Exception:
            pass
    return out


def detect_sections(
    pdf_path: Path | str,
    progress_cb: Optional[ProgressCB] = None,
    should_stop: Optional[StopCheck] = None,
    dark_threshold: float = 0.02,   # 非白比例低于此 → 视为「非正文页」（封面/扉页/空白）
    min_block: int = 40,             # 正文区块至少多少页才算「一本书」
    max_inner_gap: int = 1,          # 区块内允许的连续空白页数（章间空白=1，小说间≥2）
) -> list[tuple[int, int]]:
    """基于页面留白，把 PDF 切成若干「大段正文区块」——即一本本拼合的小说。

    原理（经验证对扫描版小说合集有效）：
      · 正文页文字密集（非白比例高），封面/扉页/空白页大面积留白（比例 < dark_threshold）；
      · 章节之间的空白通常只有 1 页（max_inner_gap=1 容忍），而两本小说之间
        会有 ≥2 页连续空白（含扉页/版权页）——这就是分界信号；
      1. 标记每页正文/非正文；
      2. 连续正文页合成区块，容忍 ≤ max_inner_gap 页内部空白；
      3. 把每个区块的起始页**回溯**到前面的空白簇——让起始页落在扉页上
        （方便后续 OCR 取书名），多包含 1~3 页空白无害；
      4. 只保留页数 ≥ min_block 的区块（过滤前言、版权页等短碎块），短碎块并入前一本。

    无需 OCR，850 页首次约 1~2 分钟（结果缓存），二次近乎免费。返回 [(起始页, 结束页), ...]。
    """
    densities = _page_densities(pdf_path, progress_cb, should_stop)
    if not densities:
        return [(1, 1)]
    n = len(densities)
    sparse = [d < dark_threshold for d in densities]
    dense = [not s for s in sparse]

    # 合成正文区块：连续正文页推进；连续空白 ≤ max_inner_gap 仍算同区块
    raw: list[tuple[int, int]] = []  # 0-based [start, end]
    i = 0
    while i < n:
        if not dense[i]:
            i += 1
            continue
        start = i
        last_dense = i        # 保证 end ≥ start，绝不死循环
        gap = 0
        j = i + 1
        while j < n:
            if dense[j]:
                gap = 0
                last_dense = j
            else:
                gap += 1
                if gap > max_inner_gap:
                    break
            j += 1
        raw.append((start, last_dense))
        i = last_dense + 1

    # 回溯起始页到前导空白簇（让起始页 = 扉页，便于 OCR 书名）
    def backfill(s: int) -> int:
        while s > 0 and sparse[s - 1]:
            s -= 1
        return s

    # 过滤短区块：把它并到上一个区块（小说尾的版权页之类）
    blocks: list[tuple[int, int]] = []
    for s, e in raw:
        s1 = backfill(s) + 1   # 1-based
        e1 = e + 1
        if (e1 - s1 + 1) >= min_block:
            blocks.append((s1, e1))
        elif blocks:
            ms, me = blocks[-1]
            blocks[-1] = (ms, e1)
    if not blocks:
        blocks = [(1, n)]
    return blocks


def detect_boundaries(
    pdf_path: Path | str,
    progress_cb: Optional[ProgressCB] = None,
    should_stop: Optional[StopCheck] = None,
    **kw,
) -> list[int]:
    """便捷封装：返回 detect_sections 每个区块的起始页（含第 1 页）。

    旧代码兼容；GUI 用 detect_sections 直接拿 (start,end) 区间。
    """
    total = page_count(pdf_path)
    secs = detect_sections(pdf_path, progress_cb, should_stop, **kw)
    return [s for s, _ in secs] + ([total] if secs and secs[-1][1] != total else [])


_TITLE_BAD_END = set("。！？；，、：")
# 作者 / 出版行常见标记词：含这些的行不算书名
_AUTHOR_MARK = ("著", "编", "译", "撰", "出版社", "出版", "印刷", "发行", "书局", "书社")


def _looks_like_title(line: str) -> bool:
    """一行是否像书名：2~10 个字，以中文为主，无句末标点，且不像作者/出版行。"""
    s = line.strip().replace(" ", "")
    if not (2 <= len(s) <= 10):
        return False
    if s[-1] in _TITLE_BAD_END:
        return False
    if any(m in s for m in _AUTHOR_MARK):
        return False
    cjk = sum(1 for c in s if "一" <= c <= "鿿")
    return cjk >= max(2, len(s) - 2)


def guess_title_by_ocr(pdf_path: Path | str, page_no: int, scan: int = 4) -> str:
    """OCR 从 page_no 起的若干页，挑最像书名的短行作为书名候选。

    策略：逐页 OCR，**第一个**出现标题候选行的页就是扉页（扉页总在正文前），
    只在该页的候选里挑——避免翻到正文把章节首句误当书名。扉页上"书名"行比
    "作者"行更长/更显眼，故在去掉作者行后取最长合格行。
    """
    try:
        import numpy as np
        ocr = pdf2txt._get_ocr()
        with fitz.open(pdf_path) as doc:
            n = doc.page_count
            for off in range(scan):
                pno = page_no + off
                if pno > n:
                    break
                page = doc[pno - 1]
                pix = page.get_pixmap(dpi=200)
                arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                if pix.n >= 4:
                    arr = arr[:, :, :3]
                elif pix.n == 1:
                    arr = np.stack([arr[:, :, 0]] * 3, axis=2)
                result = ocr(arr, use_cls=True)
                if not result or not result.txts:
                    continue
                page_cands = [ln.strip().replace(" ", "")
                              for ln in result.txts if _looks_like_title(ln)]
                if page_cands:
                    # 扉页找到了：取最长合格行（作者行已被 _looks_like_title 排除）
                    page_cands.sort(key=len, reverse=True)
                    return page_cands[0]
        return ""
    except Exception:
        return ""


# ----------------------------------------------------------- EPUB 生成
# 章节标题：第X章/回/节/篇/卷/折/幕，或独立成段的序/跋/楔子/引子/前言/后记/尾声
_CHAPTER_RE = re.compile(
    r"^(?:第\s*[一二三四五六七八九十百千零〇两\d]+\s*[章回节篇卷折幕折]"
    r"|^[序跋]|楔子|引子|前言|后记|尾声|序言|序章|终章|番外)"
)
# 逗号、句号结尾的不是标题（正文句）；标题一般短
_MAX_TITLE_LEN = 24


def _is_chapter_heading(para: str) -> bool:
    s = para.strip()
    if not s or len(s) > _MAX_TITLE_LEN * 2:
        return False
    # 去掉首尾标点后判断
    if len(s) > _MAX_TITLE_LEN and not _CHAPTER_RE.match(s):
        return False
    return bool(_CHAPTER_RE.match(s)) and len(s) <= _MAX_TITLE_LEN * 2


def split_chapters(text: str) -> list[tuple[str, list[str]]]:
    """把全文按章节标题切成 [(章节名, [段落...]), ...]。

    段落以空行分隔；遇到疑似章节标题行就开新章。识别不出章节则返回单章。
    """
    paras = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    chapters: list[tuple[str, list[str]]] = []
    cur_title: Optional[str] = None
    cur: list[str] = []

    def flush():
        nonlocal cur_title, cur
        if cur_title is not None or cur:
            chapters.append((cur_title or "正文", cur))
        cur = []

    for p in paras:
        # 去掉页码分隔标记行
        if re.match(r"^=+\s*第\s*\d+\s*页", p):
            continue
        if _is_chapter_heading(p):
            flush()
            cur_title = p.strip().splitlines()[0][:_MAX_TITLE_LEN]
        else:
            cur.append(p)
    flush()

    if not chapters:
        chapters.append(("正文", []))
    # 只有一章且无名 → 调用方会用书名兜底
    return chapters


def _paragraphs_to_html(paras: list[str]) -> str:
    out = []
    for p in paras:
        # 段内换行转 <br/>；整体 XML 转义
        safe = html.escape(p).replace("\n", "<br/>")
        out.append(f"    <p>{safe}</p>")
    return "\n".join(out) or "    <p>　</p>"


_XHTML_TPL = """<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta charset="utf-8"/>
  <title>{title}</title>
  <link rel="stylesheet" type="text/css" href="style.css"/>
</head>
<body>
  <h2>{heading}</h2>
{body}
</body>
</html>"""

_CSS = """body { font-family: "Noto Serif CJK SC", "Source Han Serif SC", "Songti SC", serif;
         line-height: 1.8; margin: 5%; }
h2 { text-align: center; margin: 1.5em 0; }
p { text-indent: 2em; margin: 0.3em 0; text-align: justify; }"""

_CONTAINER_XML = """<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>"""


def _build_opf(book_title: str, author: str, uid: str, dt: str,
               n_chapters: int, lang: str) -> str:
    manifest = "\n    ".join(
        f'<item id="ch{i}" href="chapter{i}.xhtml" media-type="application/xhtml+xml"/>'
        for i in range(1, n_chapters + 1)
    )
    manifest += f'\n    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>'
    manifest += f'\n    <item id="css" href="style.css" media-type="text/css"/>'
    spine = "\n    ".join(f'<itemref idref="ch{i}"/>' for i in range(1, n_chapters + 1))
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="2.0" unique-identifier="bookid">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf">
    <dc:identifier id="bookid">{uid}</dc:identifier>
    <dc:title>{html.escape(book_title)}</dc:title>
    <dc:creator opf:role="aut">{html.escape(author or "佚名")}</dc:creator>
    <dc:language>{lang}</dc:language>
    <dc:date>{dt}</dc:date>
  </metadata>
  <manifest>
    {manifest}
  </manifest>
  <spine toc="ncx">
    {spine}
  </spine>
</package>"""


def _build_ncx(book_title: str, author: str, uid: str,
               chapters: list[tuple[str, list[str]]]) -> str:
    points = []
    for i, (title, _) in enumerate(chapters, start=1):
        points.append(f"""    <navPoint id="nav{i}" playOrder="{i}">
      <navLabel><text>{html.escape(title)}</text></navLabel>
      <content src="chapter{i}.xhtml"/>
    </navPoint>""")
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
  <head>
    <meta name="dtb:uid" content="{uid}"/>
    <meta name="dtb:depth" content="1"/>
    <meta name="dtb:totalPageCount" content="0"/>
    <meta name="dtb:maxPageNumber" content="0"/>
  </head>
  <docTitle><text>{html.escape(book_title)}</text></docTitle>
  <docAuthor><text>{html.escape(author or "佚名")}</text></docAuthor>
  <navMap>
{chr(10).join(points)}
  </navMap>
</ncx>"""


def write_epub(out_path: Path | str, book_title: str, author: str,
               text: str, lang: str = "zh-CN") -> int:
    """把全文写成 EPUB（手写 zip）。返回章节数。"""
    chapters = split_chapters(text)
    # 第一章无名且只有一章 → 用书名
    if len(chapters) == 1 and (not chapters[0][0] or chapters[0][0] == "正文"):
        chapters = [(book_title, chapters[0][1])]
    # 章节太多（>400）说明误识别，退回单章
    if len(chapters) > 400:
        paras = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
        chapters = [(book_title, paras)]

    uid = "urn:uuid:" + str(uuid.uuid4())
    dt = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    files: list[tuple[str, bytes, int]] = []
    # mimetype 必须第一个且不压缩
    files.append(("mimetype", b"application/epub+zip", zipfile.ZIP_STORED))
    files.append(("META-INF/container.xml", _CONTAINER_XML.encode("utf-8"), zipfile.ZIP_DEFLATED))
    files.append(("OEBPS/style.css", _CSS.encode("utf-8"), zipfile.ZIP_DEFLATED))
    files.append(("OEBPS/content.opf",
                  _build_opf(book_title, author, uid, dt, len(chapters), lang).encode("utf-8"),
                  zipfile.ZIP_DEFLATED))
    files.append(("OEBPS/toc.ncx",
                  _build_ncx(book_title, author, uid, chapters).encode("utf-8"),
                  zipfile.ZIP_DEFLATED))
    for i, (title, paras) in enumerate(chapters, start=1):
        xhtml = _XHTML_TPL.format(
            title=html.escape(title),
            heading=html.escape(title),
            body=_paragraphs_to_html(paras),
        )
        files.append((f"OEBPS/chapter{i}.xhtml", xhtml.encode("utf-8"), zipfile.ZIP_DEFLATED))

    with zipfile.ZipFile(out, "w") as z:
        for name, data, comp in files:
            z.writestr(name, data, compress_type=comp)
    return len(chapters)


def write_txt(out_path: Path | str, text: str) -> None:
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(text, encoding="utf-8", newline="\n")


# ----------------------------------------------------------- 高层入口
@dataclass
class BookSpec:
    """一本要转的书：书名 + 页区间 + 作者。"""
    title: str
    start: int
    end: int
    author: str = ""


def convert_range(
    pdf_path: Path | str,
    book: BookSpec,
    fmt: str,                      # "txt" 或 "epub"
    opts: ConvertOptions,
    out_dir: Path | str,
    progress_cb: Optional[ProgressCB] = None,
    should_stop: Optional[StopCheck] = None,
    pool: Optional[ProcessPoolExecutor] = None,
    workers: int | None = None,
) -> Path:
    """转一本书（一个页区间）为 TXT 或 EPUB，返回输出文件路径。"""
    out_dir = Path(out_dir)
    safe_title = re.sub(r'[\\/:*?"<>|]', "_", book.title.strip()) or "output"
    text, n_pages, n_ocr = extract_text(
        pdf_path, book.start, book.end, opts, pool, workers, progress_cb, should_stop
    )
    if fmt.lower() == "epub":
        out = out_dir / f"{safe_title}.epub"
        write_epub(out, book.title, book.author, text)
    else:
        out = out_dir / f"{safe_title}.txt"
        write_txt(out, text)
    return out


def page_count(pdf_path: Path | str) -> int:
    with fitz.open(pdf_path) as doc:
        return doc.page_count


def is_scanned_pdf(pdf_path: Path | str, sample: int = 10) -> bool:
    """抽样判断是否扫描型 PDF（几乎无文本层）。"""
    with fitz.open(pdf_path) as doc:
        n = doc.page_count
        step = max(1, n // sample)
        chars = 0
        checked = 0
        for i in range(0, n, step):
            chars += len(doc[i].get_text().strip())
            checked += 1
            if checked >= sample:
                break
        return chars < checked * 20  # 平均每页 <20 字符 → 判为扫描
