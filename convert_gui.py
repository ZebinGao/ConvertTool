#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_gui.py — PDF 转 TXT / EPUB 的图形界面（tkinter）。

后端是 pdf_tool.py（复用 pdf2txt.py 的转换逻辑），GUI 负责：
  - 选 PDF、选输出格式（TXT / EPUB）与输出目录；
  - 自动检测「拼合小说」的分界，并允许手动增删改每个区间；
  - 配置 OCR / 版面 / 繁简等选项；
  - 后台线程跑转换（OCR 很慢，不阻塞界面），实时进度条 + 日志，可中断。

运行：
    python convert_gui.py
"""

from __future__ import annotations

import os
import queue
import re
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# 引擎
import pdf_tool
from pdf2txt import ConvertOptions

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


# ------------------------------------------------------------ 工作线程
class ConvertWorker(threading.Thread):
    """后台转换线程。通过 queue 把 (level, text) 日志和进度推给主线程刷新 UI。"""

    def __init__(self, jobs, msg_q: queue.Queue, opts: ConvertOptions,
                 fmt: str, out_dir: Path, workers: int = 0):
        super().__init__(daemon=True)
        self.jobs = jobs                       # list[(title, start, end, author)]
        self.q = msg_q
        self.opts = opts
        self.fmt = fmt
        self.out_dir = out_dir
        self.workers = workers                 # 0 = 自动（cpu-2）
        self._stop = threading.Event()
        self._pool = None

    def stop(self):
        self._stop.set()

    def should_stop(self):
        return self._stop.is_set()

    def run(self):
        pdf_path = self.jobs[0]["pdf"] if self.jobs else None
        try:
            if self.opts.ocr and pdf_path:
                uncached = sum(
                    pdf_tool.uncached_ocr_count(pdf_path, j["start"], j["end"],
                                                self.opts.ocr_dpi, self.opts.ocr_threshold)
                    for j in self.jobs)
                if uncached > 0:
                    self.q.put(("info", f"待 OCR {uncached} 页（每本书各自启动进程池，"
                                        f"命中缓存页秒过）…"))
                else:
                    self.q.put(("info", "全部命中缓存，仅重新走文本清洗…"))
            n_books = len(self.jobs)
            for idx, job in enumerate(self.jobs, start=1):
                if self.should_stop():
                    self.q.put(("warn", "已中断。"))
                    break
                self.q.put(("info", f"▶ [{idx}/{n_books}] 《{job['title']}》"
                                     f" 第 {job['start']}-{job['end']} 页 → {self.fmt.upper()}"))

                book = pdf_tool.BookSpec(
                    title=job["title"], start=job["start"],
                    end=job["end"], author=job.get("author", ""),
                )

                def prog(done, total, msg, _job=job, _idx=idx):
                    pct = int(done / total * 100) if total else 0
                    self.q.put(("progress", (pct, f"[{_idx}/{n_books}] 《{_job['title']}》：{msg}")))

                # 每本书各自创建/关闭进程池（pool=None）：一本的池崩溃不会拖垮下一本，
                # 且全缓存的书根本不启动池。代价是每本多加载一次模型（~8s），可忽略。
                out = pdf_tool.convert_range(
                    pdf_path, book, self.fmt, self.opts, self.out_dir,
                    progress_cb=prog, should_stop=self.should_stop,
                    pool=None, workers=self.workers if self.workers > 0 else None,
                )
                if self.should_stop():
                    self.q.put(("warn", "已中断。"))
                    break
                self.q.put(("ok", f"✔ 《{job['title']}》完成 → {out}"))
            completed = not self.should_stop()
            self.q.put(("done", completed))
        except Exception as e:
            self.q.put(("error", f"转换出错：{e}"))
            self.q.put(("done", False))


class DetectWorker(threading.Thread):
    """后台分界检测线程。"""

    def __init__(self, pdf_path: str, msg_q: queue.Queue):
        super().__init__(daemon=True)
        self.pdf_path = pdf_path
        self.q = msg_q
        self._stop = threading.Event()

    def stop(self):
        self._stop.set()

    def should_stop(self):
        return self._stop.is_set()

    def run(self):
        try:
            n = pdf_tool.page_count(self.pdf_path)

            def prog(done, total, msg):
                pct = int(done / total * 100) if total else 0
                self.q.put(("detect_prog", (pct, msg)))

            sections = pdf_tool.detect_sections(
                self.pdf_path, progress_cb=prog, should_stop=self.should_stop)
            if self.should_stop():
                return
            # 对每本书的起始页 OCR 拿书名（只 OCR 4 页左右，很快）
            titles = []
            for s, _e in sections:
                if self.should_stop():
                    return
                self.q.put(("detect_prog", (None, f"OCR 第 {s} 页识别书名…")))
                t = pdf_tool.guess_title_by_ocr(self.pdf_path, s)
                titles.append(t)
            self.q.put(("detect_done", (sections, titles)))
        except Exception as e:
            self.q.put(("error", f"分界检测出错：{e}"))


# ------------------------------------------------------------ 主窗口
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF → TXT / EPUB 转换工具")
        self.geometry("900x780")
        self.minsize(820, 720)

        self.pdf_path = tk.StringVar()
        self.out_dir = tk.StringVar(value=str(Path.cwd() / "output"))
        self.fmt = tk.StringVar(value="txt")           # txt / epub
        self.split = tk.BooleanVar(value=True)          # 是否拆分多本
        self.ocr = tk.BooleanVar(value=True)
        self.unwrap = tk.BooleanVar(value=True)
        self.t2s = tk.BooleanVar(value=False)
        self.page_break = tk.BooleanVar(value=False)
        self.layout = tk.StringVar(value="auto")
        self.ocr_dpi = tk.IntVar(value=300)
        self.ocr_threshold = tk.IntVar(value=10)
        self.workers = tk.IntVar(value=0)   # 0 = 自动（CPU 核数 - 2）

        self.page_total = 0
        self.book_rows: list[dict] = []     # 动态行控件
        self.msg_q: queue.Queue = queue.Queue()
        self.worker: ConvertWorker | None = None
        self.detector: DetectWorker | None = None

        self._build_ui()
        self.after(150, self._drain_queue)

    # ----------------------------------------------------- UI 构建
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True, padx=10, pady=10)

        # --- 输入 PDF ---
        f1 = ttk.LabelFrame(root, text="① 输入 PDF")
        f1.pack(fill="x", **pad)
        row = ttk.Frame(f1); row.pack(fill="x", padx=8, pady=8)
        ttk.Entry(row, textvariable=self.pdf_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览…", command=self._browse_pdf).pack(side="left", padx=(6, 0))
        self.pdf_info = ttk.Label(row, text="（未选择）", foreground="#666")
        self.pdf_info.pack(side="left", padx=(10, 0))

        # --- 拆分设置 ---
        f2 = ttk.LabelFrame(root, text="② 拆分（多本拼合的小说）")
        f2.pack(fill="x", **pad)
        top = ttk.Frame(f2); top.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Checkbutton(top, text="拆分为多本书",
                        variable=self.split, command=self._on_split_toggle).pack(side="left")
        ttk.Button(top, text="自动检测分界",
                   command=self._start_detect).pack(side="left", padx=(16, 0))
        self.detect_lbl = ttk.Label(top, text="", foreground="#666")
        self.detect_lbl.pack(side="left", padx=(10, 0))

        self.books_frame = ttk.Frame(f2)
        self.books_frame.pack(fill="x", padx=8, pady=(2, 8))
        # 表头
        hdr = ttk.Frame(self.books_frame); hdr.pack(fill="x")
        ttk.Label(hdr, text="书名", width=28).pack(side="left")
        ttk.Label(hdr, text="起始页", width=10).pack(side="left")
        ttk.Label(hdr, text="结束页", width=10).pack(side="left")
        ttk.Label(hdr, text="作者（可选）", width=20).pack(side="left")
        ttk.Label(hdr, text="", width=6).pack(side="left")
        self.rows_host = ttk.Frame(self.books_frame); self.rows_host.pack(fill="x")
        # 初始 1 行
        self._add_book_row()

        # --- 输出 ---
        f3 = ttk.LabelFrame(root, text="③ 输出")
        f3.pack(fill="x", **pad)
        ro = ttk.Frame(f3); ro.pack(fill="x", padx=8, pady=8)
        ttk.Radiobutton(ro, text="TXT 纯文本", variable=self.fmt, value="txt").pack(side="left")
        ttk.Radiobutton(ro, text="EPUB 电子书", variable=self.fmt, value="epub").pack(side="left", padx=(16, 0))
        ro2 = ttk.Frame(f3); ro2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(ro2, text="输出目录：").pack(side="left")
        ttk.Entry(ro2, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(ro2, text="浏览…", command=self._browse_out).pack(side="left")

        # --- 选项 ---
        f4 = ttk.LabelFrame(root, text="④ 转换选项（扫描书请保持 OCR 开启）")
        f4.pack(fill="x", **pad)
        g = ttk.Frame(f4); g.pack(fill="x", padx=8, pady=8)
        ttk.Checkbutton(g, text="OCR 识别（扫描/图片页）", variable=self.ocr).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(g, text="合并软换行（推荐）", variable=self.unwrap).grid(row=0, column=1, sticky="w", padx=20)
        ttk.Checkbutton(g, text="繁体转简体", variable=self.t2s).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(g, text="每页插入分隔标记", variable=self.page_break).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Label(g, text="OCR DPI：").grid(row=1, column=1, sticky="e", pady=(6, 0), padx=20)
        ttk.Spinbox(g, from_=150, to=600, increment=50, width=6,
                    textvariable=self.ocr_dpi).grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Label(g, text="版面：").grid(row=2, column=0, sticky="w", pady=(6, 0))
        for i, (lbl, val) in enumerate([("自动检测", "auto"), ("横排", "h"), ("竖排", "v")]):
            ttk.Radiobutton(g, text=lbl, variable=self.layout, value=val).grid(
                row=2, column=1 + i, sticky="w", pady=(6, 0),
                padx=(0 if i == 0 else 16, 0))
        ttk.Label(g, text="并行进程：").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Spinbox(g, from_=0, to=8, increment=1, width=6,
                    textvariable=self.workers).grid(row=3, column=1, sticky="w", pady=(6, 0))
        ttk.Label(g, text="0=自动（CPU核/2，并行OCR）", foreground="#666").grid(
            row=3, column=2, columnspan=2, sticky="w", pady=(6, 0))

        # --- 进度 + 日志 ---
        f5 = ttk.LabelFrame(root, text="⑤ 进度")
        f5.pack(fill="both", expand=True, **pad)
        pb = ttk.Frame(f5); pb.pack(fill="x", padx=8, pady=(8, 2))
        self.progress = ttk.Progressbar(pb, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)
        self.progress_label = ttk.Label(pb, text="就绪", width=24, anchor="e")
        self.progress_label.pack(side="left", padx=(8, 0))
        self.log = tk.Text(f5, height=8, state="disabled", wrap="word",
                           background="#1e1e1e", foreground="#dcdcdc",
                           font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=8, pady=(2, 8))

        # --- 操作按钮 ---
        btns = ttk.Frame(root); btns.pack(fill="x", **pad)
        self.start_btn = ttk.Button(btns, text="开始转换", command=self._start_convert)
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(btns, text="停止", command=self._stop_convert, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="打开输出目录", command=self._open_out).pack(side="right")

        # 让日志配色
        self.log.tag_config("info", foreground="#9cdcfe")
        self.log.tag_config("ok", foreground="#9ceb94")
        self.log.tag_config("warn", foreground="#ffd166")
        self.log.tag_config("error", foreground="#ff8b8b")

    # ----------------------------------------------------- 书目行管理
    def _add_book_row(self, title="", start="", end="", author=""):
        row = ttk.Frame(self.rows_host); row.pack(fill="x", pady=1)
        t = ttk.Entry(row, width=28); t.insert(0, title); t.pack(side="left")
        s = ttk.Entry(row, width=10); s.insert(0, str(start)); s.pack(side="left", padx=(6, 0))
        e = ttk.Entry(row, width=10); e.insert(0, str(end)); e.pack(side="left", padx=(6, 0))
        a = ttk.Entry(row, width=20); a.insert(0, author); a.pack(side="left", padx=(6, 0))
        rec = {"frame": row, "title": t, "start": s, "end": e, "author": a}
        def del_row(_r=row):
            if len(self.book_rows) <= 1:
                return  # 至少留一行
            _r.destroy()
            self.book_rows = [r for r in self.book_rows if r["frame"] is not _r]
        ttk.Button(row, text="✕", width=3, command=del_row).pack(side="left", padx=(6, 0))
        self.book_rows.append(rec)

    def _on_split_toggle(self):
        state = "normal" if self.split.get() else "disabled"
        for r in self.book_rows:
            for w in (r["title"], r["start"], r["end"], r["author"]):
                w.configure(state=state)

    def _collect_jobs(self) -> list[dict] | None:
        """从界面收集要转的书目。校验失败返回 None 并弹错。"""
        pdf = self.pdf_path.get().strip().strip('"')
        if not pdf or not Path(pdf).is_file():
            messagebox.showerror("错误", "请先选择有效的 PDF 文件。")
            return None
        jobs = []
        if self.split.get():
            for i, r in enumerate(self.book_rows, start=1):
                title = r["title"].get().strip() or f"第{i}部"
                s = r["start"].get().strip()
                e = r["end"].get().strip()
                try:
                    s_i = int(s); e_i = int(e)
                except ValueError:
                    messagebox.showerror("错误", f"《{title}》的起始/结束页必须是数字。")
                    return None
                if s_i < 1 or e_i < s_i or s_i > self.page_total:
                    messagebox.showerror("错误", f"《{title}》的页码范围无效（全书 {self.page_total} 页）。")
                    return None
                e_i = min(e_i, self.page_total)
                jobs.append({"pdf": pdf, "title": title, "start": s_i, "end": e_i,
                             "author": r["author"].get().strip()})
            if not jobs:
                messagebox.showerror("错误", "请至少添加一本书。")
                return None
        else:
            jobs.append({"pdf": pdf, "title": Path(pdf).stem,
                         "start": 1, "end": self.page_total, "author": ""})
        return jobs

    # ----------------------------------------------------- 浏览/打开
    def _browse_pdf(self):
        p = filedialog.askopenfilename(
            title="选择 PDF", filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if p:
            self.pdf_path.set(p)
            self._on_pdf_chosen(p)

    def _on_pdf_chosen(self, p: str):
        try:
            self.page_total = pdf_tool.page_count(p)
            scanned = pdf_tool.is_scanned_pdf(p)
            info = f"{self.page_total} 页" + ("  · 扫描型（需 OCR）" if scanned else "  · 文本型")
            self.pdf_info.configure(text=info)
            self._log("info", f"已载入：{p}（{self.page_total} 页，"
                              f"{'扫描型，建议开 OCR' if scanned else '文本型'}）")
            # 若是扫描型且未填过分界，自动启动检测
            if scanned and not self.book_rows[0]["start"].get().strip():
                self._start_detect()
        except Exception as e:
            self.pdf_info.configure(text="读取失败")
            messagebox.showerror("错误", f"无法读取 PDF：{e}")

    def _browse_out(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.out_dir.set(d)

    def _open_out(self):
        d = self.out_dir.get()
        os.makedirs(d, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(d)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", d])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", d])
        except Exception as e:
            messagebox.showwarning("提示", f"无法打开目录：{e}\n路径：{d}")

    # ----------------------------------------------------- 分界检测
    def _start_detect(self):
        pdf = self.pdf_path.get().strip().strip('"')
        if not pdf or not Path(pdf).is_file():
            messagebox.showerror("错误", "请先选择 PDF。")
            return
        if self.detector and self.detector.is_alive():
            return
        self.detect_lbl.configure(text="检测中…")
        self.detector = DetectWorker(pdf, self.msg_q)
        self.detector.start()

    def _apply_detected(self, sections, titles):
        """把检测结果填入书目行。"""
        # 清空旧行（保留一行再删）
        for r in self.book_rows:
            r["frame"].destroy()
        self.book_rows = []
        n = len(sections)
        for i, (s, e) in enumerate(sections):
            t = titles[i] if i < len(titles) and titles[i] else f"第{i + 1}部"
            # 去掉书名里的换行、多余空白
            t = re.sub(r"\s+", "", t)
            self._add_book_row(title=t, start=s, end=e)
        self._on_split_toggle()
        self.split.set(True)
        self.detect_lbl.configure(text=f"检测到 {n} 本")

    # ----------------------------------------------------- 转换
    def _build_opts(self) -> ConvertOptions:
        return ConvertOptions(
            ocr=self.ocr.get(),
            unwrap=self.unwrap.get(),
            t2s=self.t2s.get(),
            page_break=self.page_break.get(),
            layout=self.layout.get(),
            ocr_dpi=int(self.ocr_dpi.get()),
            ocr_threshold=int(self.ocr_threshold.get()),
        )

    def _start_convert(self):
        jobs = self._collect_jobs()
        if not jobs:
            return
        out_dir = Path(self.out_dir.get())
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出目录：{e}")
            return
        if self.ocr.get() and jobs:
            total_pages = sum(j["end"] - j["start"] + 1 for j in jobs)
            if total_pages > 50:
                w = int(self.workers.get())
                wn = w if w > 0 else max(1, (os.cpu_count() or 4) // 2)
                if not messagebox.askyesno(
                        "确认 OCR",
                        f"将对约 {total_pages} 页执行 OCR（{wn} 进程并行，已缓存页秒过）。\n"
                        f"首次仍需较长时间，但不会再因内存中断。继续？"):
                    return
        opts = self._build_opts()
        self.worker = ConvertWorker(jobs, self.msg_q, opts, self.fmt.get(), out_dir,
                                    workers=int(self.workers.get()))
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._clear_log()
        self._log("info", f"开始转换：{len(jobs)} 本书，格式 {self.fmt.get().upper()}，"
                          f"OCR={'开' if opts.ocr else '关'}")
        self.worker.start()

    def _stop_convert(self):
        if self.worker and self.worker.is_alive():
            self.worker.stop()
            self._log("warn", "已请求停止，正在等待当前页处理完毕…")

    # ----------------------------------------------------- 队列消费
    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == "progress":
                    pct, msg = payload
                    self.progress["value"] = pct
                    self.progress_label.configure(text=f"{pct}%")
                elif kind == "detect_prog":
                    pct, msg = payload
                    self.detect_lbl.configure(text=msg or "检测中…")
                    if pct is not None:
                        self.progress["value"] = pct
                        self.progress_label.configure(text=f"{pct}%")
                elif kind == "detect_done":
                    sections, titles = payload
                    self._apply_detected(sections, titles)
                    self._log("ok", f"分界检测完成：{len(sections)} 本 → "
                                    + ", ".join(
                                        f"{t or '第'+str(i+1)+'部'}({s}-{e})"
                                        for i, ((s, e), t) in enumerate(zip(sections, titles))))
                elif kind in ("info", "ok", "warn", "error"):
                    self._log(kind, payload)
                elif kind == "done":
                    completed = payload
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    if completed:
                        self.progress["value"] = 100
                        self.progress_label.configure(text="100%")
                        self._log("ok", "全部完成 ✔")
                    else:
                        self._log("warn", "已停止。")
        except queue.Empty:
            pass
        self.after(150, self._drain_queue)

    # ----------------------------------------------------- 日志
    def _log(self, level: str, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", level)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.progress["value"] = 0
        self.progress_label.configure(text="0%")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
