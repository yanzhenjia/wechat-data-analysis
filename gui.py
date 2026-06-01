from __future__ import annotations

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from gzh_analyzer import AnalysisOutputs, choose_default_folder, run_analysis


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("公众号文章流量分析工具")
        self.geometry("780x520")
        self.minsize(780, 520)

        self.input_var = tk.StringVar(value=str(choose_default_folder("data")))
        self.output_var = tk.StringVar(value=str(choose_default_folder("output")))
        self.status_var = tk.StringVar(value="选择数据目录后，点击“开始分析”。")

        self.last_outputs: AnalysisOutputs | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        wrapper = ttk.Frame(self, padding=14)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(
            wrapper,
            text="公众号文章流量整合与分析",
            font=("Microsoft YaHei UI", 15, "bold"),
        ).pack(anchor="w", pady=(0, 12))

        tip = ttk.Label(
            wrapper,
            text="自动去重导入文章数据，同步输出 Excel 统计表和 HTML 可视化报告。",
            foreground="#5d6b73",
        )
        tip.pack(anchor="w", pady=(0, 14))

        self._build_path_row(wrapper, "数据目录", self.input_var, self._choose_input)
        self._build_path_row(wrapper, "输出目录", self.output_var, self._choose_output)

        action_row = ttk.Frame(wrapper)
        action_row.pack(fill="x", pady=(6, 12))

        self.run_btn = ttk.Button(action_row, text="开始分析", command=self._run)
        self.run_btn.pack(side="left")

        self.open_output_btn = ttk.Button(action_row, text="打开输出目录", command=self._open_output_dir)
        self.open_output_btn.pack(side="left", padx=(10, 0))

        self.open_excel_btn = ttk.Button(
            action_row, text="打开 Excel", command=self._open_excel, state="disabled"
        )
        self.open_excel_btn.pack(side="left", padx=(10, 0))

        self.open_html_btn = ttk.Button(
            action_row, text="打开 HTML 报告", command=self._open_html, state="disabled"
        )
        self.open_html_btn.pack(side="left", padx=(10, 0))

        log_frame = ttk.Frame(wrapper)
        log_frame.pack(fill="both", expand=True)
        ttk.Label(log_frame, text="运行日志").pack(anchor="w")

        self.log = tk.Text(log_frame, height=14, wrap="word")
        self.log.pack(fill="both", expand=True, pady=(6, 0))

        status = ttk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken")
        status.pack(side="bottom", fill="x")

    def _build_path_row(self, parent: ttk.Frame, label: str, variable: tk.StringVar, callback) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, width=10).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="浏览...", command=callback).pack(side="left")

    def _append_log(self, message: str) -> None:
        self.log.insert("end", message + "\n")
        self.log.see("end")

    def _set_busy(self, busy: bool) -> None:
        self.run_btn.config(state="disabled" if busy else "normal")

    def _choose_input(self) -> None:
        selected = filedialog.askdirectory(title="选择数据目录", initialdir=self.input_var.get() or str(choose_default_folder("data")))
        if selected:
            self.input_var.set(selected)

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(title="选择输出目录", initialdir=self.output_var.get() or str(choose_default_folder("output")))
        if selected:
            self.output_var.set(selected)

    def _run(self) -> None:
        input_dir = Path(self.input_var.get()).expanduser().resolve()
        output_dir = Path(self.output_var.get()).expanduser().resolve()

        if not input_dir.exists() or not input_dir.is_dir():
            messagebox.showerror("输入目录错误", f"输入目录不存在：\n{input_dir}")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        self.last_outputs = None
        self._set_busy(True)
        self.open_excel_btn.config(state="disabled")
        self.open_html_btn.config(state="disabled")
        self._append_log("=" * 60)
        self._append_log(f"输入目录：{input_dir}")
        self._append_log(f"输出目录：{output_dir}")
        self.status_var.set("正在分析，请稍候...")

        def worker() -> None:
            try:
                outputs = run_analysis(input_dir, output_dir)
                self.last_outputs = outputs

                def done() -> None:
                    self._append_log(f"Excel：{outputs.excel_path}")
                    self._append_log(f"HTML： {outputs.html_path}")
                    self._append_log(
                        f"扫描 {outputs.total_files} 个文件，导入 {outputs.imported_articles} 个，"
                        f"跳过重复 {outputs.skipped_duplicates} 个，失败 {outputs.failed_files} 个。"
                    )
                    self.status_var.set("分析完成。现在可以直接打开 Excel 或 HTML 报告。")
                    self.open_excel_btn.config(state="normal")
                    self.open_html_btn.config(state="normal")
                    self._set_busy(False)

                self.after(0, done)
            except Exception as exc:
                def fail() -> None:
                    self._append_log(f"[错误] {exc}")
                    self.status_var.set("分析失败，请查看日志。")
                    self._set_busy(False)
                    messagebox.showerror("分析失败", str(exc))

                self.after(0, fail)

        threading.Thread(target=worker, daemon=True).start()

    def _open_output_dir(self) -> None:
        output_dir = Path(self.output_var.get()).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        os.startfile(str(output_dir))  # type: ignore[attr-defined]

    def _open_excel(self) -> None:
        if self.last_outputs and self.last_outputs.excel_path.exists():
            os.startfile(str(self.last_outputs.excel_path))  # type: ignore[attr-defined]

    def _open_html(self) -> None:
        if self.last_outputs and self.last_outputs.html_path.exists():
            os.startfile(str(self.last_outputs.html_path))  # type: ignore[attr-defined]


def main() -> None:
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # type: ignore[attr-defined]
    except Exception:
        pass

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
