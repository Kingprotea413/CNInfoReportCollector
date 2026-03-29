from __future__ import annotations

import argparse
import json
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from cninfo_pipeline import AVAILABLE_UNIT_LABELS, DEFAULT_UNIT_LABEL, AnnualReportPipeline


APP_ID = "CNInfoReportCollector"
CONFIG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / APP_ID
CONFIG_PATH = CONFIG_DIR / "config.json"
DEFAULT_OUTPUT_DIR = Path.home() / "Desktop" / "财报输出"


def resolve_unit_label(candidate: str | None) -> str:
    return candidate if candidate in AVAILABLE_UNIT_LABELS else DEFAULT_UNIT_LABEL


def load_settings() -> dict[str, str]:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}
    return {}


def save_settings(settings: dict[str, str]) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


class AppWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.pipeline = AnnualReportPipeline()
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.settings = load_settings()
        self.output_dir = Path(self.settings.get("output_dir", str(DEFAULT_OUTPUT_DIR)))
        self.unit_label = resolve_unit_label(self.settings.get("unit_label"))

        self.root.title("巨潮年报资产负债表采集器")
        self.root.geometry("680x360")
        self.root.resizable(False, False)

        container = ttk.Frame(self.root, padding=20)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="输入公司名称或证券代码，采集多年年报资产负债表并导出 Excel。",
        ).pack(anchor="w")

        company_row = ttk.Frame(container)
        company_row.pack(fill="x", pady=(18, 10))

        ttk.Label(company_row, text="公司").pack(side="left")
        self.company_var = tk.StringVar(value="长江电力")
        self.entry = ttk.Entry(company_row, textvariable=self.company_var, width=36)
        self.entry.pack(side="left", padx=(10, 10), fill="x", expand=True)

        self.start_button = ttk.Button(company_row, text="开始采集", command=self.start_collection)
        self.start_button.pack(side="left")

        output_row = ttk.Frame(container)
        output_row.pack(fill="x", pady=(2, 10))

        ttk.Label(output_row, text="保存位置").pack(side="left")
        self.output_dir_var = tk.StringVar(value=str(self.output_dir))
        output_entry = ttk.Entry(output_row, textvariable=self.output_dir_var, state="readonly", width=58)
        output_entry.pack(side="left", padx=(10, 10), fill="x", expand=True)
        ttk.Button(output_row, text="选择文件夹", command=self.choose_output_dir).pack(side="left")

        unit_row = ttk.Frame(container)
        unit_row.pack(fill="x", pady=(2, 10))

        ttk.Label(unit_row, text="导出单位").pack(side="left")
        self.unit_var = tk.StringVar(value=self.unit_label)
        self.unit_combobox = ttk.Combobox(
            unit_row,
            textvariable=self.unit_var,
            state="readonly",
            values=AVAILABLE_UNIT_LABELS,
            width=12,
        )
        self.unit_combobox.pack(side="left", padx=(10, 0))
        self.unit_combobox.bind("<<ComboboxSelected>>", self.on_unit_selected)

        self.progress = ttk.Progressbar(container, length=620, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(14, 8))

        self.status_var = tk.StringVar(value="等待开始。")
        ttk.Label(container, textvariable=self.status_var, wraplength=620).pack(anchor="w")

        self.result_var = tk.StringVar(value=f"默认保存目录：{self.output_dir}")
        ttk.Label(container, textvariable=self.result_var, wraplength=620, foreground="#555555").pack(
            anchor="w",
            pady=(10, 0),
        )

        self.root.after(150, self.poll_events)

    def on_unit_selected(self, _event: object | None = None) -> None:
        self.unit_label = resolve_unit_label(self.unit_var.get())
        self.unit_var.set(self.unit_label)
        self.settings["unit_label"] = self.unit_label
        save_settings(self.settings)

    def choose_output_dir(self) -> None:
        initial_dir = self.output_dir if self.output_dir.exists() else DEFAULT_OUTPUT_DIR
        selected = filedialog.askdirectory(
            title="选择财报保存位置",
            initialdir=str(initial_dir),
            mustexist=False,
        )
        if not selected:
            return

        self.output_dir = Path(selected)
        self.output_dir_var.set(str(self.output_dir))
        self.result_var.set(f"默认保存目录：{self.output_dir}")
        self.settings["output_dir"] = str(self.output_dir)
        save_settings(self.settings)

    def start_collection(self) -> None:
        if self.worker and self.worker.is_alive():
            return

        company = self.company_var.get().strip()
        if not company:
            messagebox.showerror("输入错误", "请输入公司名称或证券代码。")
            return

        self.progress["value"] = 0
        self.status_var.set("任务已启动。")
        self.on_unit_selected()
        self.result_var.set(f"正在保存到：{self.output_dir}（单位：{self.unit_label}）")
        self.start_button.config(state="disabled")
        self.entry.config(state="disabled")
        self.unit_combobox.config(state="disabled")

        self.worker = threading.Thread(
            target=self._run_pipeline,
            args=(company, str(self.output_dir), self.unit_label),
            daemon=True,
        )
        self.worker.start()

    def _run_pipeline(self, company: str, output_dir: str, unit_label: str) -> None:
        try:
            result = self.pipeline.run(
                company_query=company,
                output_dir=output_dir,
                unit_label=unit_label,
                progress=self._publish_progress,
            )
            self.events.put(("success", result))
        except Exception as exc:  # noqa: BLE001
            self.events.put(("error", str(exc)))

    def _publish_progress(self, percent: int, message: str) -> None:
        self.events.put(("progress", (percent, message)))

    def poll_events(self) -> None:
        while True:
            try:
                event_type, payload = self.events.get_nowait()
            except queue.Empty:
                break

            if event_type == "progress":
                percent, message = payload  # type: ignore[misc]
                self.progress["value"] = percent
                self.status_var.set(message)
            elif event_type == "success":
                result = payload
                self.progress["value"] = 100
                self.status_var.set(
                    f"{result.company.secname} 采集完成，共导出 {result.annual_records} 份年报。"
                )
                self.result_var.set(
                    f"导出文件：{result.output_path.resolve()}（单位：{result.unit_label}）"
                )
                messagebox.showinfo("采集完成", f"Excel 已生成：\n{result.output_path.resolve()}")
                self.start_button.config(state="normal")
                self.entry.config(state="normal")
                self.unit_combobox.config(state="readonly")
            elif event_type == "error":
                self.status_var.set("采集失败。")
                messagebox.showerror("采集失败", str(payload))
                self.start_button.config(state="normal")
                self.entry.config(state="normal")
                self.unit_combobox.config(state="readonly")

        self.root.after(150, self.poll_events)


def run_headless(company: str, output_dir: str, unit_label: str) -> int:
    pipeline = AnnualReportPipeline()

    def reporter(percent: int, message: str) -> None:
        print(f"[{percent:>3}%] {message}")

    result = pipeline.run(
        company_query=company,
        output_dir=output_dir,
        unit_label=unit_label,
        progress=reporter,
    )
    print(
        "导出完成："
        f"{result.company.secname}({result.company.seccode}) -> {result.output_path.resolve()}"
        f" [单位：{result.unit_label}]"
    )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="巨潮资讯年报资产负债表采集器")
    parser.add_argument("--company", default="长江电力", help="公司名称或证券代码")
    parser.add_argument("--output-dir", default="outputs", help="Excel 输出目录")
    parser.add_argument(
        "--unit",
        default=DEFAULT_UNIT_LABEL,
        choices=AVAILABLE_UNIT_LABELS,
        help="导出单位",
    )
    parser.add_argument("--headless", action="store_true", help="不启动窗口，直接命令行执行")
    parser.add_argument(
        "--self-test-gui",
        action="store_true",
        help="仅用于测试窗口能否创建，窗口会自动关闭",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.headless:
        return run_headless(args.company, args.output_dir, args.unit)

    root = tk.Tk()
    AppWindow(root)
    if args.self_test_gui:
        root.after(1200, root.destroy)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
