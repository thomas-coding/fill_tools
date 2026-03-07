from __future__ import annotations

import ctypes
import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app_engine import (
    FILL_PROFILE_PATROL1,
    FILL_PROFILE_PATROL2,
    SessionBuildResult,
    build_session,
    cleanup_session_dir,
    is_supported_excel,
    progress_stats,
    sync_progress_to_source,
    write_runtime_ahk,
)


APP_NAME = "盛丹的小工具"
APP_VERSION = "V0.2"


def resource_root() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent


def set_windows_app_id() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ShengDan.Tool.V0.2")
    except Exception:
        pass


def find_ahk_executable() -> Path:
    root = resource_root()
    candidates = [
        root / "AutoHotkey-v2" / "AutoHotkey64.exe",
        root / "AutoHotkey-v2" / "AutoHotkey32.exe",
        Path(__file__).resolve().parent / "AutoHotkey-v2" / "AutoHotkey64.exe",
        Path(__file__).resolve().parent / "AutoHotkey-v2" / "AutoHotkey32.exe",
    ]
    for path in candidates:
        if path.exists():
            return path
    raise FileNotFoundError("未找到 AutoHotkey 运行时，请检查打包文件")


class PatrolAssistantApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} {APP_VERSION}")
        self.root.geometry("860x560")
        self.root.minsize(760, 520)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._try_set_window_icon()

        self.file_var = tk.StringVar()
        self.run_state_var = tk.StringVar(value="未运行")
        self.progress_var = tk.StringVar(value="0 / 0")
        self.fill_profile_var = tk.StringVar(value=FILL_PROFILE_PATROL1)
        self.tips_var = tk.StringVar()
        self._icon_photo = None

        self.session: Optional[SessionBuildResult] = None
        self.ahk_process: Optional[subprocess.Popen] = None
        self.monitor_thread: Optional[threading.Thread] = None
        self.monitor_stop = threading.Event()
        self.ui_queue: queue.Queue = queue.Queue()
        self.last_progress_mtime = 0.0

        self._build_ui()
        self.root.after(200, self._drain_queue)

    def _try_set_window_icon(self) -> None:
        ico_candidates = [
            resource_root() / "app_icon.ico",
            Path(__file__).resolve().parent / "app_icon.ico",
        ]
        for icon_path in ico_candidates:
            if not icon_path.exists():
                continue
            try:
                self.root.iconbitmap(str(icon_path))
                break
            except Exception:
                continue

        png_candidates = [
            resource_root() / "app_icon_preview.png",
            Path(__file__).resolve().parent / "app_icon_preview.png",
        ]
        for png_path in png_candidates:
            if not png_path.exists():
                continue
            try:
                self._icon_photo = tk.PhotoImage(file=str(png_path))
                self.root.iconphoto(True, self._icon_photo)
                break
            except Exception:
                continue

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        title_row = ttk.Frame(frame)
        title_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(title_row, text=APP_NAME, font=("Microsoft YaHei UI", 16, "bold")).pack(side=tk.LEFT)
        ttk.Label(title_row, text=APP_VERSION, foreground="#A35E00").pack(side=tk.LEFT, padx=(8, 0))

        module_frame = ttk.LabelFrame(frame, text="功能选择", padding=8)
        module_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Radiobutton(
            module_frame,
            text="巡检填报1（默认）",
            variable=self.fill_profile_var,
            value=FILL_PROFILE_PATROL1,
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            module_frame,
            text="巡检填报2",
            variable=self.fill_profile_var,
            value=FILL_PROFILE_PATROL2,
        ).pack(side=tk.LEFT, padx=(12, 0))

        self.fill_profile_var.trace_add("write", lambda *_: self._refresh_tips_text())

        file_row = ttk.Frame(frame)
        file_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(file_row, text="Excel文件:").pack(side=tk.LEFT)
        entry = ttk.Entry(file_row, textvariable=self.file_var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        ttk.Button(file_row, text="选择文件", command=self.select_file).pack(side=tk.LEFT)

        action_row = ttk.Frame(frame)
        action_row.pack(fill=tk.X, pady=(0, 8))
        self.run_btn = ttk.Button(action_row, text="运行", command=self.start_run)
        self.run_btn.pack(side=tk.LEFT)
        self.stop_btn = ttk.Button(action_row, text="停止", command=self.stop_run, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(8, 0))

        status_row = ttk.Frame(frame)
        status_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(status_row, text="状态:").pack(side=tk.LEFT)
        ttk.Label(status_row, textvariable=self.run_state_var).pack(side=tk.LEFT, padx=(6, 20))
        ttk.Label(status_row, text="进度:").pack(side=tk.LEFT)
        ttk.Label(status_row, textvariable=self.progress_var).pack(side=tk.LEFT, padx=(6, 0))

        self._refresh_tips_text()
        ttk.Label(frame, textvariable=self.tips_var, foreground="#444").pack(anchor=tk.W, pady=(0, 6))

        log_frame = ttk.LabelFrame(frame, text="运行日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.configure(state=tk.DISABLED)

    def log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def select_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择巡检Excel文件",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xltx *.xltm *.xls *.xlsb *.excel"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.file_var.set(path)

    def _refresh_tips_text(self) -> None:
        profile = self.fill_profile_var.get().strip() or FILL_PROFILE_PATROL1
        if profile == FILL_PROFILE_PATROL2:
            tips = "操作提示: 巡检填报2在“处置路段”按 F5 填2项（处置路段/整改描述）；文件框按 F12，提交后按 F10。"
        else:
            tips = "操作提示: 巡检填报1在“问题地址”按 F5 填4项（地址/路段/截止/描述）；文件框按 F12，提交后按 F10。"
        self.tips_var.set(tips)

    def _count_missing_photo_rows(self) -> tuple[int, list[str]]:
        if not self.session:
            return 0, []

        data_path = self.session.paths.data_tsv
        if not data_path.exists():
            return 0, []

        missing_rows: list[str] = []
        lines = data_path.read_text(encoding="utf-8").splitlines()
        for line in lines[1:]:
            if not line.strip():
                continue

            cols = line.split("\t")
            while len(cols) < 7:
                cols.append("")

            source_row = cols[0].strip()
            photo_path = cols[6].strip()
            if not photo_path:
                missing_rows.append(source_row or "?")

        return len(missing_rows), missing_rows

    def start_run(self) -> None:
        if self.ahk_process and self.ahk_process.poll() is None:
            messagebox.showinfo("提示", "程序已经在运行中")
            return

        fill_profile = self.fill_profile_var.get().strip() or FILL_PROFILE_PATROL1
        if fill_profile not in {FILL_PROFILE_PATROL1, FILL_PROFILE_PATROL2}:
            messagebox.showwarning("提示", "请选择有效的填报功能")
            return

        feature_label = "巡检填报2" if fill_profile == FILL_PROFILE_PATROL2 else "巡检填报1"

        raw = self.file_var.get().strip()
        if not raw:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return

        source = Path(raw)
        if not source.exists():
            messagebox.showerror("错误", "所选文件不存在")
            return
        if not is_supported_excel(source):
            messagebox.showerror("错误", "不支持该文件格式，请选择 xlsx/xlsm/xls/xlsb/.excel")
            return

        try:
            self.log("正在加载 Excel 并解析图片，请稍候...")
            self.root.update_idletasks()
            self.session = build_session(source, fill_profile=fill_profile)
            write_runtime_ahk(self.session.paths, self.session.fill_profile)
            progress_path = self.session.paths.progress_tsv
            if progress_path.exists():
                self.last_progress_mtime = progress_path.stat().st_mtime
            self._launch_ahk()
            self._start_monitor()
            self._update_progress_label()

            self.run_state_var.set("运行中")
            self.run_btn.configure(state=tk.DISABLED)
            self.stop_btn.configure(state=tk.NORMAL)

            self.log("Excel加载完成，可开始按 F5/F12/F10 进行填报")
            self.log(f"当前功能: {feature_label}")
            self.log(f"已加载文件: {self.session.source_excel}")
            self.log(f"识别模式: {self.session.mode}，记录数: {self.session.total_records}")

            missing_count, missing_rows = self._count_missing_photo_rows()
            if missing_count:
                preview = ",".join(missing_rows[:8])
                suffix = "..." if missing_count > 8 else ""
                self.log(f"注意: {missing_count} 条记录未解析到上传照片（源行: {preview}{suffix}），这些行按 F12 会提示路径为空")

            self.log("热键已生效：F5/F12/F10")
        except Exception as exc:
            if self.session:
                cleanup_session_dir(self.session.paths.session_dir)
                self.session = None
            messagebox.showerror("启动失败", str(exc))

    def _launch_ahk(self) -> None:
        assert self.session is not None
        ahk_exe = find_ahk_executable()
        script = self.session.paths.ahk_script

        creationflags = 0
        if sys.platform == "win32":
            creationflags = subprocess.CREATE_NO_WINDOW

        self.ahk_process = subprocess.Popen(
            [str(ahk_exe), str(script)],
            cwd=str(self.session.paths.session_dir),
            creationflags=creationflags,
        )

    def _start_monitor(self) -> None:
        self.monitor_stop.clear()

        def _loop() -> None:
            assert self.session is not None
            progress_path = self.session.paths.progress_tsv
            meta_path = self.session.paths.meta_json

            while not self.monitor_stop.wait(0.15):
                proc = self.ahk_process
                if proc is not None and proc.poll() is not None:
                    self.ui_queue.put(("ahk_exit", proc.returncode))
                    break

                if not progress_path.exists():
                    continue

                try:
                    mtime = progress_path.stat().st_mtime
                except FileNotFoundError:
                    continue

                if mtime <= self.last_progress_mtime:
                    continue

                self.last_progress_mtime = mtime
                try:
                    updated = sync_progress_to_source(meta_path, progress_path)
                    done, total = progress_stats(progress_path)
                    self.ui_queue.put(("progress", done, total, updated))
                except Exception as exc:
                    self.ui_queue.put(("error", f"回写异常: {exc}"))

        self.monitor_thread = threading.Thread(target=_loop, daemon=True)
        self.monitor_thread.start()

    def stop_run(self) -> None:
        self.monitor_stop.set()

        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=0.4)
        self.monitor_thread = None

        if self.ahk_process and self.ahk_process.poll() is None:
            self.ahk_process.terminate()
            try:
                self.ahk_process.wait(timeout=0.6)
            except subprocess.TimeoutExpired:
                self.ahk_process.kill()
        self.ahk_process = None

        self.run_state_var.set("已停止")
        self.run_btn.configure(state=tk.NORMAL)
        self.stop_btn.configure(state=tk.DISABLED)

    def _update_progress_label(self) -> None:
        if not self.session:
            self.progress_var.set("0 / 0")
            return

        done, total = progress_stats(self.session.paths.progress_tsv)
        self.progress_var.set(f"{done} / {total}")

    def _drain_queue(self) -> None:
        while True:
            try:
                item = self.ui_queue.get_nowait()
            except queue.Empty:
                break

            kind = item[0]
            if kind == "progress":
                _, done, total, updated = item
                self.progress_var.set(f"{done} / {total}")
                self.log(f"已回写Excel: {updated} 行，当前进度 {done}/{total}")
            elif kind == "error":
                _, msg = item
                self.log(msg)
            elif kind == "ahk_exit":
                _, code = item
                self.log(f"热键进程已退出，返回码: {code}")
                self.stop_run()

        self.root.after(200, self._drain_queue)

    def on_close(self) -> None:
        self.stop_run()
        if self.session:
            cleanup_session_dir(self.session.paths.session_dir)
            self.session = None
        self.root.destroy()

def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    app = PatrolAssistantApp(root)
    app.log(f"欢迎使用 {APP_NAME} {APP_VERSION}，请先选择Excel文件，然后点击运行。")
    root.mainloop()


if __name__ == "__main__":
    main()
