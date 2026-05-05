from __future__ import annotations

import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk


def get_app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent


APP_ROOT = get_app_root()
HELPER_SCRIPT = APP_ROOT / "scripts" / "openclaw_helper.ps1"
MODEL_PROVIDER_OPTIONS = {
    "OpenRouter": {"id": "openrouter", "default_model": "openrouter/auto"},
    "OpenAI": {"id": "openai", "default_model": "openai/gpt-5.5"},
    "Anthropic": {"id": "anthropic", "default_model": "anthropic/claude-sonnet-4-5"},
    "MiniMax": {"id": "minimax", "default_model": "minimax/MiniMax-M2.7"},
}
MODEL_PROVIDER_LABELS = list(MODEL_PROVIDER_OPTIONS)


class OpenClawManagerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("OpenClaw 安裝管理工具")
        self.geometry("980x720")
        self.minsize(860, 620)

        self.shell_executable = self._detect_shell_executable()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.task_running = False
        self.gateway_process: subprocess.Popen[str] | None = None
        self.gateway_stop_requested = False

        self.status_text = tk.StringVar(value="準備就緒")
        self.pwsh_var = tk.StringVar(value="未檢查")
        self.node_var = tk.StringVar(value="未檢查")
        self.npm_var = tk.StringVar(value="未檢查")
        self.git_var = tk.StringVar(value="未檢查")
        self.python_var = tk.StringVar(value="未檢查")
        self.opencode_var = tk.StringVar(value="未檢查")
        self.openclaw_var = tk.StringVar(value="未檢查")

        self._build_ui()
        self.after(150, self._drain_log_queue)
        self.refresh_status()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self, padding=16)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        ttk.Label(
            header,
            text="OpenClaw 安裝管理工具",
            font=("Microsoft JhengHei UI", 18, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(header, textvariable=self.status_text).grid(row=0, column=1, sticky="e")

        content = ttk.Frame(self, padding=(16, 0, 16, 12))
        content.grid(row=1, column=0, sticky="nsew")
        content.rowconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)

        env_frame = ttk.LabelFrame(content, text="1. 檢查環境", padding=12)
        env_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 8))
        env_frame.columnconfigure(1, weight=1)

        self._add_status_row(env_frame, 0, "PowerShell 7", self.pwsh_var)
        self._add_status_row(env_frame, 1, "Node.js", self.node_var)
        self._add_status_row(env_frame, 2, "npm", self.npm_var)
        self._add_status_row(env_frame, 3, "Git", self.git_var)
        self._add_status_row(env_frame, 4, "Python", self.python_var)
        self._add_status_row(env_frame, 5, "OpenCode", self.opencode_var)
        self._add_status_row(env_frame, 6, "OpenClaw", self.openclaw_var)

        action_frame = ttk.LabelFrame(content, text="2. 安裝", padding=12)
        action_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        action_frame.columnconfigure(0, weight=1)

        ttk.Button(action_frame, text="重新檢查", command=self.refresh_status).grid(row=0, column=0, sticky="ew")
        ttk.Button(action_frame, text="安裝全部尚未有的環境套件", command=self.install_prerequisites).grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(action_frame, text="安裝 OpenClaw 4.1", command=self.install_openclaw).grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(action_frame, text="安裝最新版 OpenClaw", command=self.install_openclaw_latest).grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(action_frame, text="移除 OpenClaw", command=self.uninstall_openclaw).grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(action_frame, text="設定預設 API Key", command=self.open_api_key_settings).grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(action_frame, text="OpenClaw初始化", command=self.setup_openclaw).grid(row=6, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(action_frame, text="打開交談視窗", command=self.open_dashboard).grid(row=7, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(action_frame, text="備份 .openclaw 資料夾", command=self.archive_openclaw_directory).grid(row=8, column=0, sticky="ew", pady=(10, 0))

        notes = ttk.Label(
            action_frame,
            text=(
                "環境套件安裝會補齊 PowerShell 7、Node.js、npm、Git、Python、OpenCode。\n"
                "OpenCode 透過 npm 全域安裝 opencode-ai。\n"
                "OpenClaw 4.1 會固定安裝 openclaw@2026.4.1。\n"
                "OpenClaw初始化 會執行 openclaw setup。\n"
                "打開交談視窗 會執行 openclaw dashboard。\n"
                "最新版 OpenClaw 與移除 OpenClaw 也會透過 npm 全域執行。\n"
                "備份 .openclaw 資料夾 會把目前的 .openclaw 改名成 .openclaw-yyyymmdd-hhmmss。\n"
                "設定預設 API Key 會寫入 OpenClaw auth store，並切換對應供應商的預設模型。"
            ),
            justify="left",
        )
        notes.grid(row=9, column=0, sticky="w", pady=(12, 0))

        gateway_frame = ttk.LabelFrame(content, text="3. Gateway", padding=12)
        gateway_frame.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(12, 0))
        gateway_frame.columnconfigure(0, weight=1)

        self.start_gateway_button = ttk.Button(gateway_frame, text="啟動 OpenClaw Gateway", command=self.start_gateway)
        self.start_gateway_button.grid(row=0, column=0, sticky="ew")
        self.stop_gateway_button = ttk.Button(gateway_frame, text="停止 OpenClaw Gateway", command=self.stop_gateway)
        self.stop_gateway_button.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        gateway_notes = ttk.Label(
            gateway_frame,
            text=(
                "啟動會執行 openclaw gateway run --force，並將輸出寫入下方日誌。\n"
                "停止只會關閉由這個工具啟動的 gateway 行程。"
            ),
            justify="left",
        )
        gateway_notes.grid(row=2, column=0, sticky="w", pady=(12, 0))

        self._update_gateway_buttons()

        log_frame = ttk.LabelFrame(self, text="執行日誌", padding=16)
        log_frame.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_widget = tk.Text(log_frame, wrap="word", font=("Consolas", 10), state="disabled")
        self.log_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_widget.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_widget.configure(yscrollcommand=scrollbar.set)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _add_status_row(self, parent: ttk.Widget, row: int, label: str, variable: tk.StringVar) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Label(parent, textvariable=variable).grid(row=row, column=1, sticky="w", pady=4)

    def _detect_shell_executable(self) -> str:
        for candidate in ("pwsh", "powershell"):
            resolved = shutil.which(candidate)
            if resolved:
                return resolved
        raise RuntimeError("找不到 pwsh 或 powershell，無法執行安裝腳本。")

    def _drain_log_queue(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_widget.configure(state="normal")
            self.log_widget.insert("end", f"{message}\n")
            self.log_widget.see("end")
            self.log_widget.configure(state="disabled")
        self.after(150, self._drain_log_queue)

    def log(self, message: str) -> None:
        self.log_queue.put(message.rstrip())

    def set_status(self, message: str) -> None:
        self.after(0, lambda: self.status_text.set(message))

    def clear_log(self) -> None:
        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.configure(state="disabled")

    def _run_task(self, label: str, target) -> None:
        if self.task_running:
            messagebox.showinfo("工作進行中", "請先等待目前作業完成。")
            return

        def worker() -> None:
            self.task_running = True
            self.set_status(label)
            try:
                target()
            except Exception as exc:
                error_message = str(exc)
                self.log(f"[錯誤] {error_message}")
                self.after(0, lambda message=error_message: messagebox.showerror("執行失敗", message))
            finally:
                self.task_running = False
                self.set_status("準備就緒")

        threading.Thread(target=worker, daemon=True).start()

    def _powershell_command(self, extra_args: list[str]) -> list[str]:
        args = [self.shell_executable, "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass"]
        return args + extra_args

    def _run_helper_json(self, action: str) -> dict:
        command = self._powershell_command(["-File", str(HELPER_SCRIPT), "-Action", action])
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or f"{action} 執行失敗")
        payload = (result.stdout or "").strip()
        if not payload:
            raise RuntimeError(f"{action} 沒有回傳可解析的輸出。")
        return json.loads(payload)

    def _stream_process(self, command: list[str]) -> int:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        assert process.stdout is not None
        for line in process.stdout:
            self.log(line.rstrip())

        return process.wait()

    def _is_gateway_running(self) -> bool:
        return self.gateway_process is not None and self.gateway_process.poll() is None

    def _update_gateway_buttons(self) -> None:
        start_state = "disabled" if self._is_gateway_running() else "normal"
        stop_state = "normal" if self._is_gateway_running() else "disabled"
        self.start_gateway_button.configure(state=start_state)
        self.stop_gateway_button.configure(state=stop_state)

    def _resolve_openclaw_path(self) -> str:
        status = self._run_helper_json("status")
        tool_info = status.get("tools", {}).get("openclaw") or {}
        openclaw_path = tool_info.get("path")
        if not tool_info.get("installed") or not openclaw_path:
            raise RuntimeError("尚未找到 openclaw。請先安裝 OpenClaw 後再啟動 Gateway。")
        return str(openclaw_path)

    def _build_openclaw_cli_command(self, extra_args: list[str]) -> list[str]:
        openclaw_path = self._resolve_openclaw_path()
        path_suffix = Path(openclaw_path).suffix.lower()

        if path_suffix == ".ps1":
            return self._powershell_command(["-File", openclaw_path, *extra_args])
        if path_suffix in {".cmd", ".bat"}:
            return [os.environ.get("COMSPEC", "cmd.exe"), "/c", openclaw_path, *extra_args]
        return [openclaw_path, *extra_args]

    def _build_openclaw_command(self) -> list[str]:
        return self._build_openclaw_cli_command(["gateway", "run", "--force"])

    def _launch_gateway_process(self) -> subprocess.Popen[str]:
        command = self._build_openclaw_command()
        self.log("[執行] 啟動 OpenClaw Gateway")
        self.log(f"[資訊] 指令: {' '.join(command)}")
        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creationflags,
        )
        self.gateway_process = process
        self.gateway_stop_requested = False
        self.after(0, self._update_gateway_buttons)
        self.log(f"[完成] OpenClaw Gateway 已啟動，PID {process.pid}")
        threading.Thread(target=self._monitor_gateway_process, args=(process,), daemon=True).start()
        return process

    def _run_openclaw_command(
        self,
        args: list[str],
        *,
        input_text: str | None = None,
        capture_output: bool = False,
    ) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            self._build_openclaw_cli_command(args),
            input=input_text,
            capture_output=capture_output,
            text=True,
            check=False,
            encoding="utf-8",
            errors="replace",
        )

    def _run_openclaw_json_command(self, args: list[str]) -> dict:
        result = self._run_openclaw_command(args, capture_output=True)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "OpenClaw 指令執行失敗")
        payload = (result.stdout or "").strip()
        if not payload:
            raise RuntimeError("OpenClaw 沒有回傳可解析的輸出。")
        return json.loads(payload)

    def _get_provider_label_from_default_model(self) -> str:
        try:
            status = self._run_openclaw_json_command(["models", "status", "--json"])
        except Exception:
            return "OpenAI"

        default_model = status.get("resolvedDefault") or status.get("defaultModel") or ""
        provider_id = default_model.split("/", 1)[0] if "/" in default_model else ""
        for label, provider in MODEL_PROVIDER_OPTIONS.items():
            if provider["id"] == provider_id:
                return label
        return "OpenAI"

    def open_api_key_settings(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("設定預設 API Key")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=16)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        provider_var = tk.StringVar(value=self._get_provider_label_from_default_model())
        api_key_var = tk.StringVar()
        model_hint_var = tk.StringVar()

        ttk.Label(frame, text="供應商").grid(row=0, column=0, sticky="w", pady=(0, 8))
        provider_combo = ttk.Combobox(
            frame,
            textvariable=provider_var,
            values=MODEL_PROVIDER_LABELS,
            state="readonly",
        )
        provider_combo.grid(row=0, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="API Key").grid(row=1, column=0, sticky="w", pady=(0, 8))
        api_key_entry = ttk.Entry(frame, textvariable=api_key_var, show="*")
        api_key_entry.grid(row=1, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, textvariable=model_hint_var, justify="left").grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Label(
            frame,
            text="儲存後會更新該供應商的 token，並切換 OpenClaw 預設模型。",
            justify="left",
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(button_frame, text="取消", command=dialog.destroy).grid(row=0, column=0, padx=(0, 8))

        def update_model_hint(*_args: object) -> None:
            provider = MODEL_PROVIDER_OPTIONS[provider_var.get()]
            model_hint_var.set(f"預設模型: {provider['default_model']}")

        def submit() -> None:
            provider_label = provider_var.get()
            api_key = api_key_var.get().strip()
            if provider_label not in MODEL_PROVIDER_OPTIONS:
                messagebox.showerror("設定失敗", "請選擇有效的供應商。", parent=dialog)
                return
            if not api_key:
                messagebox.showerror("設定失敗", "請輸入 API Key。", parent=dialog)
                return
            self.configure_default_api_key(provider_label, api_key, dialog)

        ttk.Button(button_frame, text="儲存", command=submit).grid(row=0, column=1)

        provider_combo.bind("<<ComboboxSelected>>", update_model_hint)
        update_model_hint()
        api_key_entry.focus_set()

    def configure_default_api_key(self, provider_label: str, api_key: str, dialog: tk.Toplevel) -> None:
        provider = MODEL_PROVIDER_OPTIONS[provider_label]
        provider_id = provider["id"]
        default_model = provider["default_model"]

        def work() -> None:
            self.log(f"[執行] 設定 {provider_label} API Key")
            token_result = self._run_openclaw_command(
                ["models", "auth", "paste-token", "--provider", provider_id],
                input_text=f"{api_key}\n",
                capture_output=True,
            )
            if token_result.returncode != 0:
                raise RuntimeError(token_result.stderr.strip() or token_result.stdout.strip() or f"設定 {provider_label} API Key 失敗")

            model_result = self._run_openclaw_command(["models", "set", default_model], capture_output=True)
            if model_result.returncode != 0:
                raise RuntimeError(model_result.stderr.strip() or model_result.stdout.strip() or f"切換 {provider_label} 預設模型失敗")

            self.log(f"[完成] {provider_label} API Key 已更新")
            self.log(f"[完成] 預設模型已切換為 {default_model}")
            self.after(0, dialog.destroy)
            self.after(0, lambda: messagebox.showinfo("設定完成", f"已設定 {provider_label} API Key，並切換預設模型為 {default_model}。"))

        self._run_task(f"正在設定 {provider_label} API Key", work)

    def _monitor_gateway_process(self, process: subprocess.Popen[str]) -> None:
        assert process.stdout is not None
        for line in process.stdout:
            self.log(line.rstrip())

        return_code = process.wait()
        was_user_stop = self.gateway_stop_requested
        if self.gateway_process is process:
            self.gateway_process = None
        self.after(0, self._update_gateway_buttons)

        if was_user_stop:
            self.log("[完成] OpenClaw Gateway 已停止")
        elif return_code == 0:
            self.log("[警告] OpenClaw Gateway 自行結束，正在自動重新啟動")
            try:
                self._launch_gateway_process()
            except Exception as exc:
                self.log(f"[錯誤] OpenClaw Gateway 自動重啟失敗: {exc}")
        else:
            self.log(f"[警告] OpenClaw Gateway 已結束，代碼 {return_code}，正在自動重新啟動")
            try:
                self._launch_gateway_process()
            except Exception as exc:
                self.log(f"[錯誤] OpenClaw Gateway 自動重啟失敗: {exc}")

    def start_gateway(self) -> None:
        if self._is_gateway_running():
            messagebox.showinfo("Gateway 已啟動", "OpenClaw Gateway 已經在執行中。")
            return

        def work() -> None:
            self.set_status("正在啟動 OpenClaw Gateway")
            try:
                self._launch_gateway_process()
            except Exception as exc:
                error_message = str(exc)
                self.log(f"[錯誤] {error_message}")
                self.after(0, lambda message=error_message: messagebox.showerror("啟動失敗", message))
            finally:
                self.set_status("準備就緒")

        threading.Thread(target=work, daemon=True).start()

    def stop_gateway(self) -> None:
        process = self.gateway_process
        if process is None or process.poll() is not None:
            self.gateway_process = None
            self._update_gateway_buttons()
            messagebox.showinfo("Gateway 未執行", "目前沒有由這個工具啟動的 OpenClaw Gateway。")
            return

        def work() -> None:
            self.set_status("正在停止 OpenClaw Gateway")
            self.log("[執行] 停止 OpenClaw Gateway")
            self.gateway_stop_requested = True
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self.log("[資訊] Gateway 未在 10 秒內結束，改用強制停止")
                process.kill()
                process.wait(timeout=5)
            finally:
                self.set_status("準備就緒")

        threading.Thread(target=work, daemon=True).start()

    def refresh_status(self) -> None:
        def work() -> None:
            self.log("[資訊] 重新檢查環境")
            status = self._run_helper_json("status")
            self.after(0, lambda: self._apply_status(status))

        self._run_task("正在重新檢查環境", work)

    def _apply_status(self, status: dict) -> None:
        tools = status.get("tools", {})
        self.pwsh_var.set(self._format_tool_status(tools.get("pwsh")))
        self.node_var.set(self._format_tool_status(tools.get("node")))
        self.npm_var.set(self._format_tool_status(tools.get("npm")))
        self.git_var.set(self._format_tool_status(tools.get("git")))
        self.python_var.set(self._format_tool_status(tools.get("python")))
        self.opencode_var.set(self._format_tool_status(tools.get("opencode")))
        self.openclaw_var.set(self._format_tool_status(tools.get("openclaw")))
        self.log("[資訊] 環境檢查完成")

    def _format_tool_status(self, tool_info: dict | None) -> str:
        if not tool_info:
            return "未檢查"
        if not tool_info.get("installed"):
            return "未安裝"
        version = tool_info.get("version") or "已安裝"
        path = tool_info.get("path")
        if path:
            return f"{version} | {path}"
        return version

    def _get_openclaw_config_dir(self) -> Path:
        return Path.home() / ".openclaw"

    def _build_openclaw_archive_dir(self) -> Path:
        source_dir = self._get_openclaw_config_dir()
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        target_dir = source_dir.parent / f".openclaw-{timestamp}"
        suffix = 1
        while target_dir.exists():
            target_dir = source_dir.parent / f".openclaw-{timestamp}-{suffix}"
            suffix += 1
        return target_dir

    def _run_helper_action_task(
        self,
        *,
        action: str,
        version: str | None = None,
        task_label: str,
        start_log: str,
        success_log: str,
        error_message: str,
    ) -> None:
        def work() -> None:
            self.log(start_log)
            helper_args = ["-File", str(HELPER_SCRIPT), "-Action", action]
            if version is not None:
                helper_args.extend(["-Version", version])
            command = self._powershell_command(helper_args)
            return_code = self._stream_process(command)
            if return_code != 0:
                raise RuntimeError(error_message)
            self.log(success_log)
            status = self._run_helper_json("status")
            self.after(0, lambda: self._apply_status(status))

        self._run_task(task_label, work)

    def _run_openclaw_action_task(
        self,
        *,
        args: list[str],
        task_label: str,
        start_log: str,
        success_log: str,
        error_message: str,
    ) -> None:
        def work() -> None:
            self.log(start_log)
            command = self._build_openclaw_cli_command(args)
            return_code = self._stream_process(command)
            if return_code != 0:
                raise RuntimeError(error_message)
            self.log(success_log)

        self._run_task(task_label, work)

    def install_prerequisites(self) -> None:
        self._run_helper_action_task(
            action="install-prerequisites",
            task_label="正在安裝缺少的環境套件",
            start_log="[執行] 安裝全部尚未有的環境套件",
            success_log="[完成] 環境套件安裝流程完成",
            error_message="安裝環境套件失敗，請查看日誌。",
        )

    def install_openclaw(self) -> None:
        self._run_helper_action_task(
            action="install-openclaw",
            version="2026.4.1",
            task_label="正在安裝 OpenClaw 4.1",
            start_log="[執行] 安裝 OpenClaw 4.1",
            success_log="[完成] OpenClaw 4.1 安裝完成",
            error_message="安裝 OpenClaw 4.1 失敗，請查看日誌。",
        )

    def install_openclaw_latest(self) -> None:
        self._run_helper_action_task(
            action="install-openclaw-latest",
            task_label="正在安裝最新版 OpenClaw",
            start_log="[執行] 安裝最新版 OpenClaw",
            success_log="[完成] 最新版 OpenClaw 安裝完成",
            error_message="安裝最新版 OpenClaw 失敗，請查看日誌。",
        )

    def uninstall_openclaw(self) -> None:
        self._run_helper_action_task(
            action="uninstall-openclaw",
            task_label="正在移除 OpenClaw",
            start_log="[執行] 移除 OpenClaw",
            success_log="[完成] OpenClaw 已移除",
            error_message="移除 OpenClaw 失敗，請查看日誌。",
        )

    def archive_openclaw_directory(self) -> None:
        def work() -> None:
            source_dir = self._get_openclaw_config_dir()
            if not source_dir.exists():
                raise RuntimeError(f"找不到 {source_dir}，無法備份。")
            if not source_dir.is_dir():
                raise RuntimeError(f"{source_dir} 不是資料夾，無法備份。")

            target_dir = self._build_openclaw_archive_dir()
            self.log(f"[執行] 備份 {source_dir.name} 資料夾")
            self.log(f"[資訊] 重新命名為 {target_dir.name}")
            source_dir.rename(target_dir)
            self.log(f"[完成] 已將 {source_dir.name} 改名為 {target_dir.name}")

        self._run_task(
            "正在備份 .openclaw 資料夾",
            work,
        )

    def setup_openclaw(self) -> None:
        self._run_openclaw_action_task(
            args=["setup"],
            task_label="正在初始化 OpenClaw",
            start_log="[執行] OpenClaw初始化",
            success_log="[完成] OpenClaw 初始化完成",
            error_message="OpenClaw 初始化失敗，請查看日誌。",
        )

    def open_dashboard(self) -> None:
        def work() -> None:
            command = self._build_openclaw_cli_command(["dashboard"])
            self.log("[執行] 打開 OpenClaw 交談視窗")
            self.log(f"[資訊] 指令: {' '.join(command)}")
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            subprocess.Popen(
                command,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL,
                text=True,
                creationflags=creationflags,
            )
            self.log("[完成] 已啟動 OpenClaw 交談視窗")

        self._run_task("正在打開 OpenClaw 交談視窗", work)

    def _on_close(self) -> None:
        if self._is_gateway_running():
            self.gateway_stop_requested = True
            assert self.gateway_process is not None
            self.gateway_process.terminate()
        self.destroy()


def _request_elevation() -> None:
    if sys.platform != "win32":
        return
    import ctypes

    shell32 = ctypes.windll.shell32

    try:
        if shell32.IsUserAnAdmin():
            return
    except Exception as exc:
        raise SystemExit(f"無法確認管理者權限: {exc}") from exc

    parameters = None
    if not getattr(sys, "frozen", False):
        script = str(Path(sys.argv[0]).resolve())
        parameters = f'"{script}"'

    result = shell32.ShellExecuteW(None, "runas", sys.executable, parameters, None, 1)
    if result <= 32:
        try:
            ctypes.windll.user32.MessageBoxW(
                None,
                "OpenClaw 安裝管理工具需要管理者權限。請允許 UAC 提示後重新啟動。",
                "需要管理者權限",
                0x10,
            )
        except Exception:
            pass
        raise SystemExit("需要管理者權限才能執行 OpenClaw Manager。")

    sys.exit(0)


def main() -> None:
    if not HELPER_SCRIPT.exists():
        raise SystemExit(f"找不到輔助腳本: {HELPER_SCRIPT}")
    _request_elevation()
    app = OpenClawManagerApp()
    app.mainloop()


if __name__ == "__main__":
    main()