#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Invoice Word Builder - Desktop Application
Cross-platform GUI built with Flet (Material Design 3).
"""

import threading
import json
import os
import sys
from datetime import date
from pathlib import Path

import flet as ft

from engine import RunConfig, RunResult, run_pipeline, fmt_money, STATUS_PASS, STATUS_BLOCKED

ACCENT = ft.Colors.BLUE_700
ACCENT_LIGHT = ft.Colors.BLUE_50
SUCCESS_COLOR = ft.Colors.GREEN_700
ERROR_COLOR = ft.Colors.RED_700
WARNING_COLOR = ft.Colors.ORANGE_700
SURFACE = ft.Colors.WHITE
APP_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = APP_DIR.parent


def _settings_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    elif sys_platform := os.environ.get("XDG_CONFIG_HOME"):
        base = Path(sys_platform)
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path.home() / ".config"
    return base / "InvoiceWordBuilder" / "settings.json"


def _default_settings() -> dict[str, str | bool]:
    return {
        "invoice_dir": str(PROJECT_ROOT / "发票"),
        "reimburse_template": str(PROJECT_ROOT / "第三组报账说明.docx"),
        "acceptance_template": str(PROJECT_ROOT / "第三组验收单.docx"),
        "output_dir": str(APP_DIR / "output"),
        "storage_location": "工训楼",
        "expected_buyer_name": "北京理工大学教育基金会",
        "reimburse_name": "第四组报账说明.docx",
        "acceptance_name": "第四组验收单.docx",
        "ocr_mode": "off",
        "allow_risky_generate": False,
        "autosave": True,
    }


class InvoiceApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "发票 Word 生成器"
        self.page.window.width = 960
        self.page.window.height = 720
        self.page.window.min_width = 800
        self.page.window.min_height = 600
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.theme = ft.Theme(
            color_scheme_seed=ft.Colors.BLUE,
            font_family="Microsoft YaHei, PingFang SC, sans-serif",
        )
        self.page.padding = 0

        self.result: RunResult | None = None
        self.settings_file = _settings_path()
        self.settings = self._load_settings()
        self.file_picker = ft.FilePicker()
        self.page.services.append(self.file_picker)
        self.page.update()
        self._build_ui()
        self.page.add(self._layout)

    def _build_ui(self):
        self.invoice_dir_field = ft.TextField(
            label="发票文件夹",
            hint_text="选择包含 PDF/XML 发票的文件夹",
            value=str(self.settings.get("invoice_dir", "")),
            expand=True,
            read_only=True,
            border_radius=8,
        )
        self.reimburse_tpl_field = ft.TextField(
            label="报账说明模板",
            hint_text="选择 .docx 模板文件",
            value=str(self.settings.get("reimburse_template", "")),
            expand=True,
            read_only=True,
            border_radius=8,
        )
        self.acceptance_tpl_field = ft.TextField(
            label="验收单模板",
            hint_text="选择 .docx 模板文件",
            value=str(self.settings.get("acceptance_template", "")),
            expand=True,
            read_only=True,
            border_radius=8,
        )
        self.output_dir_field = ft.TextField(
            label="输出目录",
            hint_text="选择输出文件夹",
            value=str(self.settings.get("output_dir", "")),
            expand=True,
            read_only=True,
            border_radius=8,
        )
        self.from_xlsx_field = ft.TextField(
            label="从中间表生成（可选）",
            hint_text="选择修正后的 .xlsx 文件",
            expand=True,
            read_only=True,
            border_radius=8,
        )

        self.date_field = ft.TextField(
            label="文档日期",
            value=date.today().strftime("%Y年%m月%d日"),
            border_radius=8,
            width=200,
        )
        self.storage_field = ft.TextField(
            label="存放地点",
            value=str(self.settings.get("storage_location", "工训楼")),
            border_radius=8,
            width=160,
        )
        self.buyer_name_field = ft.TextField(
            label="期望购买方抬头（可选）",
            value=str(self.settings.get("expected_buyer_name", "")),
            border_radius=8,
            expand=True,
        )
        self.reimburse_name_field = ft.TextField(
            label="报账说明文件名",
            value=str(self.settings.get("reimburse_name", "第四组报账说明.docx")),
            border_radius=8,
            width=200,
        )
        self.acceptance_name_field = ft.TextField(
            label="验收单文件名",
            value=str(self.settings.get("acceptance_name", "第四组验收单.docx")),
            border_radius=8,
            width=200,
        )
        self.ocr_mode_dropdown = ft.Dropdown(
            label="阿里云 OCR",
            value=str(self.settings.get("ocr_mode", "off")),
            width=220,
            border_radius=8,
            options=[
                ft.DropdownOption(key="off", text="关闭"),
                ft.DropdownOption(key="auto", text="自动补救"),
                ft.DropdownOption(key="always", text="强制 OCR"),
            ],
        )
        self.risky_checkbox = ft.Checkbox(
            label="即使有校验错误也强制生成 Word",
            value=bool(self.settings.get("allow_risky_generate", False)),
        )
        self.autosave_checkbox = ft.Checkbox(
            label="自动记住本次选择",
            value=bool(self.settings.get("autosave", True)),
        )

        self.progress_bar = ft.ProgressBar(visible=False, color=ACCENT, bgcolor=ACCENT_LIGHT)
        self.progress_text = ft.Text("", size=13, color=ft.Colors.GREY_700)
        self.status_banner = ft.Container(visible=False, padding=12, border_radius=8)

        self.run_button = ft.Button(
            content=ft.Text("开始生成", color=ft.Colors.WHITE, weight=ft.FontWeight.W_500),
            icon=ft.Icons.PLAY_ARROW_ROUNDED,
            icon_color=ft.Colors.WHITE,
            style=ft.ButtonStyle(bgcolor=ACCENT, padding=ft.Padding(left=32, top=14, right=32, bottom=14)),
            on_click=self._on_run,
        )
        self.save_defaults_button = ft.TextButton(
            content="保存为默认",
            icon=ft.Icons.SAVE_OUTLINED,
            on_click=self._save_defaults_clicked,
        )
        self.settings_button = ft.IconButton(
            ft.Icons.SETTINGS_OUTLINED,
            tooltip="设置与说明",
            icon_color=ft.Colors.WHITE,
            on_click=self._open_settings,
        )

        self.results_column = ft.Column(visible=False, spacing=8)

        # Build layout
        file_section = ft.Container(
            content=ft.Column([
                ft.Text("文件设置", size=16, weight=ft.FontWeight.W_600, color=ACCENT),
                ft.Row([self.invoice_dir_field, ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="选择文件夹", on_click=self._pick_invoice_dir)]),
                ft.Row([self.reimburse_tpl_field, ft.IconButton(ft.Icons.DESCRIPTION, tooltip="选择模板", on_click=self._pick_reimburse_tpl)]),
                ft.Row([self.acceptance_tpl_field, ft.IconButton(ft.Icons.DESCRIPTION, tooltip="选择模板", on_click=self._pick_acceptance_tpl)]),
                ft.Row([self.output_dir_field, ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="选择输出目录", on_click=self._pick_output_dir)]),
                ft.Row([self.from_xlsx_field, ft.IconButton(ft.Icons.TABLE_CHART, tooltip="选择中间表", on_click=self._pick_from_xlsx)]),
            ], spacing=10),
            padding=20,
            border_radius=12,
            bgcolor=SURFACE,
            shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.with_opacity(0.06, ft.Colors.BLACK)),
        )

        params_section = ft.Container(
            content=ft.Column([
                ft.Text("参数设置", size=16, weight=ft.FontWeight.W_600, color=ACCENT),
                ft.Row([self.date_field, self.storage_field, self.reimburse_name_field, self.acceptance_name_field, self.ocr_mode_dropdown], spacing=12, wrap=True),
                ft.Row([self.buyer_name_field], spacing=12),
                ft.Row([self.risky_checkbox, self.autosave_checkbox], spacing=16, wrap=True),
            ], spacing=10),
            padding=20,
            border_radius=12,
            bgcolor=SURFACE,
            shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.with_opacity(0.06, ft.Colors.BLACK)),
        )

        action_section = ft.Container(
            content=ft.Column([
                ft.Row([self.run_button], alignment=ft.MainAxisAlignment.CENTER),
                self.progress_bar,
                ft.Row([self.progress_text], alignment=ft.MainAxisAlignment.CENTER),
                self.status_banner,
                self.results_column,
            ], spacing=12, horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
            padding=20,
            border_radius=12,
            bgcolor=SURFACE,
            shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.with_opacity(0.06, ft.Colors.BLACK)),
        )

        header = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.RECEIPT_LONG, color=ft.Colors.WHITE, size=28),
                ft.Text("发票 Word 生成器", size=22, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                ft.Container(expand=True),
                ft.Text("默认设置已自动加载", color=ft.Colors.WHITE70, size=12),
                self.settings_button,
            ], spacing=12),
            padding=ft.Padding(left=28, top=18, right=28, bottom=18),
            bgcolor=ACCENT,
        )

        self._layout = ft.Column([
            header,
            ft.Container(
                content=ft.Column([
                    self._readiness_panel(),
                    file_section,
                    params_section,
                    ft.Row([self.save_defaults_button], alignment=ft.MainAxisAlignment.END),
                    action_section,
                ], spacing=16, scroll=ft.ScrollMode.AUTO),
                expand=True,
                padding=ft.Padding(left=24, top=16, right=24, bottom=16),
                bgcolor=ft.Colors.GREY_100,
            ),
        ], expand=True, spacing=0)

    async def _pick_invoice_dir(self, e):
        result = await self.file_picker.get_directory_path(dialog_title="选择发票文件夹")
        if result:
            self.invoice_dir_field.value = result
            self._autosave_settings()
            self.page.update()

    async def _pick_reimburse_tpl(self, e):
        result = await self.file_picker.pick_files(dialog_title="选择报账说明模板", allowed_extensions=["docx"])
        if result:
            self.reimburse_tpl_field.value = result[0].path
            self._autosave_settings()
            self.page.update()

    async def _pick_acceptance_tpl(self, e):
        result = await self.file_picker.pick_files(dialog_title="选择验收单模板", allowed_extensions=["docx"])
        if result:
            self.acceptance_tpl_field.value = result[0].path
            self._autosave_settings()
            self.page.update()

    async def _pick_output_dir(self, e):
        result = await self.file_picker.get_directory_path(dialog_title="选择输出目录")
        if result:
            self.output_dir_field.value = result
            self._autosave_settings()
            self.page.update()

    async def _pick_from_xlsx(self, e):
        result = await self.file_picker.pick_files(dialog_title="选择中间表 XLSX", allowed_extensions=["xlsx"])
        if result:
            self.from_xlsx_field.value = result[0].path
            self.page.update()

    def _update_progress(self, message: str, value: float):
        self.progress_text.value = message
        self.progress_bar.value = value
        self.page.update()

    def _validate_inputs(self) -> str | None:
        if not self.from_xlsx_field.value:
            if not self.invoice_dir_field.value:
                return "请选择发票文件夹"
            if not Path(self.invoice_dir_field.value).is_dir():
                return f"发票文件夹不存在：{self.invoice_dir_field.value}"
        if not self.reimburse_tpl_field.value:
            return "请选择报账说明模板"
        if not Path(self.reimburse_tpl_field.value).is_file():
            return f"报账说明模板不存在：{self.reimburse_tpl_field.value}"
        if not self.acceptance_tpl_field.value:
            return "请选择验收单模板"
        if not Path(self.acceptance_tpl_field.value).is_file():
            return f"验收单模板不存在：{self.acceptance_tpl_field.value}"
        if not self.output_dir_field.value:
            return "请选择输出目录"
        if self.from_xlsx_field.value and not Path(self.from_xlsx_field.value).is_file():
            return f"中间表不存在：{self.from_xlsx_field.value}"
        if self.ocr_mode_dropdown.value in {"auto", "always"} and not self._ocr_ready():
            return "阿里云 OCR 需要先配置 ALIBABA_CLOUD_ACCESS_KEY_ID 和 ALIBABA_CLOUD_ACCESS_KEY_SECRET"
        return None

    def _on_run(self, e):
        error = self._validate_inputs()
        if error:
            self._show_banner(error, "error")
            return
        self._autosave_settings(force=bool(self.autosave_checkbox.value))

        self.run_button.disabled = True
        self.progress_bar.visible = True
        self.progress_bar.value = 0
        self.progress_text.value = "准备中…"
        self.status_banner.visible = False
        self.results_column.visible = False
        self.page.update()

        config = RunConfig(
            invoice_dir=Path(self.invoice_dir_field.value or "."),
            reimburse_template=Path(self.reimburse_tpl_field.value),
            acceptance_template=Path(self.acceptance_tpl_field.value),
            output_dir=Path(self.output_dir_field.value),
            from_xlsx=Path(self.from_xlsx_field.value) if self.from_xlsx_field.value else None,
            storage_location=self.storage_field.value or "工训楼",
            document_date=self.date_field.value or date.today().strftime("%Y年%m月%d日"),
            expected_buyer_name=self.buyer_name_field.value or "",
            expected_buyer_tax_id="",
            ocr_mode=self.ocr_mode_dropdown.value or "off",
            reimburse_name=self.reimburse_name_field.value or "报账说明.docx",
            acceptance_name=self.acceptance_name_field.value or "验收单.docx",
            allow_risky_generate=self.risky_checkbox.value,
        )

        def worker():
            try:
                result = run_pipeline(config, progress=self._update_progress)
                self.result = result
                self._show_result(result)
            except Exception as exc:
                self._show_banner(f"运行出错：{exc}", "error")
            finally:
                self.run_button.disabled = False
                self.progress_bar.visible = False
                self.page.update()

        threading.Thread(target=worker, daemon=True).start()

    def _load_settings(self) -> dict[str, str | bool]:
        settings = _default_settings()
        try:
            if self.settings_file.is_file():
                loaded = json.loads(self.settings_file.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    settings.update(loaded)
        except Exception:
            pass
        return settings

    def _current_settings(self) -> dict[str, str | bool]:
        return {
            "invoice_dir": self.invoice_dir_field.value or "",
            "reimburse_template": self.reimburse_tpl_field.value or "",
            "acceptance_template": self.acceptance_tpl_field.value or "",
            "output_dir": self.output_dir_field.value or "",
            "storage_location": self.storage_field.value or "工训楼",
            "expected_buyer_name": self.buyer_name_field.value or "",
            "reimburse_name": self.reimburse_name_field.value or "第四组报账说明.docx",
            "acceptance_name": self.acceptance_name_field.value or "第四组验收单.docx",
            "ocr_mode": self.ocr_mode_dropdown.value or "off",
            "allow_risky_generate": bool(self.risky_checkbox.value),
            "autosave": bool(self.autosave_checkbox.value),
        }

    def _save_settings(self, settings: dict[str, str | bool]):
        self.settings_file.parent.mkdir(parents=True, exist_ok=True)
        self.settings_file.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
        self.settings = settings

    def _autosave_settings(self, force: bool = False):
        if force or self.autosave_checkbox.value:
            self._save_settings(self._current_settings())

    def _save_defaults_clicked(self, e):
        self._save_settings(self._current_settings())
        self._show_banner(f"已保存默认设置：{self.settings_file}", "success")

    def _apply_defaults(self, settings: dict[str, str | bool]):
        self.invoice_dir_field.value = str(settings.get("invoice_dir", ""))
        self.reimburse_tpl_field.value = str(settings.get("reimburse_template", ""))
        self.acceptance_tpl_field.value = str(settings.get("acceptance_template", ""))
        self.output_dir_field.value = str(settings.get("output_dir", ""))
        self.storage_field.value = str(settings.get("storage_location", "工训楼"))
        self.buyer_name_field.value = str(settings.get("expected_buyer_name", ""))
        self.reimburse_name_field.value = str(settings.get("reimburse_name", "第四组报账说明.docx"))
        self.acceptance_name_field.value = str(settings.get("acceptance_name", "第四组验收单.docx"))
        self.ocr_mode_dropdown.value = str(settings.get("ocr_mode", "off"))
        self.risky_checkbox.value = bool(settings.get("allow_risky_generate", False))
        self.autosave_checkbox.value = bool(settings.get("autosave", True))

    def _restore_script_defaults(self, e):
        settings = _default_settings()
        self._apply_defaults(settings)
        self._save_settings(settings)
        self._show_banner("已恢复脚本默认路径和参数", "success")
        self.page.pop_dialog()
        self.page.update()

    def _ocr_ready(self) -> bool:
        return bool(os.environ.get("ALIBABA_CLOUD_ACCESS_KEY_ID") and os.environ.get("ALIBABA_CLOUD_ACCESS_KEY_SECRET"))

    def _readiness_panel(self) -> ft.Container:
        return ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.INFO_OUTLINE, color=ACCENT, size=20),
                ft.Text("推荐给复杂发票附上 XML，命名为“序号.xml”。若校验失败，修正中间表的浅黄色列后用“从中间表生成”重跑。", size=13, color=ft.Colors.GREY_800, expand=True),
                ft.TextButton(content="查看前提", icon=ft.Icons.HELP_OUTLINE, on_click=self._open_settings),
            ], spacing=8),
            padding=ft.Padding(left=14, top=10, right=14, bottom=10),
            border_radius=8,
            bgcolor=ft.Colors.BLUE_50,
        )

    def _open_settings(self, e):
        ocr_status = "已检测到密钥" if self._ocr_ready() else "未检测到密钥"
        dialog = ft.AlertDialog(
            modal=False,
            title=ft.Text("设置与前提说明", weight=ft.FontWeight.W_600),
            content=ft.Container(
                width=720,
                content=ft.Column([
                    self._settings_section("默认路径", [
                        f"发票文件夹：{self.invoice_dir_field.value or '未设置'}",
                        f"报账说明模板：{self.reimburse_tpl_field.value or '未设置'}",
                        f"验收单模板：{self.acceptance_tpl_field.value or '未设置'}",
                        f"输出目录：{self.output_dir_field.value or '未设置'}",
                        f"设置文件：{self.settings_file}",
                    ]),
                    self._settings_section("OCR 前提", [
                        f"当前状态：{ocr_status}",
                        "auto：PDF 文本解析失败或金额不闭合时才调用阿里云 OCR。",
                        "always：没有 XML 匹配的 PDF 都走阿里云 OCR。",
                        "使用 OCR 前需要在系统环境变量中配置 ALIBABA_CLOUD_ACCESS_KEY_ID 和 ALIBABA_CLOUD_ACCESS_KEY_SECRET。",
                    ]),
                    self._settings_section("校验规则", [
                        "同一批发票的购买方抬头必须一致。",
                        "每张发票都会校验发票价税合计与物品明细含税金额合计，差 0.01 也会默认停止生成 Word。",
                        "校验失败时仍会导出 invoice_intermediate.xlsx，优先修正浅黄色列，修正后可在“从中间表生成”中重新生成。",
                        "勾选强制生成只适合临时核对版，审计文件仍会保留风险标记。",
                    ]),
                    self._settings_section("推荐做法", [
                        "复杂发票、明细较多的发票，强烈建议同时放入 XML，解析会比 PDF 文本和 OCR 更稳。",
                        "XML 命名用“序号.xml”即可，例如 28.xml；PDF 仍建议保留序号和金额，便于人工核对。",
                        "发票 PDF 建议以“序号+品类+金额+发票.pdf”命名，例如 28+电子元件+390.05+发票.pdf。",
                        "查验单、付款截图可以放在同一个发票文件夹中，不会进入 Word 明细。",
                    ]),
                ], spacing=12, scroll=ft.ScrollMode.AUTO),
            ),
            actions=[
                ft.TextButton(content="恢复脚本默认", icon=ft.Icons.RESTART_ALT, on_click=self._restore_script_defaults),
                ft.TextButton(content="保存当前为默认", icon=ft.Icons.SAVE_OUTLINED, on_click=self._save_defaults_from_dialog),
                ft.TextButton(content="关闭", on_click=lambda _: self.page.pop_dialog()),
            ],
        )
        self.page.show_dialog(dialog)

    def _settings_section(self, title: str, lines: list[str]) -> ft.Container:
        return ft.Container(
            content=ft.Column([
                ft.Text(title, size=14, weight=ft.FontWeight.W_600, color=ACCENT),
                *[ft.Text(line, size=12, color=ft.Colors.GREY_800, selectable=True) for line in lines],
            ], spacing=4),
            padding=12,
            border_radius=8,
            bgcolor=ft.Colors.GREY_50,
            border=ft.Border(
                top=ft.BorderSide(1, ft.Colors.GREY_200),
                right=ft.BorderSide(1, ft.Colors.GREY_200),
                bottom=ft.BorderSide(1, ft.Colors.GREY_200),
                left=ft.BorderSide(1, ft.Colors.GREY_200),
            ),
        )

    def _save_defaults_from_dialog(self, e):
        self._save_settings(self._current_settings())
        self.page.pop_dialog()
        self._show_banner(f"已保存默认设置：{self.settings_file}", "success")

    def _show_banner(self, message: str, level: str = "info"):
        color_map = {"error": ERROR_COLOR, "success": SUCCESS_COLOR, "warning": WARNING_COLOR}
        bg_map = {"error": ft.Colors.RED_50, "success": ft.Colors.GREEN_50, "warning": ft.Colors.ORANGE_50}
        icon_map = {"error": ft.Icons.ERROR_OUTLINE, "success": ft.Icons.CHECK_CIRCLE_OUTLINE, "warning": ft.Icons.WARNING_AMBER}

        self.status_banner.bgcolor = bg_map.get(level, ft.Colors.BLUE_50)
        self.status_banner.content = ft.Row([
            ft.Icon(icon_map.get(level, ft.Icons.INFO_OUTLINE), color=color_map.get(level, ACCENT), size=20),
            ft.Text(message, color=color_map.get(level, ACCENT), size=14, expand=True),
        ], spacing=8)
        self.status_banner.visible = True
        self.page.update()

    def _show_result(self, result: RunResult):
        if result.success:
            self._show_banner(f"生成成功！共 {len(result.invoices)} 张发票，总额 {fmt_money(result.total)} 元", "success")
        else:
            self._show_banner(result.error, "error")

        controls = []

        # Summary stats
        stats_row = ft.Row([
            self._stat_chip("发票数", str(len(result.invoices)), ft.Icons.RECEIPT),
            self._stat_chip("物品行", str(len(result.items)), ft.Icons.INVENTORY_2),
            self._stat_chip("总额", f"¥{fmt_money(result.total)}", ft.Icons.PAYMENTS),
        ], spacing=12, wrap=True)
        controls.append(stats_row)

        # Output files
        files = []
        if result.reimburse_doc:
            files.append(("报账说明", result.reimburse_doc))
        if result.acceptance_doc:
            files.append(("验收单", result.acceptance_doc))
        if result.intermediate_xlsx:
            files.append(("中间表", result.intermediate_xlsx))
            output_dir = result.intermediate_xlsx.parent
            for label, filename in [
                ("发票审计 CSV", "invoice_audit.csv"),
                ("验收明细 CSV", "acceptance_items.csv"),
                ("审计 JSON", "invoice_audit.json"),
            ]:
                path = output_dir / filename
                if path.exists():
                    files.append((label, path))
            files.append(("输出目录", output_dir))

        if files:
            file_chips = ft.Row([
                ft.Chip(
                    label=ft.Text(name, size=12),
                    leading=ft.Icon(ft.Icons.INSERT_DRIVE_FILE, size=16),
                    bgcolor=ft.Colors.BLUE_50,
                    on_click=lambda _, p=path: self._open_file(p),
                )
                for name, path in files
            ], spacing=8, wrap=True)
            controls.append(ft.Container(content=file_chips, padding=ft.Padding(left=0, top=8, right=0, bottom=0)))

        # Issues table
        if result.issues:
            controls.append(ft.Text(f"问题清单（{len(result.issues)} 项）", size=14, weight=ft.FontWeight.W_600, color=WARNING_COLOR))
            issue_rows = []
            for issue in result.issues[:20]:
                level_icon = ft.Icon(ft.Icons.ERROR, color=ERROR_COLOR, size=16) if issue.level == "error" else ft.Icon(ft.Icons.WARNING, color=WARNING_COLOR, size=16)
                issue_rows.append(
                    ft.Container(
                        content=ft.Row([
                            level_icon,
                            ft.Text(f"[{issue.key}]" if issue.key else "", size=12, width=40),
                            ft.Text(issue.message, size=12, expand=True, overflow=ft.TextOverflow.ELLIPSIS),
                        ], spacing=8),
                        padding=ft.Padding(left=8, top=4, right=8, bottom=4),
                        border_radius=4,
                        bgcolor=ft.Colors.RED_50 if issue.level == "error" else ft.Colors.ORANGE_50,
                    )
                )
            controls.append(ft.Column(issue_rows, spacing=4))
            if len(result.issues) > 20:
                controls.append(ft.Text(f"…还有 {len(result.issues) - 20} 项，详见中间表", size=12, italic=True, color=ft.Colors.GREY_600))

        # Invoice summary table
        if result.invoices:
            controls.append(ft.Text("发票概览", size=14, weight=ft.FontWeight.W_600, color=ACCENT))
            rows = []
            for inv in result.invoices[:30]:
                status_color = SUCCESS_COLOR if inv.status == STATUS_PASS else ERROR_COLOR if inv.status == STATUS_BLOCKED else WARNING_COLOR
                rows.append(
                    ft.DataRow(cells=[
                        ft.DataCell(ft.Text(str(inv.key), size=12)),
                        ft.DataCell(ft.Text(inv.invoice_no[-8:] if len(inv.invoice_no) > 8 else inv.invoice_no, size=12)),
                        ft.DataCell(ft.Text(inv.seller[:12], size=12, overflow=ft.TextOverflow.ELLIPSIS)),
                        ft.DataCell(ft.Text(f"¥{fmt_money(inv.total)}", size=12)),
                        ft.DataCell(ft.Container(
                            content=ft.Text(inv.status, size=11, color=ft.Colors.WHITE),
                            bgcolor=status_color,
                            padding=ft.Padding(left=8, top=2, right=8, bottom=2),
                            border_radius=10,
                        )),
                    ])
                )
            table = ft.DataTable(
                columns=[
                    ft.DataColumn(ft.Text("序号", size=12)),
                    ft.DataColumn(ft.Text("发票号", size=12)),
                    ft.DataColumn(ft.Text("销售方", size=12)),
                    ft.DataColumn(ft.Text("总额", size=12)),
                    ft.DataColumn(ft.Text("状态", size=12)),
                ],
                rows=rows,
                border_radius=8,
                heading_row_height=36,
                data_row_min_height=32,
                data_row_max_height=40,
            )
            controls.append(ft.Container(content=table, border_radius=8, bgcolor=ft.Colors.WHITE))

        self.results_column.controls = controls
        self.results_column.visible = True
        self.page.update()

    def _stat_chip(self, label: str, value: str, icon: str) -> ft.Container:
        return ft.Container(
            content=ft.Row([
                ft.Icon(icon, size=18, color=ACCENT),
                ft.Column([
                    ft.Text(value, size=16, weight=ft.FontWeight.BOLD),
                    ft.Text(label, size=11, color=ft.Colors.GREY_600),
                ], spacing=0),
            ], spacing=8),
            padding=ft.Padding(left=16, top=10, right=16, bottom=10),
            border_radius=10,
            bgcolor=ACCENT_LIGHT,
        )

    def _open_file(self, path: Path):
        import subprocess, sys
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        elif sys.platform == "win32":
            subprocess.Popen(["start", "", str(path)], shell=True)
        else:
            subprocess.Popen(["xdg-open", str(path)])


def main(page: ft.Page):
    InvoiceApp(page)


if __name__ == "__main__":
    ft.run(main)
