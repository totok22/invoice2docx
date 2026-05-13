#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Invoice Word Builder - Desktop Application
Cross-platform GUI built with Flet (Material Design 3).
"""

import json
import os
import subprocess
import sys
import threading
from datetime import date
from pathlib import Path
from typing import Any

import flet as ft

from engine import RunConfig, RunResult, STATUS_BLOCKED, STATUS_PASS, fmt_money, run_pipeline

ACCENT = ft.Colors.BLUE_700
ACCENT_LIGHT = ft.Colors.BLUE_50
SUCCESS_COLOR = ft.Colors.GREEN_700
ERROR_COLOR = ft.Colors.RED_700
WARNING_COLOR = ft.Colors.ORANGE_700
SURFACE = ft.Colors.WHITE
APP_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = APP_DIR.parent
APP_VERSION = "V0.0"
APP_BUILD_DATE = "2026-05-13"
APP_AUTHOR = "totok22"


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


def _default_template_schemes() -> list[dict[str, str]]:
    return [
        {
            "name": "默认模板",
            "reimburse_template": str(PROJECT_ROOT / "第三组报账说明.docx"),
            "acceptance_template": str(PROJECT_ROOT / "第三组验收单.docx"),
            "reimburse_name": "第四组报账说明.docx",
            "acceptance_name": "第四组验收单.docx",
        }
    ]


def _default_person_profiles() -> list[dict[str, str]]:
    return [
        {
            "name": "默认资料",
            "season": "",
            "student_id": "",
            "person_name": "",
            "contact": "",
            "bank_name": "",
            "bank_card": "",
        }
    ]


def _default_settings() -> dict[str, Any]:
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
        "template_schemes": _default_template_schemes(),
        "selected_template_scheme": "默认模板",
        "person_profiles": _default_person_profiles(),
        "selected_person_profile": "默认资料",
    }


class InvoiceApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "发票 Word 生成器"
        self.page.window.width = 980
        self.page.window.height = 760
        self.page.window.min_width = 860
        self.page.window.min_height = 640
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.theme = ft.Theme(
            color_scheme_seed=ft.Colors.BLUE,
            font_family="Microsoft YaHei, PingFang SC, sans-serif",
        )
        self.page.padding = 0

        self.result: RunResult | None = None
        self.settings_file = _settings_path()
        self.settings = self._load_settings()
        self.template_schemes = self._normalize_template_schemes(self.settings.get("template_schemes"))
        self.person_profiles = self._normalize_person_profiles(self.settings.get("person_profiles"))
        self.selected_template_scheme = self._pick_existing_name(
            str(self.settings.get("selected_template_scheme", "默认模板")),
            self.template_schemes,
        )
        self.selected_person_profile = self._pick_existing_name(
            str(self.settings.get("selected_person_profile", "默认资料")),
            self.person_profiles,
        )

        self.file_picker = ft.FilePicker()
        self.page.services.append(self.file_picker)
        self.page.update()
        self._build_ui()
        self.page.add(self._layout)

    def _build_ui(self):
        self.template_scheme_dropdown = ft.Dropdown(
            label="模板方案",
            width=220,
            border_radius=8,
            value=self.selected_template_scheme,
            on_select=self._on_template_scheme_change,
        )
        self.profile_dropdown = ft.Dropdown(
            label="资料档案",
            width=220,
            border_radius=8,
            value=self.selected_person_profile,
            on_select=self._on_person_profile_change,
        )
        self.template_hint = ft.Text(size=12, color=ft.Colors.GREY_700)
        self.profile_hint = ft.Text(size=12, color=ft.Colors.GREY_700)

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
            hint_text="模板方案会自动带出，也可手动改",
            value=str(self.settings.get("reimburse_template", "")),
            expand=True,
            read_only=True,
            border_radius=8,
        )
        self.acceptance_tpl_field = ft.TextField(
            label="验收单模板",
            hint_text="模板方案会自动带出，也可手动改",
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
            label="期望购买方抬头",
            value=str(self.settings.get("expected_buyer_name", "")),
            border_radius=8,
            expand=True,
        )
        self.reimburse_name_field = ft.TextField(
            label="报账说明文件名",
            value=str(self.settings.get("reimburse_name", "第四组报账说明.docx")),
            border_radius=8,
            width=220,
        )
        self.acceptance_name_field = ft.TextField(
            label="验收单文件名",
            value=str(self.settings.get("acceptance_name", "第四组验收单.docx")),
            border_radius=8,
            width=220,
        )
        self.ocr_mode_dropdown = ft.Dropdown(
            label="阿里云 OCR",
            value=str(self.settings.get("ocr_mode", "off")),
            width=180,
            border_radius=8,
            options=[
                ft.DropdownOption(key="off", text="关闭"),
                ft.DropdownOption(key="auto", text="自动补救"),
                ft.DropdownOption(key="always", text="强制 OCR"),
            ],
        )
        self.risky_checkbox = ft.Checkbox(
            label="有校验错误也生成 Word",
            value=bool(self.settings.get("allow_risky_generate", False)),
        )
        self.autosave_checkbox = ft.Checkbox(
            label="自动记住本次选择",
            value=bool(self.settings.get("autosave", True)),
        )

        self.progress_bar = ft.ProgressBar(visible=False, color=ACCENT, bgcolor=ACCENT_LIGHT)
        self.progress_text = ft.Text("", size=13, color=ft.Colors.GREY_700)
        self.status_banner = ft.Container(visible=False, padding=12, border_radius=8)
        self.results_column = ft.Column(visible=False, spacing=8)

        self.run_button = ft.Button(
            content=ft.Text("开始生成", color=ft.Colors.WHITE, weight=ft.FontWeight.W_500),
            icon=ft.Icons.PLAY_ARROW_ROUNDED,
            icon_color=ft.Colors.WHITE,
            style=ft.ButtonStyle(bgcolor=ACCENT, padding=ft.Padding(left=32, top=14, right=32, bottom=14)),
            on_click=self._on_run,
        )
        self.save_defaults_button = ft.TextButton(
            content="保存当前设置",
            icon=ft.Icons.SAVE_OUTLINED,
            on_click=self._save_defaults_clicked,
        )
        self.settings_button = ft.IconButton(
            ft.Icons.SETTINGS_OUTLINED,
            tooltip="设置",
            icon_color=ft.Colors.WHITE,
            on_click=self._open_settings,
        )

        self._refresh_template_scheme_dropdown()
        self._refresh_person_profile_dropdown()
        self._apply_template_scheme(self.selected_template_scheme, persist=False, update_page=False)
        self._apply_person_profile(self.selected_person_profile, persist=False, update_page=False)

        file_section = ft.Container(
            content=ft.Column(
                [
                    ft.Text("文件与模板", size=16, weight=ft.FontWeight.W_600, color=ACCENT),
                    ft.Row(
                        [
                            self.template_scheme_dropdown,
                            ft.TextButton(content="管理模板方案", icon=ft.Icons.TUNE, on_click=self._open_settings),
                            ft.Container(content=self.template_hint, expand=True),
                        ],
                        spacing=12,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    ft.Row([self.invoice_dir_field, ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="选择文件夹", on_click=self._pick_invoice_dir)]),
                    ft.Row([self.reimburse_tpl_field, ft.IconButton(ft.Icons.DESCRIPTION, tooltip="选择报账说明模板", on_click=self._pick_reimburse_tpl)]),
                    ft.Row([self.acceptance_tpl_field, ft.IconButton(ft.Icons.DESCRIPTION, tooltip="选择验收单模板", on_click=self._pick_acceptance_tpl)]),
                    ft.Row([self.output_dir_field, ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="选择输出目录", on_click=self._pick_output_dir)]),
                    ft.Row([self.from_xlsx_field, ft.IconButton(ft.Icons.TABLE_CHART, tooltip="选择中间表", on_click=self._pick_from_xlsx)]),
                ],
                spacing=10,
            ),
            padding=20,
            border_radius=12,
            bgcolor=SURFACE,
            shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.with_opacity(0.06, ft.Colors.BLACK)),
        )

        params_section = ft.Container(
            content=ft.Column(
                [
                    ft.Text("参数与档案", size=16, weight=ft.FontWeight.W_600, color=ACCENT),
                    ft.Row(
                        [
                            self.profile_dropdown,
                            ft.TextButton(content="管理资料档案", icon=ft.Icons.BADGE_OUTLINED, on_click=self._open_settings),
                            ft.Container(content=self.profile_hint, expand=True),
                        ],
                        spacing=12,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    ft.Row([self.date_field, self.storage_field, self.reimburse_name_field, self.acceptance_name_field, self.ocr_mode_dropdown], spacing=12, wrap=True),
                    ft.Row([self.buyer_name_field], spacing=12),
                    ft.Row([self.risky_checkbox, self.autosave_checkbox], spacing=16, wrap=True),
                ],
                spacing=10,
            ),
            padding=20,
            border_radius=12,
            bgcolor=SURFACE,
            shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.with_opacity(0.06, ft.Colors.BLACK)),
        )

        action_section = ft.Container(
            content=ft.Column(
                [
                    ft.Row([self.run_button], alignment=ft.MainAxisAlignment.CENTER),
                    self.progress_bar,
                    ft.Row([self.progress_text], alignment=ft.MainAxisAlignment.CENTER),
                    self.status_banner,
                    self.results_column,
                ],
                spacing=12,
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            ),
            padding=20,
            border_radius=12,
            bgcolor=SURFACE,
            shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.with_opacity(0.06, ft.Colors.BLACK)),
        )

        header = ft.Container(
            content=ft.Row(
                [
                    ft.Icon(ft.Icons.RECEIPT_LONG, color=ft.Colors.WHITE, size=28),
                    ft.Text("发票 Word 生成器", size=22, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                    ft.Container(expand=True),
                    ft.Text(f"{APP_AUTHOR}  {APP_VERSION}", color=ft.Colors.WHITE70, size=12),
                    self.settings_button,
                ],
                spacing=12,
            ),
            padding=ft.Padding(left=28, top=18, right=28, bottom=18),
            bgcolor=ACCENT,
        )

        self._layout = ft.Column(
            [
                header,
                ft.Container(
                    content=ft.Column(
                        [
                            self._readiness_panel(),
                            file_section,
                            params_section,
                            ft.Row([self.save_defaults_button], alignment=ft.MainAxisAlignment.END),
                            action_section,
                        ],
                        spacing=16,
                        scroll=ft.ScrollMode.AUTO,
                    ),
                    expand=True,
                    padding=ft.Padding(left=24, top=16, right=24, bottom=16),
                    bgcolor=ft.Colors.GREY_100,
                ),
            ],
            expand=True,
            spacing=0,
        )

    def _normalize_template_schemes(self, raw: Any) -> list[dict[str, str]]:
        schemes: list[dict[str, str]] = []
        for item in raw if isinstance(raw, list) else []:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).strip()
            if not name:
                continue
            schemes.append(
                {
                    "name": name,
                    "reimburse_template": str(item.get("reimburse_template", "")).strip(),
                    "acceptance_template": str(item.get("acceptance_template", "")).strip(),
                    "reimburse_name": str(item.get("reimburse_name", "第四组报账说明.docx")).strip() or "第四组报账说明.docx",
                    "acceptance_name": str(item.get("acceptance_name", "第四组验收单.docx")).strip() or "第四组验收单.docx",
                }
            )
        return schemes or _default_template_schemes()

    def _normalize_person_profiles(self, raw: Any) -> list[dict[str, str]]:
        profiles: list[dict[str, str]] = []
        for item in raw if isinstance(raw, list) else []:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).strip()
            if not name:
                continue
            profiles.append(
                {
                    "name": name,
                    "season": str(item.get("season", "")).strip(),
                    "student_id": str(item.get("student_id", "")).strip(),
                    "person_name": str(item.get("person_name", "")).strip(),
                    "contact": str(item.get("contact", "")).strip(),
                    "bank_name": str(item.get("bank_name", "")).strip(),
                    "bank_card": str(item.get("bank_card", "")).strip(),
                }
            )
        return profiles or _default_person_profiles()

    def _pick_existing_name(self, wanted: str, rows: list[dict[str, str]]) -> str:
        names = [row["name"] for row in rows]
        if wanted in names:
            return wanted
        return names[0]

    def _get_template_scheme(self, name: str) -> dict[str, str] | None:
        for scheme in self.template_schemes:
            if scheme["name"] == name:
                return scheme
        return None

    def _get_person_profile(self, name: str) -> dict[str, str] | None:
        for profile in self.person_profiles:
            if profile["name"] == name:
                return profile
        return None

    def _refresh_template_scheme_dropdown(self):
        self.template_scheme_dropdown.options = [ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.template_schemes]
        self.template_scheme_dropdown.value = self.selected_template_scheme

    def _refresh_person_profile_dropdown(self):
        self.profile_dropdown.options = [ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.person_profiles]
        self.profile_dropdown.value = self.selected_person_profile

    def _refresh_template_hint(self):
        self.template_hint.value = f"当前：{self.selected_template_scheme}"

    def _refresh_profile_hint(self):
        profile = self._get_person_profile(self.selected_person_profile) or {}
        parts = [profile.get("season", ""), profile.get("person_name", ""), profile.get("student_id", "")]
        compact = " / ".join(part for part in parts if part)
        if not compact:
            compact = "当前档案未填写详细信息"
        self.profile_hint.value = compact

    def _apply_template_scheme(self, name: str, persist: bool = True, update_page: bool = True):
        scheme = self._get_template_scheme(name)
        if not scheme:
            return
        self.selected_template_scheme = scheme["name"]
        self.reimburse_tpl_field.value = scheme.get("reimburse_template", "")
        self.acceptance_tpl_field.value = scheme.get("acceptance_template", "")
        self.reimburse_name_field.value = scheme.get("reimburse_name", "第四组报账说明.docx") or "第四组报账说明.docx"
        self.acceptance_name_field.value = scheme.get("acceptance_name", "第四组验收单.docx") or "第四组验收单.docx"
        self.template_scheme_dropdown.value = self.selected_template_scheme
        self._refresh_template_hint()
        if persist:
            self._autosave_settings()
        if update_page:
            self.page.update()

    def _apply_person_profile(self, name: str, persist: bool = True, update_page: bool = True):
        profile = self._get_person_profile(name)
        if not profile:
            return
        self.selected_person_profile = profile["name"]
        self.profile_dropdown.value = self.selected_person_profile
        self._refresh_profile_hint()
        if persist:
            self._autosave_settings()
        if update_page:
            self.page.update()

    def _sync_selected_template_scheme_from_fields(self):
        scheme = self._get_template_scheme(self.selected_template_scheme)
        if not scheme:
            return
        scheme["reimburse_template"] = self.reimburse_tpl_field.value or ""
        scheme["acceptance_template"] = self.acceptance_tpl_field.value or ""
        scheme["reimburse_name"] = self.reimburse_name_field.value or "第四组报账说明.docx"
        scheme["acceptance_name"] = self.acceptance_name_field.value or "第四组验收单.docx"

    def _upsert_template_scheme(self, scheme: dict[str, str]):
        for idx, row in enumerate(self.template_schemes):
            if row["name"] == scheme["name"]:
                self.template_schemes[idx] = scheme
                return
        self.template_schemes.append(scheme)

    def _upsert_person_profile(self, profile: dict[str, str]):
        for idx, row in enumerate(self.person_profiles):
            if row["name"] == profile["name"]:
                self.person_profiles[idx] = profile
                return
        self.person_profiles.append(profile)

    def _load_settings(self) -> dict[str, Any]:
        settings = _default_settings()
        try:
            if self.settings_file.is_file():
                loaded = json.loads(self.settings_file.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    settings.update(loaded)
        except Exception:
            pass
        return settings

    def _current_settings(self) -> dict[str, Any]:
        self._sync_selected_template_scheme_from_fields()
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
            "template_schemes": self.template_schemes,
            "selected_template_scheme": self.selected_template_scheme,
            "person_profiles": self.person_profiles,
            "selected_person_profile": self.selected_person_profile,
        }

    def _save_settings(self, settings: dict[str, Any]):
        self.settings_file.parent.mkdir(parents=True, exist_ok=True)
        self.settings_file.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
        self.settings = settings

    def _autosave_settings(self, force: bool = False):
        if force or self.autosave_checkbox.value:
            self._save_settings(self._current_settings())

    def _apply_defaults(self, settings: dict[str, Any]):
        self.template_schemes = self._normalize_template_schemes(settings.get("template_schemes"))
        self.person_profiles = self._normalize_person_profiles(settings.get("person_profiles"))
        self.selected_template_scheme = self._pick_existing_name(str(settings.get("selected_template_scheme", "")), self.template_schemes)
        self.selected_person_profile = self._pick_existing_name(str(settings.get("selected_person_profile", "")), self.person_profiles)

        self.invoice_dir_field.value = str(settings.get("invoice_dir", ""))
        self.output_dir_field.value = str(settings.get("output_dir", ""))
        self.storage_field.value = str(settings.get("storage_location", "工训楼"))
        self.buyer_name_field.value = str(settings.get("expected_buyer_name", ""))
        self.ocr_mode_dropdown.value = str(settings.get("ocr_mode", "off"))
        self.risky_checkbox.value = bool(settings.get("allow_risky_generate", False))
        self.autosave_checkbox.value = bool(settings.get("autosave", True))

        self._refresh_template_scheme_dropdown()
        self._refresh_person_profile_dropdown()
        self._apply_template_scheme(self.selected_template_scheme, persist=False, update_page=False)
        self._apply_person_profile(self.selected_person_profile, persist=False, update_page=False)

    def _save_defaults_clicked(self, e):
        self._save_settings(self._current_settings())
        self._show_banner(f"已保存：{self.settings_file}", "success")

    def _restore_script_defaults(self, e):
        settings = _default_settings()
        self._apply_defaults(settings)
        self._save_settings(self._current_settings())
        self._show_banner("已恢复默认模板、默认档案和基础参数", "success")
        self.page.pop_dialog()
        self.page.update()

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
            self._sync_selected_template_scheme_from_fields()
            self._autosave_settings()
            self.page.update()

    async def _pick_acceptance_tpl(self, e):
        result = await self.file_picker.pick_files(dialog_title="选择验收单模板", allowed_extensions=["docx"])
        if result:
            self.acceptance_tpl_field.value = result[0].path
            self._sync_selected_template_scheme_from_fields()
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

    def _on_template_scheme_change(self, e):
        if self.template_scheme_dropdown.value:
            self._apply_template_scheme(self.template_scheme_dropdown.value)

    def _on_person_profile_change(self, e):
        if self.profile_dropdown.value:
            self._apply_person_profile(self.profile_dropdown.value)

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

    def _ocr_ready(self) -> bool:
        return bool(os.environ.get("ALIBABA_CLOUD_ACCESS_KEY_ID") and os.environ.get("ALIBABA_CLOUD_ACCESS_KEY_SECRET"))

    def _readiness_panel(self) -> ft.Container:
        return ft.Container(
            content=ft.Row(
                [
                    ft.Icon(ft.Icons.INFO_OUTLINE, color=ACCENT, size=20),
                    ft.Text("先选模板方案和资料档案。复杂发票优先补 XML；校验失败就改中间表再重跑。", size=13, color=ft.Colors.GREY_800, expand=True),
                    ft.TextButton(content="设置", icon=ft.Icons.TUNE, on_click=self._open_settings),
                ],
                spacing=8,
            ),
            padding=ft.Padding(left=14, top=10, right=14, bottom=10),
            border_radius=8,
            bgcolor=ft.Colors.BLUE_50,
        )

    def _open_settings(self, e):
        self.settings_template_select = ft.Dropdown(
            label="模板方案",
            value=self.selected_template_scheme,
            border_radius=8,
            options=[ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.template_schemes],
            on_select=self._settings_template_selected,
        )
        self.settings_template_name = ft.TextField(label="方案名称", value=self.selected_template_scheme, border_radius=8)
        self.settings_template_summary = ft.Text(size=12, color=ft.Colors.GREY_700)

        self.settings_profile_select = ft.Dropdown(
            label="资料档案",
            value=self.selected_person_profile,
            border_radius=8,
            options=[ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.person_profiles],
            on_select=self._settings_profile_selected,
        )
        self.settings_profile_name = ft.TextField(label="档案名称", value=self.selected_person_profile, border_radius=8)
        self.settings_profile_season = ft.TextField(label="赛季", border_radius=8)
        self.settings_profile_student_id = ft.TextField(label="学号", border_radius=8)
        self.settings_profile_person_name = ft.TextField(label="姓名", border_radius=8)
        self.settings_profile_contact = ft.TextField(label="联系方式", border_radius=8)
        self.settings_profile_bank_name = ft.TextField(label="开户行", border_radius=8)
        self.settings_profile_bank_card = ft.TextField(label="卡号", border_radius=8)

        self._load_settings_template_editor(self.selected_template_scheme)
        self._load_settings_profile_editor(self.selected_person_profile)

        ocr_status = "已配置" if self._ocr_ready() else "未配置"
        dialog = ft.AlertDialog(
            modal=False,
            title=ft.Text("设置", weight=ft.FontWeight.W_600),
            content=ft.Container(
                width=760,
                content=ft.Column(
                    [
                        self._settings_block(
                            "模板方案",
                            ft.Column(
                                [
                                    ft.Row([self.settings_template_select, self.settings_template_name], spacing=12),
                                    self.settings_template_summary,
                                    ft.Row(
                                        [
                                            ft.TextButton(content="用当前界面覆盖选中方案", icon=ft.Icons.SAVE_AS_OUTLINED, on_click=self._save_selected_template_scheme),
                                            ft.TextButton(content="按名称另存新方案", icon=ft.Icons.ADD_BOX_OUTLINED, on_click=self._save_new_template_scheme),
                                            ft.TextButton(content="删除选中方案", icon=ft.Icons.DELETE_OUTLINE, on_click=self._delete_template_scheme),
                                        ],
                                        spacing=8,
                                        wrap=True,
                                    ),
                                ],
                                spacing=10,
                            ),
                        ),
                        self._settings_block(
                            "资料档案",
                            ft.Column(
                                [
                                    ft.Row([self.settings_profile_select, self.settings_profile_name], spacing=12),
                                    ft.Row([self.settings_profile_season, self.settings_profile_student_id, self.settings_profile_person_name], spacing=12, wrap=True),
                                    ft.Row([self.settings_profile_contact, self.settings_profile_bank_name, self.settings_profile_bank_card], spacing=12, wrap=True),
                                    ft.Row(
                                        [
                                            ft.TextButton(content="保存当前档案", icon=ft.Icons.SAVE_AS_OUTLINED, on_click=self._save_selected_profile),
                                            ft.TextButton(content="按名称另存新档案", icon=ft.Icons.PERSON_ADD_ALT_1_OUTLINED, on_click=self._save_new_profile),
                                            ft.TextButton(content="删除选中档案", icon=ft.Icons.DELETE_OUTLINE, on_click=self._delete_profile),
                                        ],
                                        spacing=8,
                                        wrap=True,
                                    ),
                                ],
                                spacing=10,
                            ),
                        ),
                        self._settings_block(
                            "运行前提",
                            ft.Column(
                                [
                                    ft.Text(f"OCR：{ocr_status}", size=12, color=ft.Colors.GREY_800),
                                    ft.Text(f"设置文件：{self.settings_file}", size=12, color=ft.Colors.GREY_800, selectable=True),
                                    ft.Text("同一批发票必须是同一购买方。模板不是任意 Word，必须符合本项目表格结构。", size=12, color=ft.Colors.GREY_800),
                                ],
                                spacing=6,
                            ),
                        ),
                    ],
                    spacing=12,
                    scroll=ft.ScrollMode.AUTO,
                ),
            ),
            actions=[
                ft.TextButton(content="恢复默认", icon=ft.Icons.RESTART_ALT, on_click=self._restore_script_defaults),
                ft.TextButton(content="保存全部设置", icon=ft.Icons.SAVE_OUTLINED, on_click=self._save_defaults_from_dialog),
                ft.TextButton(content="关闭", on_click=lambda _: self.page.pop_dialog()),
            ],
        )
        self.page.show_dialog(dialog)

    def _settings_block(self, title: str, content: ft.Control) -> ft.Container:
        return ft.Container(
            content=ft.Column(
                [
                    ft.Text(title, size=14, weight=ft.FontWeight.W_600, color=ACCENT),
                    content,
                ],
                spacing=8,
            ),
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

    def _load_settings_template_editor(self, name: str):
        scheme = self._get_template_scheme(name) or self.template_schemes[0]
        self.settings_template_select.value = scheme["name"]
        self.settings_template_name.value = scheme["name"]
        reimburse = Path(scheme.get("reimburse_template", "")).name or "未设置"
        acceptance = Path(scheme.get("acceptance_template", "")).name or "未设置"
        self.settings_template_summary.value = (
            f"报账说明：{reimburse} | 验收单：{acceptance} | "
            f"输出：{scheme.get('reimburse_name', '')} / {scheme.get('acceptance_name', '')}"
        )

    def _load_settings_profile_editor(self, name: str):
        profile = self._get_person_profile(name) or self.person_profiles[0]
        self.settings_profile_select.value = profile["name"]
        self.settings_profile_name.value = profile["name"]
        self.settings_profile_season.value = profile.get("season", "")
        self.settings_profile_student_id.value = profile.get("student_id", "")
        self.settings_profile_person_name.value = profile.get("person_name", "")
        self.settings_profile_contact.value = profile.get("contact", "")
        self.settings_profile_bank_name.value = profile.get("bank_name", "")
        self.settings_profile_bank_card.value = profile.get("bank_card", "")

    def _settings_template_selected(self, e):
        if self.settings_template_select.value:
            self._load_settings_template_editor(self.settings_template_select.value)
            self.page.update()

    def _settings_profile_selected(self, e):
        if self.settings_profile_select.value:
            self._load_settings_profile_editor(self.settings_profile_select.value)
            self.page.update()

    def _save_selected_template_scheme(self, e):
        name = self.settings_template_select.value or self.settings_template_name.value.strip()
        if not name:
            self._show_banner("模板方案名称不能为空", "error")
            return
        scheme = {
            "name": name,
            "reimburse_template": self.reimburse_tpl_field.value or "",
            "acceptance_template": self.acceptance_tpl_field.value or "",
            "reimburse_name": self.reimburse_name_field.value or "第四组报账说明.docx",
            "acceptance_name": self.acceptance_name_field.value or "第四组验收单.docx",
        }
        self._upsert_template_scheme(scheme)
        self.selected_template_scheme = name
        self._refresh_template_scheme_dropdown()
        self._apply_template_scheme(name, persist=False, update_page=False)
        self.settings_template_select.options = [ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.template_schemes]
        self._load_settings_template_editor(name)
        self._save_settings(self._current_settings())
        self._show_banner(f"已保存模板方案：{name}", "success")
        self.page.update()

    def _save_new_template_scheme(self, e):
        name = self.settings_template_name.value.strip()
        if not name:
            self._show_banner("模板方案名称不能为空", "error")
            return
        scheme = {
            "name": name,
            "reimburse_template": self.reimburse_tpl_field.value or "",
            "acceptance_template": self.acceptance_tpl_field.value or "",
            "reimburse_name": self.reimburse_name_field.value or "第四组报账说明.docx",
            "acceptance_name": self.acceptance_name_field.value or "第四组验收单.docx",
        }
        self._upsert_template_scheme(scheme)
        self.selected_template_scheme = name
        self._refresh_template_scheme_dropdown()
        self._apply_template_scheme(name, persist=False, update_page=False)
        self.settings_template_select.options = [ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.template_schemes]
        self._load_settings_template_editor(name)
        self._save_settings(self._current_settings())
        self._show_banner(f"已新增模板方案：{name}", "success")
        self.page.update()

    def _delete_template_scheme(self, e):
        name = self.settings_template_select.value or ""
        if len(self.template_schemes) <= 1:
            self._show_banner("至少保留一个模板方案", "warning")
            return
        self.template_schemes = [row for row in self.template_schemes if row["name"] != name]
        self.selected_template_scheme = self.template_schemes[0]["name"]
        self._refresh_template_scheme_dropdown()
        self._apply_template_scheme(self.selected_template_scheme, persist=False, update_page=False)
        self.settings_template_select.options = [ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.template_schemes]
        self._load_settings_template_editor(self.selected_template_scheme)
        self._save_settings(self._current_settings())
        self._show_banner(f"已删除模板方案：{name}", "success")
        self.page.update()

    def _profile_from_editor(self, name: str) -> dict[str, str]:
        return {
            "name": name,
            "season": self.settings_profile_season.value.strip(),
            "student_id": self.settings_profile_student_id.value.strip(),
            "person_name": self.settings_profile_person_name.value.strip(),
            "contact": self.settings_profile_contact.value.strip(),
            "bank_name": self.settings_profile_bank_name.value.strip(),
            "bank_card": self.settings_profile_bank_card.value.strip(),
        }

    def _save_selected_profile(self, e):
        name = self.settings_profile_select.value or self.settings_profile_name.value.strip()
        if not name:
            self._show_banner("资料档案名称不能为空", "error")
            return
        profile = self._profile_from_editor(name)
        self._upsert_person_profile(profile)
        self.selected_person_profile = name
        self._refresh_person_profile_dropdown()
        self._apply_person_profile(name, persist=False, update_page=False)
        self.settings_profile_select.options = [ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.person_profiles]
        self._load_settings_profile_editor(name)
        self._save_settings(self._current_settings())
        self._show_banner(f"已保存资料档案：{name}", "success")
        self.page.update()

    def _save_new_profile(self, e):
        name = self.settings_profile_name.value.strip()
        if not name:
            self._show_banner("资料档案名称不能为空", "error")
            return
        profile = self._profile_from_editor(name)
        self._upsert_person_profile(profile)
        self.selected_person_profile = name
        self._refresh_person_profile_dropdown()
        self._apply_person_profile(name, persist=False, update_page=False)
        self.settings_profile_select.options = [ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.person_profiles]
        self._load_settings_profile_editor(name)
        self._save_settings(self._current_settings())
        self._show_banner(f"已新增资料档案：{name}", "success")
        self.page.update()

    def _delete_profile(self, e):
        name = self.settings_profile_select.value or ""
        if len(self.person_profiles) <= 1:
            self._show_banner("至少保留一个资料档案", "warning")
            return
        self.person_profiles = [row for row in self.person_profiles if row["name"] != name]
        self.selected_person_profile = self.person_profiles[0]["name"]
        self._refresh_person_profile_dropdown()
        self._apply_person_profile(self.selected_person_profile, persist=False, update_page=False)
        self.settings_profile_select.options = [ft.DropdownOption(key=row["name"], text=row["name"]) for row in self.person_profiles]
        self._load_settings_profile_editor(self.selected_person_profile)
        self._save_settings(self._current_settings())
        self._show_banner(f"已删除资料档案：{name}", "success")
        self.page.update()

    def _save_defaults_from_dialog(self, e):
        self._save_settings(self._current_settings())
        self.page.pop_dialog()
        self._show_banner(f"已保存：{self.settings_file}", "success")

    def _show_banner(self, message: str, level: str = "info"):
        color_map = {"error": ERROR_COLOR, "success": SUCCESS_COLOR, "warning": WARNING_COLOR}
        bg_map = {"error": ft.Colors.RED_50, "success": ft.Colors.GREEN_50, "warning": ft.Colors.ORANGE_50}
        icon_map = {"error": ft.Icons.ERROR_OUTLINE, "success": ft.Icons.CHECK_CIRCLE_OUTLINE, "warning": ft.Icons.WARNING_AMBER}

        self.status_banner.bgcolor = bg_map.get(level, ft.Colors.BLUE_50)
        self.status_banner.content = ft.Row(
            [
                ft.Icon(icon_map.get(level, ft.Icons.INFO_OUTLINE), color=color_map.get(level, ACCENT), size=20),
                ft.Text(message, color=color_map.get(level, ACCENT), size=14, expand=True),
            ],
            spacing=8,
        )
        self.status_banner.visible = True
        self.page.update()

    def _show_result(self, result: RunResult):
        if result.success:
            self._show_banner(f"生成成功，共 {len(result.invoices)} 张发票，总额 {fmt_money(result.total)} 元", "success")
        else:
            self._show_banner(result.error, "error")

        controls: list[ft.Control] = []
        controls.append(
            ft.Row(
                [
                    self._stat_chip("发票数", str(len(result.invoices)), ft.Icons.RECEIPT),
                    self._stat_chip("物品行", str(len(result.items)), ft.Icons.INVENTORY_2),
                    self._stat_chip("总额", f"¥{fmt_money(result.total)}", ft.Icons.PAYMENTS),
                ],
                spacing=12,
                wrap=True,
            )
        )

        files: list[tuple[str, Path]] = []
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
            controls.append(
                ft.Container(
                    content=ft.Row(
                        [
                            ft.Chip(
                                label=ft.Text(name, size=12),
                                leading=ft.Icon(ft.Icons.INSERT_DRIVE_FILE, size=16),
                                bgcolor=ft.Colors.BLUE_50,
                                on_click=lambda _, p=path: self._open_file(p),
                            )
                            for name, path in files
                        ],
                        spacing=8,
                        wrap=True,
                    ),
                    padding=ft.Padding(left=0, top=8, right=0, bottom=0),
                )
            )

        if result.issues:
            controls.append(ft.Text(f"问题清单（{len(result.issues)} 项）", size=14, weight=ft.FontWeight.W_600, color=WARNING_COLOR))
            issue_rows = []
            for issue in result.issues[:20]:
                level_icon = ft.Icon(ft.Icons.ERROR, color=ERROR_COLOR, size=16) if issue.level == "error" else ft.Icon(ft.Icons.WARNING, color=WARNING_COLOR, size=16)
                issue_rows.append(
                    ft.Container(
                        content=ft.Row(
                            [
                                level_icon,
                                ft.Text(f"[{issue.key}]" if issue.key else "", size=12, width=40),
                                ft.Text(issue.message, size=12, expand=True, overflow=ft.TextOverflow.ELLIPSIS),
                            ],
                            spacing=8,
                        ),
                        padding=ft.Padding(left=8, top=4, right=8, bottom=4),
                        border_radius=4,
                        bgcolor=ft.Colors.RED_50 if issue.level == "error" else ft.Colors.ORANGE_50,
                    )
                )
            controls.append(ft.Column(issue_rows, spacing=4))
            if len(result.issues) > 20:
                controls.append(ft.Text(f"还有 {len(result.issues) - 20} 项，详见中间表", size=12, italic=True, color=ft.Colors.GREY_600))

        if result.invoices:
            controls.append(ft.Text("发票概览", size=14, weight=ft.FontWeight.W_600, color=ACCENT))
            rows = []
            for inv in result.invoices[:30]:
                status_color = SUCCESS_COLOR if inv.status == STATUS_PASS else ERROR_COLOR if inv.status == STATUS_BLOCKED else WARNING_COLOR
                rows.append(
                    ft.DataRow(
                        cells=[
                            ft.DataCell(ft.Text(str(inv.key), size=12)),
                            ft.DataCell(ft.Text(inv.invoice_no[-8:] if len(inv.invoice_no) > 8 else inv.invoice_no, size=12)),
                            ft.DataCell(ft.Text(inv.seller[:12], size=12, overflow=ft.TextOverflow.ELLIPSIS)),
                            ft.DataCell(ft.Text(f"¥{fmt_money(inv.total)}", size=12)),
                            ft.DataCell(
                                ft.Container(
                                    content=ft.Text(inv.status, size=11, color=ft.Colors.WHITE),
                                    bgcolor=status_color,
                                    padding=ft.Padding(left=8, top=2, right=8, bottom=2),
                                    border_radius=10,
                                )
                            ),
                        ]
                    )
                )
            controls.append(
                ft.Container(
                    content=ft.DataTable(
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
                    ),
                    border_radius=8,
                    bgcolor=ft.Colors.WHITE,
                )
            )

        self.results_column.controls = controls
        self.results_column.visible = True
        self.page.update()

    def _stat_chip(self, label: str, value: str, icon: str) -> ft.Container:
        return ft.Container(
            content=ft.Row(
                [
                    ft.Icon(icon, size=18, color=ACCENT),
                    ft.Column(
                        [
                            ft.Text(value, size=16, weight=ft.FontWeight.BOLD),
                            ft.Text(label, size=11, color=ft.Colors.GREY_600),
                        ],
                        spacing=0,
                    ),
                ],
                spacing=8,
            ),
            padding=ft.Padding(left=16, top=10, right=16, bottom=10),
            border_radius=10,
            bgcolor=ACCENT_LIGHT,
        )

    def _open_file(self, path: Path):
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
