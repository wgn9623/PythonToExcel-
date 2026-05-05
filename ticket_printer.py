from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
import csv
import hashlib
import json
import logging
import os
from openpyxl import Workbook
from openpyxl.styles import Font
from pathlib import Path
import shutil
import sys
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import uuid

import pythoncom
import win32com.client as win32
import win32timezone


SEQUENCE_SHEET = "顺序号"
HISTORY_SHEET = "打印记录"
RENDER_SHEET = "打印临时"
DEFAULT_TEMPLATE_SHEETS = [
    "模板1_转账支票",
    "模板2_内部划转",
    "模板3_进帐单",
    "模板4_票汇凭证",
    "模板5_结算支票",
    "模板6_现金支票",
]
TEMPLATE_DISPLAY_ALIASES = {
    "模板2_内部帐户资金划转": "模板2_内部划转",
}
TEMPLATE_SHEET_ALIASES = {
    "模板2_内部划转": ["模板2_内部划转", "模板2_内部帐户资金划转"],
    "模板2_内部帐户资金划转": ["模板2_内部划转", "模板2_内部帐户资金划转"],
}
THEME_LABELS = {
    "excel": "经典Excel风",
    "light_blue": "淡蓝风",
    "bank": "紫色风",
    "classic": "传统经典风",
}
PROGRAM_HELP_TEXT = """一、程序功能

1. 录入付款人、收款人、金额、日期、备注等票据信息。
2. 将录入内容加入待打印列表，并支持修改、删除、批量选择。
3. 按所选模板进行打印预览或直接打印。
4. 支持打印记录查询、票据核验、默认模板与默认付款人设置。
5. 支持账户信息维护、导入和模板下载。

二、模板文件修改说明

1. 模板文件允许新增模板工作表。
2. 可以修改模板中的文字内容。
3. 不能增删行。
4. 不能合并或取消合并单元格。
5. 不能破坏现有填充规则和字段落位，否则会影响预览、打印或数据写入。

如需调整模板，请优先复制现有模板后再修改文字内容。"""
THEME_PRESETS = {
    "excel": {
        "root_bg": "#f2f2f2",
        "panel_bg": "#f8f8f8",
        "panel_alt_bg": "#f5f5f5",
        "title_fg": "#303030",
        "muted_fg": "#666666",
        "text_fg": "#3d3d3d",
        "accent_bg": "#7f9db9",
        "accent_active": "#7392af",
        "accent_pressed": "#6785a1",
        "secondary_bg": "#e6e6e6",
        "secondary_active": "#dddddd",
        "secondary_pressed": "#d2d2d2",
        "secondary_fg": "#404040",
        "border": "#b7b7b7",
        "entry_border": "#b7b7b7",
        "entry_bg": "#ffffff",
        "table_header_bg": "#e5e5e5",
        "table_header_fg": "#404040",
        "table_even_bg": "#ffffff",
        "table_odd_bg": "#fafafa",
        "table_selected_bg": "#d9e8fb",
        "status_idle": "#666666",
        "status_processing": "#4d6d8d",
        "status_success": "#4d6b52",
        "status_error": "#9b5b4a",
    },
    "light_blue": {
        "root_bg": "#edf4fb",
        "panel_bg": "#f7fbff",
        "panel_alt_bg": "#f2f8fe",
        "title_fg": "#274766",
        "muted_fg": "#6a84a1",
        "text_fg": "#355c7d",
        "accent_bg": "#6f98c2",
        "accent_active": "#648db5",
        "accent_pressed": "#587fa6",
        "secondary_bg": "#dfeaf6",
        "secondary_active": "#d4e2f0",
        "secondary_pressed": "#c8d9eb",
        "secondary_fg": "#476784",
        "border": "#b8cce0",
        "entry_border": "#c7d9ea",
        "entry_bg": "#fcfeff",
        "table_header_bg": "#dde9f5",
        "table_header_fg": "#476784",
        "table_even_bg": "#fcfeff",
        "table_odd_bg": "#f4f8fc",
        "table_selected_bg": "#e1ebf6",
        "status_idle": "#6a84a1",
        "status_processing": "#6f98c2",
        "status_success": "#567a63",
        "status_error": "#b06a4b",
    },
    "bank": {
        "root_bg": "#f3effb",
        "panel_bg": "#fcfaff",
        "panel_alt_bg": "#f7f2fd",
        "title_fg": "#4b2f63",
        "muted_fg": "#7b6a92",
        "text_fg": "#5c4476",
        "accent_bg": "#8a63b8",
        "accent_active": "#7c58aa",
        "accent_pressed": "#704e9b",
        "secondary_bg": "#ece3f7",
        "secondary_active": "#e2d6f2",
        "secondary_pressed": "#d7c8ee",
        "secondary_fg": "#61497f",
        "border": "#cbbadf",
        "entry_border": "#dacdee",
        "entry_bg": "#fffefe",
        "table_header_bg": "#eadff7",
        "table_header_fg": "#5f467d",
        "table_even_bg": "#fffefe",
        "table_odd_bg": "#faf7fd",
        "table_selected_bg": "#e6daf5",
        "status_idle": "#7b6a92",
        "status_processing": "#7a56ad",
        "status_success": "#5e7a63",
        "status_error": "#a65b7b",
    },
    "classic": {
        "root_bg": "#f3f0e8",
        "panel_bg": "#fbf8f1",
        "panel_alt_bg": "#f7f3ea",
        "title_fg": "#3f3426",
        "muted_fg": "#7a6b59",
        "text_fg": "#4c4031",
        "accent_bg": "#8b6a43",
        "accent_active": "#7c5e3b",
        "accent_pressed": "#6f5333",
        "secondary_bg": "#e7dcc9",
        "secondary_active": "#ddd0bb",
        "secondary_pressed": "#d3c4ac",
        "secondary_fg": "#5e4d38",
        "border": "#b9aa92",
        "entry_border": "#c8bba6",
        "entry_bg": "#fffdf8",
        "table_header_bg": "#e6dccb",
        "table_header_fg": "#5a4a36",
        "table_even_bg": "#fffdf8",
        "table_odd_bg": "#f8f3ea",
        "table_selected_bg": "#eadfcf",
        "status_idle": "#7a6b59",
        "status_processing": "#8b6a43",
        "status_success": "#5f6f52",
        "status_error": "#9a5b45",
    },
}
THREE_COPY_TEMPLATES = {"模板2_内部划转", "模板2_内部帐户资金划转", "模板5_结算支票"}
ROW_STRIDE = {"2copy": 22, "3copy": 33}
COPY_OFFSETS = {"2copy": [0, 11], "3copy": [0, 11, 22]}
DEFAULT_WORKBOOK_NAME = "template_workbook.xls"
DATA_FILE_NAME = "ticket_printer_data.json"
MIN_SERIAL_NUMBER = 1
SERIAL_NUMBER_WIDTH = 5
MACHINE_PREFIX_LENGTH = 4
WATERMARK_TEXT = "内部结算专用"
LARGE_AMOUNT_WARNING = Decimal("50000.00")
REQUIRED_SHEETS = {SEQUENCE_SHEET}


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


LOG_PATH = app_dir() / "ticket_printer.log"


logging.basicConfig(
    filename=str(LOG_PATH),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)


@dataclass
class DataRow:
    payee_name: str
    payee_account: str
    payee_bank: str
    payer_name: str
    payer_account: str
    amount: Decimal
    note: str
    date_text: str
    payer_bank: str
    in_city: str
    in_county: str
    extra_1: str = ""
    extra_2: str = ""
    printed_count: int = 0
    last_serial: str = ""
    last_security_code: str = ""
    last_printed_at: str = ""


class TicketDataStore:
    def __init__(self, data_path: Path) -> None:
        self.data_path = data_path
        self.existed = self.data_path.exists()
        self.data = self._load()

    def _load(self) -> dict:
        if not self.data_path.exists():
            return self._default_data()
        try:
            with self.data_path.open("r", encoding="utf-8") as file:
                data = json.load(file)
        except Exception:
            logging.exception("读取 JSON 数据文件失败，已使用空数据")
            data = {}
        default_data = self._default_data()
        for key, value in default_data.items():
            data.setdefault(key, value)
        data.setdefault("sequence", default_data["sequence"])
        data.setdefault("machines", default_data["machines"])
        data.setdefault("preferences", default_data["preferences"])
        data["sequence"] = {"current": MIN_SERIAL_NUMBER}
        machines = data.get("machines") or {}
        if not isinstance(machines, dict):
            machines = {}
        normalized_machines: dict[str, dict] = {}
        for raw_mac, item in machines.items():
            machine_mac = self._normalize_mac(raw_mac)
            if not machine_mac:
                continue
            info = item if isinstance(item, dict) else {}
            prefix = self._normalize_machine_prefix(info.get("prefix", ""))
            try:
                current = int(info.get("current") or MIN_SERIAL_NUMBER)
            except Exception:
                current = MIN_SERIAL_NUMBER
            normalized_machines[machine_mac] = {
                "prefix": prefix,
                "current": max(current, MIN_SERIAL_NUMBER),
                "created_at": str(info.get("created_at") or ""),
                "last_used_at": str(info.get("last_used_at") or ""),
                "device_name": str(info.get("device_name") or ""),
                "physical_address": self._format_physical_address(machine_mac),
            }
        data["machines"] = normalized_machines
        data["preferences"]["default_template"] = str(data["preferences"].get("default_template") or DEFAULT_TEMPLATE_SHEETS[0])
        data["preferences"]["operator_name"] = str(data["preferences"].get("operator_name") or "")
        default_payer = data["preferences"].get("default_payer") or {}
        if not isinstance(default_payer, dict):
            default_payer = {}
        data["preferences"]["default_payer"] = {
            "name": str(default_payer.get("name") or ""),
            "bank": str(default_payer.get("bank") or ""),
            "account": str(default_payer.get("account") or ""),
        }
        theme_name = str(data["preferences"].get("theme_name") or "light_blue")
        data["preferences"]["theme_name"] = theme_name if theme_name in THEME_PRESETS else "light_blue"
        return data

    @staticmethod
    def _default_data() -> dict:
        return {
            "sequence": {"current": MIN_SERIAL_NUMBER},
            "preferences": {
                "default_template": DEFAULT_TEMPLATE_SHEETS[0],
                "operator_name": "",
                "theme_name": "light_blue",
                "default_payer": {"name": "", "bank": "", "account": ""},
            },
            "payers": [],
            "payees": [],
            "printed_transfers": [],
            "machines": {},
        }

    def save(self) -> None:
        self.data_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = self.data_path.with_suffix(".tmp")
        with temp_path.open("w", encoding="utf-8") as file:
            json.dump(self.data, file, ensure_ascii=False, indent=2)
        temp_path.replace(self.data_path)

    def get_sequence(self) -> tuple[str, int]:
        machine = self.get_current_machine_config()
        if not machine:
            return "", MIN_SERIAL_NUMBER
        current = max(int(machine.get("current") or 0), MIN_SERIAL_NUMBER)
        return str(machine.get("prefix", "")), current

    def update_sequence(self, prefix: str, current: int) -> None:
        machine = self.get_current_machine_config()
        if not machine:
            raise ValueError("当前电脑尚未设置流水号前置符。")
        old_prefix, old_current = self.get_sequence()
        current = int(current)
        if prefix != old_prefix:
            raise ValueError("当前电脑的流水号前置符不匹配。")
        if current < old_current:
            raise ValueError("流水号只能自动递增，不能回退或复用。")
        machine["current"] = current
        machine["last_used_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.save()

    def take_serials(self, count: int) -> list[str]:
        prefix, current = self.get_sequence()
        if not prefix:
            raise ValueError("当前电脑尚未设置流水号前置符。")
        serials = [self.format_full_serial(prefix, current + offset) for offset in range(count)]
        self.update_sequence(prefix, current + count)
        return serials

    def initialize_sequence(self, current: int) -> None:
        machine = self.get_current_machine_config()
        if machine is None:
            return
        machine["current"] = max(int(current), MIN_SERIAL_NUMBER)
        machine["last_used_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.save()

    def get_current_machine_key(self) -> str:
        return self._normalize_mac(self._get_physical_address())

    def get_current_machine_config(self) -> dict | None:
        machine_key = self.get_current_machine_key()
        if not machine_key:
            return None
        machine = self.data.setdefault("machines", {}).get(machine_key)
        return machine if isinstance(machine, dict) else None

    def machine_prefix_exists(self, prefix: str, exclude_machine_key: str | None = None) -> bool:
        normalized_prefix = self._normalize_machine_prefix(prefix)
        if not normalized_prefix:
            return False
        for machine_key, machine in self.data.get("machines", {}).items():
            if exclude_machine_key and machine_key == exclude_machine_key:
                continue
            if self._normalize_machine_prefix(machine.get("prefix", "")) == normalized_prefix:
                return True
        return False

    def get_suggested_machine_prefix(self) -> str:
        machine_key = self.get_current_machine_key()
        if machine_key and len(machine_key) >= MACHINE_PREFIX_LENGTH:
            return machine_key[-MACHINE_PREFIX_LENGTH:]
        return "0001"

    def register_current_machine(self, prefix: str, reset_sequence: bool = True) -> dict:
        machine_key = self.get_current_machine_key()
        if not machine_key:
            raise ValueError("未读取到当前电脑的 MAC 地址，无法生成前置符。")
        normalized_prefix = self._normalize_machine_prefix(prefix)
        if len(normalized_prefix) != MACHINE_PREFIX_LENGTH:
            raise ValueError(f"前置符必须是 {MACHINE_PREFIX_LENGTH} 位字母或数字。")
        if self.machine_prefix_exists(normalized_prefix, exclude_machine_key=machine_key):
            raise ValueError("该前置符已被其他电脑占用，请换一个。")
        now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        current_record = self.data.setdefault("machines", {}).get(machine_key)
        current_value = MIN_SERIAL_NUMBER if reset_sequence or not isinstance(current_record, dict) else max(int(current_record.get("current") or MIN_SERIAL_NUMBER), MIN_SERIAL_NUMBER)
        self.data["machines"][machine_key] = {
            "prefix": normalized_prefix,
            "current": current_value,
            "created_at": str(current_record.get("created_at") if isinstance(current_record, dict) and current_record.get("created_at") else now_text),
            "last_used_at": now_text,
            "device_name": os.environ.get("COMPUTERNAME", "UNKNOWN"),
            "physical_address": self._format_physical_address(machine_key),
        }
        self.save()
        return self.data["machines"][machine_key]

    @staticmethod
    def format_full_serial(prefix: str, current: int) -> str:
        normalized_prefix = TicketDataStore._normalize_machine_prefix(prefix)
        return f"{normalized_prefix}{int(current):0{SERIAL_NUMBER_WIDTH}d}"

    @staticmethod
    def _normalize_machine_prefix(prefix: str) -> str:
        return "".join(char for char in str(prefix).upper().strip() if char.isalnum())[:MACHINE_PREFIX_LENGTH]

    @staticmethod
    def _normalize_mac(mac_text: str) -> str:
        return "".join(char for char in str(mac_text).upper() if char in "0123456789ABCDEF")

    @classmethod
    def _format_physical_address(cls, mac_text: str) -> str:
        clean = cls._normalize_mac(mac_text)
        if len(clean) != 12:
            return str(mac_text).upper()
        return ":".join(clean[index:index + 2] for index in range(0, len(clean), 2))

    @staticmethod
    def _get_physical_address() -> str:
        mac_int = uuid.getnode()
        return ":".join(f"{(mac_int >> shift) & 0xFF:02X}" for shift in range(40, -1, -8))

    def upsert_account(self, account_type: str, name: str, bank: str, account: str) -> None:
        name = name.strip()
        bank = bank.strip()
        account = account.strip()
        if not name or not bank or not account:
            raise ValueError("名称、开户行、账号都不能为空。")
        table_name = self._account_table_name(account_type)
        table = self.data[table_name]
        record = {
            "name": name,
            "bank": bank,
            "account": account,
            "is_favorite": False,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        for index, item in enumerate(table):
            if str(item.get("account", "")).strip() == account:
                table[index] = {**record, **item, "name": name, "bank": bank, "account": account, "updated_at": record["updated_at"]}
                self.save()
                return
        table.append(record)
        self.save()

    def update_account(
        self,
        account_type: str,
        original_account: str,
        name: str,
        bank: str,
        account: str,
    ) -> None:
        name = name.strip()
        bank = bank.strip()
        account = account.strip()
        original_account = original_account.strip()
        if not name or not bank or not account:
            raise ValueError("名称、开户行、账号都不能为空。")
        table_name = self._account_table_name(account_type)
        table = self.data[table_name]
        target_index = None
        for index, item in enumerate(table):
            if str(item.get("account", "")).strip() == original_account:
                target_index = index
                break
        if target_index is None:
            raise ValueError("未找到要修改的账户。")

        for index, item in enumerate(table):
            if index != target_index and str(item.get("account", "")).strip() == account:
                raise ValueError("新账号已存在，不能与其他账户重复。")

        table[target_index] = {
            **table[target_index],
            "name": name,
            "bank": bank,
            "account": account,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self.save()

    def find_accounts(self, account_type: str, keyword: str) -> list[dict]:
        keyword = keyword.strip().lower()
        table = self.data[self._account_table_name(account_type)]
        if keyword:
            records = [
                item
                for item in table
                if keyword in str(item.get("name", "")).lower()
                or keyword in str(item.get("bank", "")).lower()
                or keyword in str(item.get("account", "")).lower()
            ]
        else:
            records = list(table)
        if account_type == "payer":
            records.sort(
                key=lambda item: (
                    0 if bool(item.get("is_favorite")) else 1,
                    str(item.get("name", "")),
                    str(item.get("account", "")),
                )
            )
        return records

    def toggle_account_favorite(self, account_type: str, account: str, is_favorite: bool) -> None:
        table_name = self._account_table_name(account_type)
        table = self.data[table_name]
        for index, item in enumerate(table):
            if str(item.get("account", "")).strip() == account.strip():
                table[index] = {**item, "is_favorite": bool(is_favorite)}
                self.save()
                return
        raise ValueError("未找到要设置常用的账户。")

    def append_print_record(self, record: dict) -> None:
        self.data["printed_transfers"].append(record)
        self.save()

    def void_print_record(self, record: dict, reason: str) -> dict:
        target = self._find_print_record(record)
        if target is None:
            raise ValueError("未找到要作废的打印记录。")
        if self.is_print_record_voided(target):
            raise ValueError("该记录已经作废，不能重复作废。")
        target.update(
            {
                "record_status": "voided",
                "voided_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "void_reason": reason.strip() or "未填写",
                "void_operator": self.get_operator_name(),
            }
        )
        self.save()
        return target

    def _find_print_record(self, record: dict) -> dict | None:
        for item in self.data["printed_transfers"]:
            if item is record:
                return item
            if (
                str(item.get("operation_time", "")) == str(record.get("operation_time", ""))
                and str(item.get("serial_number", "")) == str(record.get("serial_number", ""))
                and str(item.get("security_code", "")) == str(record.get("security_code", ""))
            ):
                return item
        return None

    @staticmethod
    def is_print_record_voided(record: dict) -> bool:
        return str(record.get("record_status", "")).strip().lower() == "voided"

    def get_default_template(self) -> str:
        return str(self.data.get("preferences", {}).get("default_template") or DEFAULT_TEMPLATE_SHEETS[0])

    def set_default_template(self, template_name: str) -> None:
        self.data.setdefault("preferences", {})
        self.data["preferences"]["default_template"] = template_name
        self.save()

    def get_operator_name(self) -> str:
        return str(self.data.get("preferences", {}).get("operator_name") or "")

    def set_operator_name(self, operator_name: str) -> None:
        self.data.setdefault("preferences", {})
        self.data["preferences"]["operator_name"] = operator_name.strip()
        self.save()

    def get_theme_name(self) -> str:
        theme_name = str(self.data.get("preferences", {}).get("theme_name") or "light_blue")
        return theme_name if theme_name in THEME_PRESETS else "light_blue"

    def set_theme_name(self, theme_name: str) -> None:
        if theme_name not in THEME_PRESETS:
            raise ValueError("未知主题。")
        self.data.setdefault("preferences", {})
        self.data["preferences"]["theme_name"] = theme_name
        self.save()

    def get_default_payer(self) -> dict[str, str]:
        payer = self.data.get("preferences", {}).get("default_payer") or {}
        return {
            "name": str(payer.get("name") or ""),
            "bank": str(payer.get("bank") or ""),
            "account": str(payer.get("account") or ""),
        }

    def set_default_payer(self, name: str, bank: str, account: str) -> None:
        self.data.setdefault("preferences", {})
        self.data["preferences"]["default_payer"] = {
            "name": name.strip(),
            "bank": bank.strip(),
            "account": account.strip(),
        }
        self.save()

    def clear_default_payer(self) -> None:
        self.set_default_payer("", "", "")

    def find_print_records(self, keyword: str) -> list[dict]:
        keyword = keyword.strip().lower()
        records = list(self.data["printed_transfers"])
        if not keyword:
            return records
        filtered: list[dict] = []
        for item in records:
            values = [
                item.get("operation_time", ""),
                item.get("action_type", ""),
                item.get("template", ""),
                item.get("operator_name", ""),
                item.get("serial_number", ""),
                item.get("full_serial", ""),
                item.get("security_code", ""),
                item.get("payee_name", ""),
                item.get("payee_account", ""),
                item.get("payer_name", ""),
                item.get("payer_account", ""),
                item.get("record_status", ""),
                item.get("voided_at", ""),
                item.get("void_reason", ""),
                item.get("void_operator", ""),
            ]
            if any(keyword in str(value).lower() for value in values):
                filtered.append(item)
        return filtered

    def verify_print_record(self, keyword: str) -> dict | None:
        keyword = keyword.strip().lower()
        if not keyword:
            return None
        records = self.find_print_records(keyword)
        exact_matches = [
            item
            for item in records
            if keyword in {
                str(item.get("serial_number", "")).strip().lower(),
                str(item.get("full_serial", "")).strip().lower(),
                str(item.get("security_code", "")).strip().lower(),
            }
        ]
        target = exact_matches[-1] if exact_matches else (records[-1] if records else None)
        return target

    @staticmethod
    def _account_table_name(account_type: str) -> str:
        if account_type == "payer":
            return "payers"
        if account_type == "payee":
            return "payees"
        raise ValueError(f"未知账户类型：{account_type}")


class WorkbookSession:
    def __init__(self, workbook_path: Path, use_temp: bool = True) -> None:
        self.original_path = workbook_path.resolve()
        self.use_temp = use_temp
        self.closed = False
        if self.use_temp:
            self.temp_dir = Path(tempfile.mkdtemp(prefix="ticket_printer_"))
            self.temp_path = self.temp_dir / "workbook.xls"
            shutil.copy2(self.original_path, self.temp_path)
            open_path = self.temp_path
        else:
            self.temp_dir = None
            self.temp_path = self.original_path
            open_path = self.original_path

        pythoncom.CoInitialize()
        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.EnableEvents = False
        try:
            self.excel.AutomationSecurity = 3
        except Exception:
            pass

        self.workbook = self.excel.Workbooks.Open(
            str(open_path),
            UpdateLinks=0,
            IgnoreReadOnlyRecommended=True,
        )

    def close(self, save: bool = True) -> None:
        if self.closed:
            return
        self.closed = True
        try:
            if save:
                self.workbook.Save()
                self.workbook.Close(True)
                if self.use_temp:
                    shutil.copy2(self.temp_path, self.original_path)
            else:
                self.workbook.Close(False)
        finally:
            self.excel.Quit()
            pythoncom.CoUninitialize()
            if self.temp_dir is not None:
                shutil.rmtree(self.temp_dir, ignore_errors=True)

    def sheet(self, name: str):
        return self.workbook.Worksheets(name)


class TicketPrinterService:
    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path
        self.store = TicketDataStore(app_dir() / DATA_FILE_NAME)
        if not self.store.existed:
            self._migrate_sequence_from_workbook()

    def validate_workbook(self) -> None:
        session = WorkbookSession(self.workbook_path)
        save_changes = True
        try:
            self._validate_workbook_in_session(session)
            _prefix, current = self.get_sequence_state()
            if current < MIN_SERIAL_NUMBER:
                raise ValueError(f"流水号不能小于 {MIN_SERIAL_NUMBER}。")
        finally:
            session.close(save=False)

    def list_template_sheets(self) -> list[str]:
        session = WorkbookSession(self.workbook_path)
        try:
            sheet_names = [str(sheet.Name) for sheet in session.workbook.Worksheets]
            templates = [name for name in sheet_names if name.startswith("模板")]
            if not templates:
                return list(DEFAULT_TEMPLATE_SHEETS)
            def sort_key(name: str):
                prefix = name.split("_", 1)[0]
                try:
                    order = int(prefix.replace("模板", ""))
                except Exception:
                    order = 999
                return (order, name)
            return sorted(templates, key=sort_key)
        finally:
            session.close(save=False)

    def get_sequence_state(self) -> tuple[str, int]:
        return self.store.get_sequence()

    def format_serial(self, prefix: str, current: int) -> str:
        return self.store.format_full_serial(prefix, current)

    def get_machine_prefix(self) -> str:
        prefix, _current = self.get_sequence_state()
        return prefix

    def get_machine_config(self) -> dict | None:
        return self.store.get_current_machine_config()

    def get_suggested_machine_prefix(self) -> str:
        return self.store.get_suggested_machine_prefix()

    def register_current_machine(self, prefix: str, reset_sequence: bool = True) -> dict:
        return self.store.register_current_machine(prefix, reset_sequence=reset_sequence)

    def save_account(self, account_type: str, name: str, bank: str, account: str) -> None:
        self.store.upsert_account(account_type, name, bank, account)

    def update_account(self, account_type: str, original_account: str, name: str, bank: str, account: str) -> None:
        self.store.update_account(account_type, original_account, name, bank, account)

    def find_accounts(self, account_type: str, keyword: str) -> list[dict]:
        return self.store.find_accounts(account_type, keyword)

    def toggle_account_favorite(self, account_type: str, account: str, is_favorite: bool) -> None:
        self.store.toggle_account_favorite(account_type, account, is_favorite)

    def find_print_records(self, keyword: str) -> list[dict]:
        return self.store.find_print_records(keyword)

    def void_print_record(self, record: dict, reason: str) -> dict:
        return self.store.void_print_record(record, reason)

    def verify_print_record(self, keyword: str) -> dict | None:
        return self.store.verify_print_record(keyword)

    def get_all_print_records(self) -> list[dict]:
        return self.store.find_print_records("")

    def get_default_template(self) -> str:
        return self.store.get_default_template()

    def set_default_template(self, template_name: str) -> None:
        self.store.set_default_template(template_name)

    def get_operator_name(self) -> str:
        return self.store.get_operator_name()

    def set_operator_name(self, operator_name: str) -> None:
        self.store.set_operator_name(operator_name)

    def get_theme_name(self) -> str:
        return self.store.get_theme_name()

    def set_theme_name(self, theme_name: str) -> None:
        self.store.set_theme_name(theme_name)

    def get_default_payer(self) -> dict[str, str]:
        return self.store.get_default_payer()

    def set_default_payer(self, name: str, bank: str, account: str) -> None:
        self.store.set_default_payer(name, bank, account)

    def clear_default_payer(self) -> None:
        self.store.clear_default_payer()

    def import_accounts(self, account_type: str, records: list[dict[str, str]]) -> int:
        imported = 0
        for item in records:
            self.save_account(
                account_type,
                str(item.get("name", "")),
                str(item.get("bank", "")),
                str(item.get("account", "")),
            )
            imported += 1
        return imported

    def delete_account(self, account_type: str, account: str) -> None:
        table_name = self.store._account_table_name(account_type)
        table = self.store.data[table_name]
        filtered = [item for item in table if str(item.get("account", "")).strip() != account.strip()]
        if len(filtered) == len(table):
            raise ValueError("未找到要删除的账户。")
        self.store.data[table_name] = filtered
        self.store.save()

    def _migrate_sequence_from_workbook(self) -> None:
        session = WorkbookSession(self.workbook_path)
        try:
            _prefix, current = self._read_sequence_in_session(session)
            machine = self.store.get_current_machine_config()
            if machine is not None:
                machine["current"] = max(int(current), MIN_SERIAL_NUMBER)
                self.store.save()
        except Exception:
            logging.exception("从 Excel 迁移流水号失败，已使用 JSON 默认流水号")
        finally:
            session.close(save=False)

    def save_single_row(self, row: DataRow) -> None:
        raise RuntimeError("当前版本不再使用单独的数据表保存，请直接使用预览或打印流程。")

    def build_print_sheet(self, template_name: str, backup: bool = False) -> int:
        raise RuntimeError("当前版本不再通过独立打印页批量构建，请直接使用选中记录预览或打印。")
        session = WorkbookSession(self.workbook_path)
        try:
            self._validate_template_name(template_name)
            self._validate_workbook_in_session(session)
            self._clear_print_sheet(session)
            self._copy_template(session, template_name)
            data_rows = self._read_data_rows(session)
            if not data_rows:
                raise ValueError("数据表中没有可打印的数据。")

            copy_mode = "3copy" if template_name in THREE_COPY_TEMPLATES else "2copy"
            self._fill_print_sheet(session, data_rows, copy_mode)
            self._add_sequence_numbers(session, len(data_rows), copy_mode)
            if backup:
                self._backup_data(session, data_rows, copy_mode)
            return len(data_rows)
        finally:
            session.close(save=True)

    def preview_workbook(self) -> None:
        session = WorkbookSession(self.workbook_path)
        try:
            session.excel.Visible = True
            session.sheet(RENDER_SHEET).PrintPreview()
        finally:
            session.close(save=False)

    def print_workbook(self) -> None:
        session = WorkbookSession(self.workbook_path)
        try:
            session.sheet(RENDER_SHEET).PrintOut()
        finally:
            session.close(save=False)

    def preview_single_row(self, row: DataRow, template_name: str) -> bool:
        return self.preview_rows([row], template_name)

    def print_single_row(self, row: DataRow, template_name: str) -> bool:
        return self.print_rows([row], template_name)

    def preview_rows(self, rows: list[DataRow], template_name: str) -> bool:
        return self._render_rows(rows, template_name, do_preview=True, do_print=False)

    def print_rows(self, rows: list[DataRow], template_name: str) -> bool:
        return bool(self._print_rows_individually(rows, template_name))

    def print_rows_with_plan(
        self,
        rows: list[DataRow],
        template_name: str,
        full_serials: list[str] | None = None,
        action_types: list[str] | None = None,
    ) -> list[dict] | None:
        return self._print_rows_individually(rows, template_name, full_serials=full_serials, action_types=action_types)

    def _read_sequence_in_session(self, session: WorkbookSession) -> tuple[str, int]:
        sheet = session.sheet(SEQUENCE_SHEET)
        prefix = str(sheet.Cells(2, "A").Value or "").strip()
        raw_current = sheet.Cells(2, "B").Value
        if raw_current in (None, ""):
            return prefix, 0

        try:
            current_decimal = Decimal(str(raw_current)).quantize(Decimal("1"))
        except Exception as exc:
            raise ValueError(f"顺序号表 B2 必须是数字，当前值是：{raw_current}") from exc

        return prefix, int(current_decimal)

    def _render_rows(self, rows: list[DataRow], template_name: str, do_preview: bool, do_print: bool) -> bool:
        if not rows:
            raise ValueError("请先选择至少一条数据。")
        for row in rows:
            self._validate_input_row(row)
        if do_preview and len(rows) > 1:
            return self._preview_rows_individually(rows, template_name)
        if do_print:
            return self._print_rows_individually(rows, template_name)

        session = WorkbookSession(self.workbook_path)
        save_changes = True
        try:
            self._validate_template_name(template_name)
            self._validate_workbook_in_session(session)
            render_sheet = self._prepare_render_sheet(session, template_name)
            copy_mode = "3copy" if template_name in THREE_COPY_TEMPLATES else "2copy"
            self._fill_render_sheet(render_sheet, rows, copy_mode)
            full_serials = self._add_sequence_numbers(render_sheet, session, len(rows), copy_mode)
            security_items = self._add_security_codes(render_sheet, rows, full_serials, copy_mode)
            if full_serials:
                action_type = "预览" if do_preview and not do_print else "打印"
                for row, full_serial, (security_code, printed_at) in zip(rows, full_serials, security_items):
                    try:
                        self._append_history(session, row, template_name, full_serial, action_type, security_code, printed_at)
                    except Exception:
                        logging.exception("写入打印记录失败，不阻断预览或打印")
            if do_preview:
                session.workbook.Save()
                session.excel.Visible = True
                render_sheet.Activate()
                render_sheet.PrintPreview()
            if do_print:
                session.excel.Visible = True
                render_sheet.Activate()
                if not session.excel.Dialogs(8).Show():
                    save_changes = False
                    return False
                session.workbook.Save()
            return True
        finally:
            session.close(save=save_changes)

    def _render_single_row(self, row: DataRow, template_name: str, do_preview: bool, do_print: bool) -> bool:
        return self._render_rows([row], template_name, do_preview=do_preview, do_print=do_print)

    def _preview_rows_individually(self, rows: list[DataRow], template_name: str) -> bool:
        session = WorkbookSession(self.workbook_path)
        save_changes = False
        try:
            self._validate_template_name(template_name)
            self._validate_workbook_in_session(session)
            session.excel.Visible = True

            preview_sheet_names: list[str] = []
            copy_mode = "3copy" if template_name in THREE_COPY_TEMPLATES else "2copy"
            for index, row in enumerate(rows, start=1):
                sheet_name = RENDER_SHEET if index == 1 else f"{RENDER_SHEET}_{index}"
                render_sheet = self._prepare_render_sheet(session, template_name, sheet_name=sheet_name)
                self._fill_render_sheet(render_sheet, [row], copy_mode)
                full_serial = self._add_sequence_numbers(render_sheet, session, 1, copy_mode)[0]
                security_code, printed_at = self._add_security_codes(render_sheet, [row], [full_serial], copy_mode)[0]
                try:
                    self._append_history(session, row, template_name, full_serial, "预览", security_code, printed_at)
                except Exception:
                    logging.exception("写入打印记录失败，不阻断预览")
                preview_sheet_names.append(render_sheet.Name)

            session.workbook.Save()
            session.excel.Worksheets(preview_sheet_names).Select()
            session.excel.ActiveWindow.SelectedSheets.PrintPreview()
            return True
        finally:
            session.close(save=save_changes)

    def _print_rows_individually(
        self,
        rows: list[DataRow],
        template_name: str,
        full_serials: list[str] | None = None,
        action_types: list[str] | None = None,
    ) -> list[dict] | None:
        session = WorkbookSession(self.workbook_path)
        save_changes = True
        results: list[dict] = []
        try:
            self._validate_template_name(template_name)
            self._validate_workbook_in_session(session)
            copy_mode = "3copy" if template_name in THREE_COPY_TEMPLATES else "2copy"
            session.excel.Visible = True

            for index, row in enumerate(rows):
                render_sheet = self._prepare_render_sheet(session, template_name)
                self._fill_render_sheet(render_sheet, [row], copy_mode)
                render_sheet.Activate()
                if index == 0 and not session.excel.Dialogs(9).Show():
                    save_changes = False
                    return None

                if full_serials and index < len(full_serials):
                    full_serial = full_serials[index]
                    self._write_serials_to_sheet(render_sheet, [full_serial], copy_mode)
                else:
                    full_serial = self._add_sequence_numbers(render_sheet, session, 1, copy_mode)[0]
                security_code, printed_at = self._add_security_codes(render_sheet, [row], [full_serial], copy_mode)[0]
                action_type = action_types[index] if action_types and index < len(action_types) else "打印"
                try:
                    self._append_history(session, row, template_name, full_serial, action_type, security_code, printed_at)
                except Exception:
                    logging.exception("写入打印记录失败，不阻断打印")
                session.workbook.Save()
                render_sheet.PrintOut()
                results.append(
                    {
                        "full_serial": full_serial,
                        "security_code": security_code,
                        "printed_at": printed_at,
                        "action_type": action_type,
                    }
                )
            return results
        finally:
            session.close(save=save_changes)

    def _validate_template_name(self, template_name: str) -> None:
        if not str(template_name).strip():
            raise ValueError(f"不支持的模板：{template_name}")

    @staticmethod
    def _resolve_template_sheet_name(sheet_names: set[str], template_name: str) -> str:
        candidates = TEMPLATE_SHEET_ALIASES.get(template_name, [template_name])
        for candidate in candidates:
            if candidate in sheet_names:
                return candidate
        raise ValueError(f"Excel 缺少工作表：{template_name}")

    def _validate_workbook_in_session(self, session: WorkbookSession) -> None:
        sheet_names = {sheet.Name for sheet in session.workbook.Worksheets}
        missing_sheets: list[str] = []
        for required_sheet in REQUIRED_SHEETS:
            if required_sheet not in sheet_names:
                missing_sheets.append(required_sheet)
        template_sheets = [name for name in sheet_names if str(name).startswith("模板")]
        if not template_sheets:
            missing_sheets.append("至少一张模板工作表")
        if missing_sheets:
            raise ValueError(f"Excel 缺少工作表：{', '.join(missing_sheets)}")

    def _prepare_render_sheet(self, session: WorkbookSession, template_name: str, sheet_name: str = RENDER_SHEET):
        if sheet_name == RENDER_SHEET:
            try:
                session.sheet(RENDER_SHEET).Delete()
            except Exception:
                pass
        else:
            try:
                session.sheet(sheet_name).Delete()
            except Exception:
                pass
        actual_template_name = self._resolve_template_sheet_name(
            {sheet.Name for sheet in session.workbook.Worksheets},
            template_name,
        )
        template_sheet = session.sheet(actual_template_name)
        template_sheet.Copy(Before=session.workbook.Worksheets(1))
        render_sheet = session.workbook.ActiveSheet
        render_sheet.Name = sheet_name
        return render_sheet

    def _ensure_history_sheet(self, session: WorkbookSession):
        try:
            sheet = session.sheet(HISTORY_SHEET)
        except Exception:
            sheet = session.workbook.Worksheets.Add()
            sheet.Name = HISTORY_SHEET
            headers = [
                "操作时间", "操作类型", "模板", "流水号", "唯一识别码",
                "设备标识", "收款人名称", "收款人账号", "收款人开户行",
                "付款人名称", "付款人账号", "付款人开户行", "金额",
                "日期", "备注", "汇入地市", "汇入地区县",
            ]
            for index, header in enumerate(headers, start=1):
                sheet.Cells(1, index).Value = header
        return sheet

    def _validate_input_row(self, row: DataRow) -> None:
        required_values = {
            "收款人名称": row.payee_name,
            "收款人账号": row.payee_account,
            "收款人开户行": row.payee_bank,
            "付款人名称": row.payer_name,
            "付款人账号": row.payer_account,
            "金额": str(row.amount),
            "日期": row.date_text,
            "付款人开户行": row.payer_bank,
        }
        missing = [name for name, value in required_values.items() if not str(value).strip()]
        if missing:
            raise ValueError(f"请先填写：{', '.join(missing)}")

    def _write_data_row(self, sheet, row_index: int, row: DataRow) -> None:
        sheet.Cells(row_index, "B").Value = row.payee_name
        self._write_text_cell(sheet.Cells(row_index, "C"), row.payee_account)
        sheet.Cells(row_index, "D").Value = row.payee_bank
        sheet.Cells(row_index, "E").Value = row.payer_name
        self._write_text_cell(sheet.Cells(row_index, "F"), row.payer_account)
        sheet.Cells(row_index, "G").Value = float(row.amount)
        sheet.Cells(row_index, "H").Value = row.note
        sheet.Cells(row_index, "I").Value = row.date_text
        sheet.Cells(row_index, "J").Value = row.payer_bank
        sheet.Cells(row_index, "K").Value = row.in_city
        sheet.Cells(row_index, "L").Value = row.in_county
        sheet.Cells(row_index, "M").Value = row.extra_1
        sheet.Cells(row_index, "N").Value = row.extra_2

    def _fill_render_sheet(self, sheet, data_rows: list[DataRow], copy_mode: str) -> None:
        self._prepare_batch_template_blocks(sheet, len(data_rows), copy_mode)
        start_row = 2
        for item in data_rows:
            for offset in COPY_OFFSETS[copy_mode]:
                self._write_ticket_block(sheet, start_row + offset, item, copy_mode)
            start_row += ROW_STRIDE[copy_mode]
        self._set_batch_print_layout(sheet, len(data_rows), copy_mode)
        self._add_watermark(sheet, len(data_rows), copy_mode)
        self._add_operator_footer(sheet, len(data_rows), copy_mode)

    def _prepare_batch_template_blocks(self, sheet, row_count: int, copy_mode: str) -> None:
        if row_count <= 1:
            return

        stride = ROW_STRIDE[copy_mode]
        source_start = 2
        source_end = source_start + stride - 1
        source_rows = sheet.Rows(f"{source_start}:{source_end}")

        for index in range(1, row_count):
            target_start = source_start + stride * index
            target_end = target_start + stride - 1
            source_rows.Copy(Destination=sheet.Rows(f"{target_start}:{target_end}"))

    def _set_batch_print_layout(self, sheet, row_count: int, copy_mode: str) -> None:
        if row_count < 1:
            return

        stride = ROW_STRIDE[copy_mode]
        last_row = 1 + stride * row_count
        last_column = max(17, sheet.UsedRange.Column + sheet.UsedRange.Columns.Count - 1)
        print_area = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, last_column)).Address
        sheet.PageSetup.PrintArea = print_area

        if row_count > 1:
            try:
                sheet.ResetAllPageBreaks()
            except Exception:
                pass

    def _add_watermark(self, sheet, row_count: int, copy_mode: str) -> None:
        try:
            for shape in list(sheet.Shapes):
                shape_name = str(getattr(shape, "Name", ""))
                if shape_name.startswith("TicketWatermark") or shape_name.startswith("TicketOperatorFooter"):
                    shape.Delete()
        except Exception:
            pass

        if row_count < 1:
            return

        stride = ROW_STRIDE[copy_mode]
        last_column = max(17, sheet.UsedRange.Column + sheet.UsedRange.Columns.Count - 1)
        for index in range(row_count):
            block_start = 2 + stride * index
            block_end = block_start + stride - 1
            target_range = sheet.Range(sheet.Cells(block_start, 1), sheet.Cells(block_end, last_column))
            width = 220.0
            height = 44.0
            left = float(target_range.Left + (target_range.Width - width) / 2)
            top = float(target_range.Top + (target_range.Height - height) / 2)

            try:
                shape = sheet.Shapes.AddTextbox(1, left, top, width, height)
                shape.Name = f"TicketWatermark_{index + 1}"
                shape.Line.Visible = False
                shape.Fill.Visible = False
                shape.Rotation = -18
                text_range = shape.TextFrame2.TextRange
                text_range.Text = WATERMARK_TEXT
                text_range.Font.Size = 18
                text_range.Font.Name = "微软雅黑"
                text_range.Font.Bold = False
                text_range.Font.Fill.ForeColor.RGB = 0xC8B6D8
                text_range.Font.Fill.Transparency = 0.48
                shape.TextFrame2.VerticalAnchor = 3
                shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
            except Exception:
                try:
                    shape = sheet.Shapes.AddTextbox(1, left, top, width, height)
                    shape.Name = f"TicketWatermark_{index + 1}"
                    shape.Line.Visible = False
                    shape.Fill.Visible = False
                    shape.Rotation = -18
                    shape.TextFrame.Characters().Text = WATERMARK_TEXT
                    shape.TextFrame.HorizontalAlignment = -4108
                    shape.TextFrame.VerticalAlignment = -4108
                    shape.TextFrame.Characters().Font.Size = 18
                    shape.TextFrame.Characters().Font.Name = "微软雅黑"
                    shape.TextFrame.Characters().Font.Color = 0xC8B6D8
                except Exception:
                    logging.exception("添加水印失败")

        for index in range(1, row_count):
            sheet.HPageBreaks.Add(Before=sheet.Rows(2 + stride * index))
        try:
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = row_count
        except Exception:
            pass

    def _add_operator_footer(self, sheet, row_count: int, copy_mode: str) -> None:
        operator_name = self.store.get_operator_name().strip()
        if row_count < 1 or not operator_name:
            return

        stride = ROW_STRIDE[copy_mode]
        last_column = max(17, sheet.UsedRange.Column + sheet.UsedRange.Columns.Count - 1)
        for index in range(row_count):
            block_start = 2 + stride * index
            copy_height = 11
            width = 120.0
            height = 18.0
            footer_text = f"操作员：{operator_name}"
            for copy_index, offset in enumerate(COPY_OFFSETS[copy_mode], start=1):
                copy_start = block_start + offset
                copy_end = min(copy_start + copy_height - 1, block_start + stride - 1)
                target_range = sheet.Range(sheet.Cells(copy_start, 1), sheet.Cells(copy_end, last_column))
                security_row = copy_start + 9
                security_cell = sheet.Cells(security_row, "I")
                left = float(target_range.Left + 18)
                top = float(security_cell.Top - 1)
                try:
                    shape = sheet.Shapes.AddTextbox(1, left, top, width, height)
                    shape.Name = f"TicketOperatorFooter_{index + 1}_{copy_index}"
                    shape.Line.Visible = False
                    shape.Fill.Visible = False
                    shape.Rotation = 0
                    text_range = shape.TextFrame2.TextRange
                    text_range.Text = footer_text
                    text_range.Font.Size = 8
                    text_range.Font.Name = "微软雅黑"
                    text_range.Font.Fill.ForeColor.RGB = 0x7A7A7A
                    text_range.Font.Fill.Transparency = 0.1
                    shape.TextFrame2.VerticalAnchor = 3
                    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
                except Exception:
                    try:
                        shape = sheet.Shapes.AddTextbox(1, left, top, width, height)
                        shape.Name = f"TicketOperatorFooter_{index + 1}_{copy_index}"
                        shape.Line.Visible = False
                        shape.Fill.Visible = False
                        shape.TextFrame.Characters().Text = footer_text
                        shape.TextFrame.HorizontalAlignment = -4131
                        shape.TextFrame.VerticalAlignment = -4108
                        shape.TextFrame.Characters().Font.Size = 8
                        shape.TextFrame.Characters().Font.Name = "微软雅黑"
                        shape.TextFrame.Characters().Font.Color = 0x7A7A7A
                    except Exception:
                        logging.exception("添加操作员标记失败")

    def _write_ticket_block(self, sheet, row: int, item: DataRow, copy_mode: str) -> None:
        sheet.Cells(row, "E").Value = item.date_text
        sheet.Cells(row + 1, "G").Value = item.payee_name
        sheet.Cells(row + 1, "D").Value = item.payer_name
        self._write_text_cell(sheet.Cells(row + 2, "D"), item.payer_account)
        sheet.Cells(row + 4, "D").Value = self._amount_to_chinese_upper(item.amount)
        self._write_amount_boxes(sheet, row + 5, item.amount)
        sheet.Cells(row + 8, "B").Value = item.note

        if copy_mode == "2copy":
            self._write_text_cell(sheet.Cells(row + 2, "G"), self._group_account(item.payee_account))
            sheet.Cells(row + 3, "D").Value = item.payer_bank
            sheet.Cells(row + 3, "G").Value = item.payee_bank
            sheet.Cells(row + 3, "M").Value = f"{item.in_city}{item.in_county}"
        else:
            self._write_text_cell(sheet.Cells(row + 2, "G"), item.payee_account)
            sheet.Cells(row + 3, "D").Value = item.payee_bank
            sheet.Cells(row + 3, "G").Value = item.payer_bank

    def _write_amount_boxes(self, sheet, row: int, amount: Decimal) -> None:
        numeric_text = f"{amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):.2f}".replace(".", "")
        for column in range(8, 18):
            sheet.Cells(row, column).Value = ""
        target_column = 17
        for char in reversed(numeric_text):
            if target_column < 8:
                break
            sheet.Cells(row, target_column).Value = char
            target_column -= 1
        if target_column >= 8:
            sheet.Cells(row, target_column).Value = "¥"

    def _add_sequence_numbers(self, print_sheet, session: WorkbookSession, count: int, copy_mode: str) -> list[str]:
        full_serials = self.store.take_serials(count)
        self._write_serials_to_sheet(print_sheet, full_serials, copy_mode)
        return full_serials

    def _write_serials_to_sheet(self, print_sheet, full_serials: list[str], copy_mode: str) -> None:
        start_row = 2
        for full_serial in full_serials:
            for offset in COPY_OFFSETS[copy_mode]:
                row = start_row + offset
                print_sheet.Range(f"I{row}:N{row}").MergeCells = True
                cell = print_sheet.Cells(row, "I")
                cell.Value = full_serial
                cell.Font.Name = "黑体"
                cell.Font.Bold = True
            start_row += ROW_STRIDE[copy_mode]

    def _add_security_codes(
        self,
        print_sheet,
        rows: list[DataRow],
        full_serials: list[str],
        copy_mode: str,
    ) -> list[tuple[str, str]]:
        security_items = [
            self._make_security_code(row, full_serial)
            for row, full_serial in zip(rows, full_serials)
        ]

        start_row = 2
        for security_code, _printed_at in security_items:
            for offset in COPY_OFFSETS[copy_mode]:
                row = start_row + offset + 9
                cell = print_sheet.Cells(row, "I")
                cell.Value = f"唯一识别码：{security_code}"
                cell.Font.Size = 8
                cell.Font.Color = 8421504
            start_row += ROW_STRIDE[copy_mode]

        return security_items

    @staticmethod
    def _get_physical_address() -> str:
        mac_int = uuid.getnode()
        return ":".join(f"{(mac_int >> shift) & 0xFF:02X}" for shift in range(40, -1, -8))

    @classmethod
    def _make_security_code(cls, row: DataRow, full_serial: str) -> tuple[str, str]:
        physical_address = cls._get_physical_address()
        printed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
        raw_text = "|".join(
            [
                full_serial,
                f"{row.amount:.2f}",
                row.date_text,
                row.payee_account,
                row.payer_account,
                row.payee_name,
                physical_address,
                printed_at,
            ]
        )
        digest = hashlib.sha256(raw_text.encode("utf-8")).hexdigest().upper()
        return f"{digest[:4]}-{digest[4:8]}", printed_at

    def _append_history(
        self,
        session: WorkbookSession,
        row: DataRow,
        template_name: str,
        full_serial: str,
        action_type: str,
        security_code: str = "",
        printed_at: str = "",
    ) -> None:
        serial_number = full_serial
        device_id = os.environ.get("COMPUTERNAME", "UNKNOWN")
        physical_address = self._get_physical_address()
        operator_name = self.store.get_operator_name()
        self.store.append_print_record(
            {
                "operation_time": printed_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f"),
                "action_type": action_type,
                "template": template_name,
                "operator_name": operator_name,
                "serial_number": serial_number,
                "full_serial": full_serial,
                "security_code": security_code,
                "device_id": device_id,
                "physical_address": physical_address,
                "payee_name": row.payee_name,
                "payee_account": row.payee_account,
                "payee_bank": row.payee_bank,
                "payer_name": row.payer_name,
                "payer_account": row.payer_account,
                "payer_bank": row.payer_bank,
                "amount": f"{row.amount:.2f}",
                "date": row.date_text,
                "note": row.note,
                "in_city": row.in_city,
                "in_county": row.in_county,
            }
        )

    @staticmethod
    def _group_account(account: str) -> str:
        cleaned = account.replace(" ", "")
        return " ".join(cleaned[index:index + 4] for index in range(0, len(cleaned), 4))

    @staticmethod
    def _write_text_cell(cell, value: str) -> None:
        cell.NumberFormat = "@"
        cell.Value = str(value or "")

    @staticmethod
    def _to_decimal(value) -> Decimal:
        if value in (None, ""):
            return Decimal("0.00")
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    @staticmethod
    def _amount_to_chinese_upper(amount: Decimal) -> str:
        digits = "零壹贰叁肆伍陆柒捌玖"
        integer_units = ["", "拾", "佰", "仟"]
        section_units = ["", "万", "亿", "万亿"]

        def convert_section(number: int) -> str:
            if number == 0:
                return ""
            result = ""
            zero_pending = False
            unit_index = 0
            while number > 0:
                digit = number % 10
                if digit == 0:
                    if result:
                        zero_pending = True
                else:
                    piece = digits[digit] + integer_units[unit_index]
                    result = ("零" if zero_pending else "") + piece + result
                    zero_pending = False
                number //= 10
                unit_index += 1
            return result

        quantized = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        integer_part = int(quantized)
        decimal_part = int((quantized - integer_part) * 100)

        if integer_part == 0:
            integer_text = "零元"
        else:
            parts: list[str] = []
            section_index = 0
            need_zero = False
            while integer_part > 0:
                section = integer_part % 10000
                if section:
                    section_text = convert_section(section) + section_units[section_index]
                    if need_zero:
                        section_text = "零" + section_text
                        need_zero = False
                    parts.insert(0, section_text)
                    if section < 1000:
                        need_zero = True
                elif parts:
                    need_zero = True
                integer_part //= 10000
                section_index += 1
            integer_text = "".join(parts) + "元"

        jiao = decimal_part // 10
        fen = decimal_part % 10
        if jiao == 0 and fen == 0:
            return integer_text + "整"
        if jiao == 0:
            return integer_text + "零" + digits[fen] + "分"
        if fen == 0:
            return integer_text + digits[jiao] + "角整"
        return integer_text + digits[jiao] + "角" + digits[fen] + "分"


class TicketPrinterApp:
    def __init__(self, root: tk.Tk, workbook_path: Path) -> None:
        self.root = root
        self.root.title("财务票据打印")
        self.root.geometry("1180x760")
        self.root.minsize(1120, 720)
        self.service = TicketPrinterService(workbook_path)
        self._ensure_machine_prefix()

        today = date.today().strftime("%Y-%m-%d")
        self.available_template_sheets = self.service.list_template_sheets()
        self.template_sheet_to_label = {
            sheet: TEMPLATE_DISPLAY_ALIASES.get(sheet, sheet)
            for sheet in self.available_template_sheets
        }
        self.template_label_to_sheet = {
            label: sheet for sheet, label in self.template_sheet_to_label.items()
        }
        default_template_sheet = self.service.get_default_template()
        if default_template_sheet not in self.template_sheet_to_label:
            for actual_sheet in self.available_template_sheets:
                alias_label = TEMPLATE_DISPLAY_ALIASES.get(actual_sheet, actual_sheet)
                if default_template_sheet in {actual_sheet, alias_label}:
                    default_template_sheet = actual_sheet
                    break
            else:
                default_template_sheet = self.available_template_sheets[0] if self.available_template_sheets else DEFAULT_TEMPLATE_SHEETS[0]
        current_theme_name = self.service.get_theme_name()
        self.theme_label_to_name = {label: name for name, label in THEME_LABELS.items()}
        self.theme_name_to_label = dict(THEME_LABELS)
        self.theme_var = tk.StringVar(value=self.theme_name_to_label[current_theme_name])
        self.template_var = tk.StringVar(value=self.template_sheet_to_label[default_template_sheet])
        self.default_template_var = tk.BooleanVar(value=True)
        self.default_payer_var = tk.BooleanVar(value=False)
        self.operator_var = tk.StringVar(value=self.service.get_operator_name())
        self.status_var = tk.StringVar(value="就绪")
        self.sequence_display_var = tk.StringVar(value="当前票号：读取中...")
        self.template_display_var = tk.StringVar(
            value=f"当前模板：{self.template_sheet_to_label[default_template_sheet]}"
        )
        self.queue_count_var = tk.StringVar(value="待打印 0 笔")
        self.total_amount_var = tk.StringVar(value="合计 0.00 元")
        self.selection_summary_var = tk.StringVar(value="当前未选中记录")
        self._updating_default_template_flag = False
        self._updating_default_payer_flag = False
        self._updating_theme_flag = False
        self.current_theme_name = current_theme_name
        self.current_theme = THEME_PRESETS[current_theme_name]
        self.rows: list[DataRow] = []
        self.style = ttk.Style()
        self._entry_borders: list[tk.Frame] = []
        self._entry_widgets: list[tk.Entry] = []

        self.field_vars: dict[str, tk.StringVar] = {
            "payee_name": tk.StringVar(),
            "payee_account": tk.StringVar(),
            "payee_bank": tk.StringVar(),
            "payer_name": tk.StringVar(),
            "payer_account": tk.StringVar(),
            "amount": tk.StringVar(),
            "note": tk.StringVar(),
            "date_text": tk.StringVar(value=today),
            "payer_bank": tk.StringVar(),
            "in_city": tk.StringVar(),
            "in_county": tk.StringVar(),
            "extra_1": tk.StringVar(),
            "extra_2": tk.StringVar(),
        }

        self._configure_styles()
        self._load_default_payer()
        self._build_ui()
        self._refresh_sequence_display()
        self._refresh_default_template_flag()
        self._refresh_default_payer_flag()
        self._update_status_banner()

    def _ensure_machine_prefix(self) -> None:
        machine = self.service.get_machine_config()
        if machine and str(machine.get("prefix", "")).strip():
            return

        suggested_prefix = self.service.get_suggested_machine_prefix()
        while True:
            prefix = simpledialog.askstring(
                "设置流水号前置符",
                (
                    "检测到这台电脑第一次运行财务票据打印程序。\n\n"
                    f"请为本机设置 {MACHINE_PREFIX_LENGTH} 位流水号前置符。\n"
                    f"建议值：{suggested_prefix}\n\n"
                    "确认后将把本机流水号从 00001 开始。"
                ),
                parent=self.root,
                initialvalue=suggested_prefix,
            )
            if prefix is None:
                if messagebox.askyesno("退出程序", "未设置前置符，是否退出程序？", parent=self.root):
                    self.root.destroy()
                    raise SystemExit
                continue
            normalized_prefix = TicketDataStore._normalize_machine_prefix(prefix)
            if len(normalized_prefix) != MACHINE_PREFIX_LENGTH:
                messagebox.showerror(
                    "设置失败",
                    f"前置符必须是 {MACHINE_PREFIX_LENGTH} 位字母或数字。",
                    parent=self.root,
                )
                continue
            try:
                self.service.register_current_machine(normalized_prefix, reset_sequence=True)
            except Exception as exc:
                messagebox.showerror("设置失败", str(exc), parent=self.root)
                continue
            messagebox.showinfo(
                "设置成功",
                f"本机前置符已设置为 {normalized_prefix}，流水号将从 {normalized_prefix}{MIN_SERIAL_NUMBER:0{SERIAL_NUMBER_WIDTH}d} 开始。",
                parent=self.root,
            )
            break

    def _configure_styles(self) -> None:
        theme = self.current_theme
        self.root.configure(bg=theme["root_bg"])
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.style.configure(".", font=("Microsoft YaHei UI", 10))
        self.style.configure(".", background=theme["root_bg"], foreground=theme["text_fg"])
        self.style.configure("App.TFrame", background=theme["root_bg"])
        self.style.configure(
            "Card.TLabelframe",
            background=theme["panel_bg"],
            borderwidth=0,
            relief="flat",
            bordercolor=theme["border"],
            lightcolor=theme["border"],
            darkcolor=theme["border"],
        )
        self.style.configure(
            "Card.TLabelframe.Label",
            background=theme["panel_bg"],
            foreground=theme["text_fg"],
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        self.style.configure("Card.TFrame", background=theme["panel_bg"])
        self.style.configure("HeaderTitle.TLabel", background=theme["root_bg"], foreground=theme["title_fg"], font=("Microsoft YaHei UI", 13, "bold"))
        self.style.configure("HeaderSub.TLabel", background=theme["root_bg"], foreground=theme["muted_fg"], font=("Microsoft YaHei UI", 8))
        self.style.configure("SummaryTitle.TLabel", background=theme["panel_alt_bg"], foreground=theme["muted_fg"], font=("Microsoft YaHei UI", 9))
        self.style.configure("SummaryValue.TLabel", background=theme["panel_alt_bg"], foreground=theme["title_fg"], font=("Microsoft YaHei UI", 11, "bold"))
        self.style.configure("Info.TLabel", background=theme["panel_alt_bg"], foreground=theme["muted_fg"])
        self.style.configure(
            "TLabel",
            background=theme["root_bg"],
            foreground=theme["text_fg"],
        )
        self.style.configure(
            "TEntry",
            fieldbackground=theme["entry_bg"],
            bordercolor=theme["entry_border"],
            lightcolor=theme["entry_border"],
            darkcolor=theme["entry_border"],
            borderwidth=0,
            relief="flat",
            padding=4,
        )
        self.style.configure(
            "TCombobox",
            fieldbackground=theme["entry_bg"],
            background=theme["entry_bg"],
            bordercolor=theme["entry_border"],
            lightcolor=theme["entry_border"],
            darkcolor=theme["entry_border"],
            borderwidth=1,
            relief="solid",
            padding=2,
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", theme["entry_bg"])],
            selectbackground=[("readonly", theme["entry_bg"])],
            selectforeground=[("readonly", theme["title_fg"])],
        )
        self.style.configure(
            "Accent.TButton",
            padding=(12, 7),
            font=("Microsoft YaHei UI", 10, "bold"),
            background=theme["accent_bg"],
            foreground="#ffffff",
            borderwidth=0,
            focusthickness=0,
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", theme["accent_active"]), ("pressed", theme["accent_pressed"])],
            foreground=[("disabled", "#dbeafe")],
        )
        self.style.configure(
            "Secondary.TButton",
            padding=(12, 7),
            background=theme["secondary_bg"],
            foreground=theme["secondary_fg"],
            borderwidth=0,
            focusthickness=0,
        )
        self.style.map(
            "Secondary.TButton",
            background=[("active", theme["secondary_active"]), ("pressed", theme["secondary_pressed"])],
        )
        self.style.configure(
            "Danger.TButton",
            padding=(12, 7),
            background=theme["secondary_bg"],
            foreground=theme["secondary_fg"],
            borderwidth=0,
            focusthickness=0,
        )
        self.style.map(
            "Danger.TButton",
            background=[("active", theme["secondary_active"]), ("pressed", theme["secondary_pressed"])],
        )
        self.style.configure(
            "Treeview",
            rowheight=28,
            font=("Microsoft YaHei UI", 10),
            background=theme["table_even_bg"],
            fieldbackground=theme["table_even_bg"],
            borderwidth=0,
            relief="flat",
        )
        self.style.configure(
            "Treeview.Heading",
            font=("Microsoft YaHei UI", 10, "bold"),
            background=theme["table_header_bg"],
            foreground=theme["table_header_fg"],
            relief="flat",
        )
        self.style.map(
            "Treeview",
            background=[("selected", theme["table_selected_bg"])],
            foreground=[("selected", theme["title_fg"])],
        )
        self.style.map(
            "Treeview.Heading",
            background=[("active", theme["table_header_bg"])],
        )
        self._apply_runtime_theme()

    def _create_text_entry(
        self,
        parent,
        textvariable: tk.StringVar,
        *,
        width: int | None = None,
    ) -> tuple[tk.Frame, tk.Entry]:
        outer = tk.Frame(parent, bg=self.current_theme["entry_border"], bd=0, highlightthickness=0)
        entry = tk.Entry(
            outer,
            textvariable=textvariable,
            width=width,
            relief="flat",
            bd=0,
            highlightthickness=0,
            bg=self.current_theme["entry_bg"],
            fg=self.current_theme["title_fg"],
            insertbackground=self.current_theme["title_fg"],
            font=("Microsoft YaHei UI", 10),
        )
        entry.pack(fill="both", expand=True, padx=1, pady=1, ipady=3)
        self._entry_borders.append(outer)
        self._entry_widgets.append(entry)
        return outer, entry

    def _create_tick_checkbutton(self, parent, text: str, variable: tk.BooleanVar, command):
        checkbox = tk.Checkbutton(
            parent,
            text=text,
            variable=variable,
            command=command,
            anchor="w",
            padx=2,
            pady=0,
            font=("Microsoft YaHei UI", 9),
            bg=self.current_theme["panel_bg"],
            fg=self.current_theme["text_fg"],
            activebackground=self.current_theme["panel_bg"],
            activeforeground=self.current_theme["title_fg"],
            selectcolor=self.current_theme["panel_bg"],
            relief="flat",
            bd=0,
            highlightthickness=0,
        )
        return checkbox

    def _load_default_payer(self) -> None:
        payer = self.service.get_default_payer()
        if payer.get("name"):
            self.field_vars["payer_name"].set(payer.get("name", ""))
            self.field_vars["payer_bank"].set(payer.get("bank", ""))
            self.field_vars["payer_account"].set(payer.get("account", ""))

    def _show_program_help(self) -> None:
        messagebox.showinfo("功能说明", PROGRAM_HELP_TEXT, parent=self.root)

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        header_frame = ttk.Frame(self.root, style="App.TFrame", padding=(14, 6, 14, 2))
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.columnconfigure(0, weight=1)
        header_frame.columnconfigure(1, weight=0)
        ttk.Label(header_frame, text="财务票据打印工作台", style="HeaderTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header_frame, text="主题", style="Info.TLabel").grid(row=0, column=1, sticky="e", padx=(0, 6))
        ttk.Combobox(
            header_frame,
            textvariable=self.theme_var,
            values=list(self.theme_label_to_name.keys()),
            state="readonly",
            width=18,
        ).grid(row=0, column=2, sticky="e")
        ttk.Button(
            header_frame,
            text="功能说明",
            style="Secondary.TButton",
            command=self._show_program_help,
        ).grid(row=0, column=3, sticky="e", padx=(8, 0))
        self.theme_var.trace_add("write", self._on_theme_change)
        ttk.Label(
            header_frame,
            text="内部结算专用 · 模板化打印 · 唯一识别码核验",
            style="HeaderSub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 3))
        self.status_banner = tk.Label(
            header_frame,
            textvariable=self.status_var,
            anchor="w",
            padx=2,
            pady=0,
            font=("Microsoft YaHei UI", 8),
            bg="#eef6ff",
            fg="#5d7fa3",
        )
        self.status_banner.grid(row=2, column=0, sticky="w")
        self.status_var.trace_add("write", self._update_status_banner)

        entry_frame = ttk.LabelFrame(self.root, text="录入信息", padding=14, style="Card.TLabelframe")
        entry_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=(6, 8))
        for column in range(6):
            entry_frame.columnconfigure(column, weight=1 if column in {1, 3, 5} else 0)

        fields = [
            ("收款人名称", "payee_name", 0, 0, 24),
            ("收款人账号", "payee_account", 0, 2, 24),
            ("收款人开户行", "payee_bank", 0, 4, 24),
            ("付款人名称", "payer_name", 1, 0, 24),
            ("付款人账号", "payer_account", 1, 2, 24),
            ("付款人开户行", "payer_bank", 1, 4, 24),
            ("金额", "amount", 2, 0, 14),
            ("日期", "date_text", 2, 2, 14),
            ("备注", "note", 2, 4, 24),
            ("汇入地-市", "in_city", 3, 0, 14),
            ("汇入地-县", "in_county", 3, 2, 14),
        ]
        for label, key, row_index, column_index, width in fields:
            ttk.Label(entry_frame, text=label).grid(
                row=row_index,
                column=column_index,
                sticky="e",
                padx=(0, 6),
                pady=5,
            )
            if key == "payer_name":
                payer_name_frame = ttk.Frame(entry_frame, style="Card.TFrame")
                payer_name_frame.grid(
                    row=row_index,
                    column=column_index + 1,
                    sticky="ew",
                    padx=(0, 14),
                    pady=5,
                )
                payer_name_frame.columnconfigure(0, weight=1)
                entry_outer, _entry = self._create_text_entry(payer_name_frame, self.field_vars[key], width=18)
                entry_outer.grid(row=0, column=0, sticky="ew")
                self.default_payer_checkbutton = self._create_tick_checkbutton(
                    payer_name_frame,
                    "默认",
                    self.default_payer_var,
                    self._toggle_default_payer,
                )
                self.default_payer_checkbutton.grid(row=0, column=1, sticky="w", padx=(8, 0))
            else:
                entry_outer, _entry = self._create_text_entry(entry_frame, self.field_vars[key], width=width)
                entry_outer.grid(
                    row=row_index,
                    column=column_index + 1,
                    sticky="ew",
                    padx=(0, 14),
                    pady=5,
                )

        account_buttons = ttk.Frame(entry_frame, style="Card.TFrame")
        account_buttons.grid(row=4, column=0, columnspan=6, sticky="w", pady=(8, 0))
        ttk.Button(account_buttons, text="添加付款人", style="Secondary.TButton", command=lambda: self._query_account("payer")).grid(
            row=0,
            column=0,
            padx=4,
        )
        ttk.Button(account_buttons, text="添加收款人", style="Secondary.TButton", command=lambda: self._query_account("payee")).grid(
            row=0,
            column=1,
            padx=4,
        )

        entry_buttons = ttk.Frame(entry_frame, style="Card.TFrame")
        entry_buttons.grid(row=5, column=0, columnspan=6, sticky="e", pady=(8, 0))
        ttk.Button(entry_buttons, text="加入待打队列", style="Accent.TButton", command=self._add_row).grid(row=0, column=0, padx=4)
        ttk.Button(entry_buttons, text="修改选中", style="Secondary.TButton", command=self._update_selected_row).grid(row=0, column=1, padx=4)
        ttk.Button(entry_buttons, text="删除选中", style="Danger.TButton", command=self._delete_selected_row).grid(row=0, column=2, padx=4)
        ttk.Button(entry_buttons, text="清空录入", style="Secondary.TButton", command=self._clear_form).grid(row=0, column=3, padx=4)

        table_frame = ttk.LabelFrame(self.root, text="待打印数据（可选中一条或多条后进行预览或打印）", padding=12, style="Card.TLabelframe")
        table_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=8)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("status", "payee", "account", "payer", "amount", "date", "note")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings", height=10, selectmode="extended")
        headings = {
            "status": "状态",
            "payee": "收款人",
            "account": "收款账号",
            "payer": "付款人",
            "amount": "金额",
            "date": "日期",
            "note": "备注",
        }
        widths = {
            "status": 90,
            "payee": 150,
            "account": 190,
            "payer": 150,
            "amount": 90,
            "date": 100,
            "note": 210,
        }
        for column, heading in headings.items():
            self.table.heading(column, text=heading)
            self.table.column(column, width=widths[column], anchor="w")
        self.table.grid(row=0, column=0, sticky="nsew")
        self.table.tag_configure("odd", background="#ffffff")
        self.table.tag_configure("even", background="#f6f9fc")
        self.table.tag_configure("status_done", foreground="#1d7a46")
        self.table.tag_configure("status_reprint", foreground="#b15a00")
        self.table.bind("<<TreeviewSelect>>", self._on_table_select)
        self.table.bind("<Double-1>", self._preview_selected)
        self.table.bind("<Control-a>", self._select_all_rows)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.table.configure(yscrollcommand=scrollbar.set)

        action_frame = ttk.Frame(self.root, style="App.TFrame", padding=(12, 4, 12, 12))
        action_frame.grid(row=3, column=0, sticky="ew")
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)

        tool_frame = ttk.LabelFrame(action_frame, text="打印操作", padding=12, style="Card.TLabelframe")
        tool_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        summary_frame = ttk.LabelFrame(action_frame, text="当前摘要", padding=12, style="Card.TLabelframe")
        summary_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        for column in range(3):
            summary_frame.columnconfigure(column, weight=1)

        ttk.Label(tool_frame, text="模板", style="Info.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=(0, 10))
        ttk.Combobox(
            tool_frame,
            textvariable=self.template_var,
            values=list(self.template_label_to_sheet.keys()),
            state="readonly",
            width=18,
        ).grid(row=0, column=1, sticky="w", padx=(0, 10), pady=(0, 10))
        self.template_var.trace_add("write", self._on_template_change)
        self.default_template_checkbutton = self._create_tick_checkbutton(
            tool_frame,
            "默认",
            self.default_template_var,
            self._toggle_default_template,
        )
        self.default_template_checkbutton.grid(row=0, column=2, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Label(tool_frame, text="操作员", style="Info.TLabel").grid(row=0, column=3, sticky="w", padx=(0, 6), pady=(0, 10))
        operator_outer, _operator_entry = self._create_text_entry(tool_frame, self.operator_var, width=12)
        operator_outer.grid(row=0, column=4, sticky="w", padx=(0, 8), pady=(0, 10))
        ttk.Button(tool_frame, text="保存操作员", style="Secondary.TButton", command=self._save_operator_name).grid(
            row=0,
            column=5,
            padx=4,
            pady=(0, 10),
        )
        ttk.Button(tool_frame, text="已打印记录", style="Secondary.TButton", command=self._query_print_records).grid(
            row=1,
            column=0,
            padx=4,
            pady=4,
            sticky="w",
        )
        ttk.Button(tool_frame, text="核验票据", style="Secondary.TButton", command=self._verify_print_record_dialog).grid(
            row=1,
            column=1,
            padx=4,
            pady=4,
            sticky="w",
        )
        ttk.Button(tool_frame, text="全选表格", style="Secondary.TButton", command=self._select_all_rows).grid(
            row=1,
            column=2,
            padx=4,
            pady=4,
            sticky="w",
        )
        ttk.Button(tool_frame, text="刷新序号", style="Secondary.TButton", command=self._refresh_sequence_display).grid(
            row=1,
            column=3,
            padx=4,
            pady=4,
            sticky="w",
        )
        ttk.Button(tool_frame, text="打印预览", style="Secondary.TButton", command=self._preview_selected).grid(
            row=1,
            column=4,
            padx=4,
            pady=4,
            sticky="w",
        )
        ttk.Button(tool_frame, text="直接打印", style="Accent.TButton", command=self._print_selected).grid(
            row=1,
            column=5,
            padx=4,
            pady=4,
            sticky="w",
        )

        ttk.Label(summary_frame, text="待打印笔数", style="SummaryTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(summary_frame, text="待打合计", style="SummaryTitle.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(summary_frame, text="当前选择", style="SummaryTitle.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Label(summary_frame, textvariable=self.queue_count_var, style="SummaryValue.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 10))
        ttk.Label(summary_frame, textvariable=self.total_amount_var, style="SummaryValue.TLabel").grid(row=1, column=1, sticky="w", pady=(4, 10))
        ttk.Label(summary_frame, textvariable=self.selection_summary_var, style="SummaryValue.TLabel").grid(row=1, column=2, sticky="w", pady=(4, 10))
        ttk.Label(summary_frame, textvariable=self.sequence_display_var, style="Info.TLabel").grid(
            row=2,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(4, 0),
        )
        ttk.Label(summary_frame, textvariable=self.template_display_var, style="Info.TLabel").grid(
            row=3,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(4, 0),
        )

    def _collect_form_data(self) -> DataRow:
        amount_text = self.field_vars["amount"].get().strip()
        if not amount_text:
            raise ValueError("请填写金额。")
        return DataRow(
            payee_name=self.field_vars["payee_name"].get().strip(),
            payee_account=self.field_vars["payee_account"].get().strip(),
            payee_bank=self.field_vars["payee_bank"].get().strip(),
            payer_name=self.field_vars["payer_name"].get().strip(),
            payer_account=self.field_vars["payer_account"].get().strip(),
            amount=TicketPrinterService._to_decimal(amount_text),
            note=self.field_vars["note"].get().strip(),
            date_text=self.field_vars["date_text"].get().strip(),
            payer_bank=self.field_vars["payer_bank"].get().strip(),
            in_city=self.field_vars["in_city"].get().strip(),
            in_county=self.field_vars["in_county"].get().strip(),
            extra_1=self.field_vars["extra_1"].get().strip(),
            extra_2=self.field_vars["extra_2"].get().strip(),
        )

    def _add_row(self) -> None:
        try:
            row = self._collect_form_data()
            self.service._validate_input_row(row)
        except Exception as exc:
            messagebox.showerror("录入有误", str(exc))
            return
        self.rows.append(row)
        self._refresh_table()
        if hasattr(self, "table"):
            self.table.selection_remove(*self.table.selection())
            self.table.focus("")
        self._clear_form()
        self.status_var.set("已加入表格")

    def _update_selected_row(self) -> None:
        index = self._selected_index()
        if index is None:
            messagebox.showinfo("提示", "请先在下面表格里选中一条数据。")
            return
        try:
            row = self._collect_form_data()
            self.service._validate_input_row(row)
        except Exception as exc:
            messagebox.showerror("录入有误", str(exc))
            return
        current_row = self.rows[index]
        row.printed_count = current_row.printed_count
        row.last_serial = current_row.last_serial
        row.last_security_code = current_row.last_security_code
        row.last_printed_at = current_row.last_printed_at
        self.rows[index] = row
        self._refresh_table(select_index=index)
        self.status_var.set("已修改选中数据")

    def _delete_selected_row(self) -> None:
        index = self._selected_index()
        if index is None:
            messagebox.showinfo("提示", "请先在下面表格里选中一条数据。")
            return
        del self.rows[index]
        self._refresh_table()
        self.status_var.set("已删除选中数据")

    def _generate_test_rows(self) -> None:
        today = date.today().strftime("%Y-%m-%d")
        self.rows = [
            DataRow(
                payee_name="上海测试供应商有限公司",
                payee_account="6222020202020202020",
                payee_bank="中国工商银行上海分行",
                payer_name="本单位测试付款户",
                payer_account="1002003004005006007",
                amount=Decimal("1280.50"),
                note="办公用品测试",
                date_text=today,
                payer_bank="中国农业银行本地支行",
                in_city="上海市",
                in_county="浦东新区",
            ),
            DataRow(
                payee_name="北京样例科技有限公司",
                payee_account="6228480402564890018",
                payee_bank="中国农业银行北京分行",
                payer_name="本单位测试付款户",
                payer_account="1002003004005006007",
                amount=Decimal("35600.00"),
                note="软件服务费测试",
                date_text=today,
                payer_bank="中国农业银行本地支行",
                in_city="北京市",
                in_county="海淀区",
            ),
            DataRow(
                payee_name="杭州测试材料经营部",
                payee_account="6217001210012345678",
                payee_bank="中国建设银行杭州支行",
                payer_name="本单位测试付款户",
                payer_account="1002003004005006007",
                amount=Decimal("987.65"),
                note="材料费测试",
                date_text=today,
                payer_bank="中国农业银行本地支行",
                in_city="杭州市",
                in_county="西湖区",
            ),
            DataRow(
                payee_name="广州示例劳务有限公司",
                payee_account="6212260200011122233",
                payee_bank="中国银行广州分行",
                payer_name="本单位测试付款户",
                payer_account="1002003004005006007",
                amount=Decimal("12000.00"),
                note="劳务费测试",
                date_text=today,
                payer_bank="中国农业银行本地支行",
                in_city="广州市",
                in_county="天河区",
            ),
            DataRow(
                payee_name="南京测试印刷厂",
                payee_account="6227003322110099887",
                payee_bank="交通银行南京支行",
                payer_name="本单位测试付款户",
                payer_account="1002003004005006007",
                amount=Decimal("4567.89"),
                note="印刷费测试",
                date_text=today,
                payer_bank="中国农业银行本地支行",
                in_city="南京市",
                in_county="鼓楼区",
            ),
        ]
        for row in self.rows:
            self.service.save_account("payer", row.payer_name, row.payer_bank, row.payer_account)
            self.service.save_account("payee", row.payee_name, row.payee_bank, row.payee_account)
        self._refresh_table(select_index=0)
        self.status_var.set("已生成 5 条测试数据，并写入付款人/收款人账户库")

    def _preview_selected(self) -> None:
        selected_rows = self._selected_rows()
        if not selected_rows:
            messagebox.showinfo("提示", "请先在下面表格里选中至少一条数据。")
            return

        def action() -> None:
            self.service.preview_rows(selected_rows, self._selected_template_sheet())
            self._refresh_sequence_display()

        message = "已打开打印预览" if len(selected_rows) == 1 else f"已打开 {len(selected_rows)} 条数据的打印预览"
        self._run_action(message, action)

    def _print_selected(self) -> None:
        selected_indices = self._selected_indices()
        if not selected_indices:
            messagebox.showinfo("提示", "请先在下面表格里选中至少一条数据。")
            return
        selected_rows = [self.rows[index] for index in selected_indices if 0 <= index < len(self.rows)]
        if not self._confirm_print_rows(selected_rows):
            return

        def action() -> None:
            serial_plan = [row.last_serial if row.printed_count > 0 and row.last_serial else "" for row in selected_rows]
            action_plan = ["重打" if row.printed_count > 0 else "打印" for row in selected_rows]
            print_results = self.service.print_rows_with_plan(
                selected_rows,
                self._selected_template_sheet(),
                full_serials=serial_plan,
                action_types=action_plan,
            )
            if print_results:
                for index, result in zip(selected_indices, print_results):
                    row = self.rows[index]
                    row.printed_count += 1
                    row.last_serial = str(result.get("full_serial", ""))
                    row.last_security_code = str(result.get("security_code", ""))
                    row.last_printed_at = str(result.get("printed_at", ""))
                self._refresh_table(select_index=selected_indices[0] if selected_indices else None)
                self._refresh_sequence_display()
            else:
                self.status_var.set("已取消打印")

        message = "已发送到打印机" if len(selected_rows) == 1 else f"已发送 {len(selected_rows)} 条数据到打印机"
        self._run_action(message, action)

    def _confirm_print_rows(self, rows: list[DataRow]) -> bool:
        prefix, current_serial = self.service.get_sequence_state()
        next_serial = current_serial
        warnings = self._build_print_warnings(rows)
        lines = [f"模板：{self.template_var.get()}", f"打印条数：{len(rows)}", ""]
        if warnings:
            lines.append("注意事项：")
            lines.extend(f"- {item}" for item in warnings)
            lines.append("")
        for index, row in enumerate(rows, start=1):
            if row.printed_count > 0 and row.last_serial:
                action_type = "重打"
                serial_text = row.last_serial
            else:
                action_type = "打印"
                serial_text = self.service.format_serial(prefix, next_serial)
                next_serial += 1
            lines.extend(
                [
                    f"{index}. [{action_type}] {row.payee_name} ← {row.payer_name}",
                    f"   金额：{row.amount:.2f}    流水号：{serial_text}",
                    f"   日期：{row.date_text}    备注：{row.note or '-'}",
                    "",
                ]
            )
        lines.append("确认后将开始直接打印。")
        return messagebox.askyesno("打印确认", "\n".join(lines))

    def _build_print_warnings(self, rows: list[DataRow]) -> list[str]:
        def _row_value(row: DataRow | dict, key: str, default=None):
            if isinstance(row, dict):
                return row.get(key, default)
            return getattr(row, key, default)

        def _row_amount(row: DataRow | dict) -> Decimal:
            value = _row_value(row, "amount", Decimal("0"))
            if isinstance(value, Decimal):
                return value
            try:
                return Decimal(str(value or "0"))
            except Exception:
                return Decimal("0")

        warnings: list[str] = []
        for row in rows:
            amount = _row_amount(row)
            payee_name = str(_row_value(row, "payee_name", "") or "").strip()
            if amount >= LARGE_AMOUNT_WARNING:
                warnings.append(f"{payee_name or '当前收款人'} 金额 {amount:.2f} 已达到大额提醒阈值。")

        payee_counts: dict[tuple[str, str], int] = {}
        for row in rows:
            payee_name = str(_row_value(row, "payee_name", "") or "").strip()
            date_text = str(_row_value(row, "date_text", "") or "").strip()
            key = (payee_name, date_text)
            payee_counts[key] = payee_counts.get(key, 0) + 1
        for (payee_name, date_text), count in payee_counts.items():
            if payee_name and count > 1:
                warnings.append(f"{payee_name} 在本次打印中有 {count} 笔，日期均为 {date_text}。")

        history = self.service.get_all_print_records()
        for row in rows:
            payee_name = str(_row_value(row, "payee_name", "") or "").strip()
            date_text = str(_row_value(row, "date_text", "") or "").strip()
            same_day_count = sum(
                1
                for item in history
                if str(item.get("payee_name", "")).strip() == payee_name
                and str(item.get("date", "")).strip() == date_text
            )
            if same_day_count > 0:
                warnings.append(f"{payee_name or '当前收款人'} 在 {date_text} 已有 {same_day_count} 条历史打印记录。")
        return list(dict.fromkeys(warnings))

    def _clear_form(self) -> None:
        for key, var in self.field_vars.items():
            var.set(date.today().strftime("%Y-%m-%d") if key == "date_text" else "")
        self._load_default_payer()
        self._refresh_default_payer_flag()
        self.status_var.set("录入区已清空")

    def _save_account(
        self,
        account_type: str,
        name: str | None = None,
        bank: str | None = None,
        account: str | None = None,
    ) -> None:
        fields = self._account_fields(account_type)
        try:
            self.service.save_account(
                account_type,
                name if name is not None else self.field_vars[fields["name"]].get(),
                bank if bank is not None else self.field_vars[fields["bank"]].get(),
                account if account is not None else self.field_vars[fields["account"]].get(),
            )
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))
            return
        label = "付款人" if account_type == "payer" else "收款人"
        self.status_var.set(f"已保存{label}账户")

    def _query_account(self, account_type: str) -> None:
        label = "付款人" if account_type == "payer" else "收款人"
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{label}维护")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("980x560")
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(2, weight=1)

        search_var = tk.StringVar()
        name_var = tk.StringVar()
        bank_var = tk.StringVar()
        account_var = tk.StringVar()
        favorite_var = tk.BooleanVar(value=False)
        editing_account_var = tk.StringVar(value="")
        edit_mode_var = tk.StringVar(value="当前为新增模式")
        save_button_text_var = tk.StringVar(value=f"新增{label}")
        top_frame = ttk.Frame(dialog, padding=(12, 12, 12, 8))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)
        ttk.Label(top_frame, text=f"{label}查询").grid(row=0, column=0, padx=(0, 8))
        search_outer, search_entry = self._create_text_entry(top_frame, search_var)
        search_outer.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        form_frame = ttk.LabelFrame(dialog, text=f"{label}录入", padding=12)
        form_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        form_frame.columnconfigure(1, weight=1)
        form_frame.columnconfigure(3, weight=1)
        form_frame.columnconfigure(5, weight=1)
        ttk.Label(form_frame, text=f"{label}名称").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=4)
        name_outer, _name_entry = self._create_text_entry(form_frame, name_var)
        name_outer.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=4)
        ttk.Label(form_frame, text="开户行").grid(row=0, column=2, sticky="e", padx=(0, 6), pady=4)
        bank_outer, _bank_entry = self._create_text_entry(form_frame, bank_var)
        bank_outer.grid(row=0, column=3, sticky="ew", padx=(0, 10), pady=4)
        ttk.Label(form_frame, text="账号").grid(row=0, column=4, sticky="e", padx=(0, 6), pady=4)
        account_outer, _account_entry = self._create_text_entry(form_frame, account_var)
        account_outer.grid(row=0, column=5, sticky="ew", pady=4)
        if account_type == "payer":
            ttk.Checkbutton(form_frame, text="常用付款人", variable=favorite_var).grid(
                row=1,
                column=0,
                columnspan=2,
                sticky="w",
                pady=(6, 0),
            )

        form_buttons = ttk.Frame(form_frame)
        form_buttons.grid(row=2, column=0, columnspan=6, sticky="w", pady=(8, 0))
        ttk.Label(form_frame, textvariable=edit_mode_var, foreground="#8a4b00").grid(
            row=3,
            column=0,
            columnspan=6,
            sticky="w",
            pady=(8, 0),
        )

        tree_columns = ("favorite", "name", "bank", "account", "updated") if account_type == "payer" else ("name", "bank", "account", "updated")
        tree = ttk.Treeview(dialog, columns=tree_columns, show="headings", height=12)
        headings = {"name": f"{label}名称", "bank": "开户行", "account": "账号", "updated": "更新时间"}
        widths = {"name": 180, "bank": 250, "account": 240, "updated": 140}
        if account_type == "payer":
            headings = {"favorite": "常用", **headings}
            widths = {"favorite": 60, **widths}
        for column, heading in headings.items():
            tree.heading(column, text=heading)
            tree.column(column, width=widths[column], anchor="w")
        tree.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))

        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=tree.yview)
        scrollbar.grid(row=2, column=1, sticky="ns", pady=(0, 12))
        tree.configure(yscrollcommand=scrollbar.set)

        def load_records() -> list[dict]:
            return self.service.find_accounts(account_type, search_var.get().strip())

        def refresh_tree() -> None:
            records = load_records()
            tree.delete(*tree.get_children())
            selected_index = None
            for index, item in enumerate(records):
                if editing_account_var.get().strip() and str(item.get("account", "")).strip() == editing_account_var.get().strip():
                    selected_index = str(index)
                values = (
                    ("★" if item.get("is_favorite") else ""),
                    item.get("name", ""),
                    item.get("bank", ""),
                    item.get("account", ""),
                    item.get("updated_at", ""),
                ) if account_type == "payer" else (
                    item.get("name", ""),
                    item.get("bank", ""),
                    item.get("account", ""),
                    item.get("updated_at", ""),
                )
                tree.insert(
                    "",
                    "end",
                    iid=str(index),
                    values=values,
                )
            tree.records = records  # type: ignore[attr-defined]
            if selected_index is not None:
                tree.selection_set(selected_index)
                tree.focus(selected_index)

        def selected_record() -> dict | None:
            picked = tree.selection()
            if not picked:
                return None
            records = getattr(tree, "records", [])
            if not records:
                return None
            return records[int(picked[0])]

        def write_back(_event=None) -> None:
            record = selected_record()
            if not record:
                messagebox.showinfo("提示", "请先选择一条账户信息。", parent=dialog)
                return
            self._fill_account_fields(account_type, record)
            if account_type == "payer":
                self._refresh_default_payer_flag()
            self.status_var.set(f"已带出{label}账户信息")
            dialog.destroy()

        def load_selected_to_form(_event=None) -> None:
            record = selected_record()
            if not record:
                return
            name_var.set(str(record.get("name", "")))
            bank_var.set(str(record.get("bank", "")))
            account_var.set(str(record.get("account", "")))
            favorite_var.set(bool(record.get("is_favorite")))
            editing_account_var.set(str(record.get("account", "")))
            edit_mode_var.set(f"正在修改账号：{record.get('account', '')}")
            save_button_text_var.set("保存修改")

        def clear_form() -> None:
            name_var.set("")
            bank_var.set("")
            account_var.set("")
            favorite_var.set(False)
            editing_account_var.set("")
            edit_mode_var.set("当前为新增模式")
            save_button_text_var.set(f"新增{label}")

        def save_current() -> None:
            original_account = editing_account_var.get().strip()
            if original_account:
                try:
                    self.service.update_account(
                        account_type,
                        original_account,
                        name_var.get(),
                        bank_var.get(),
                        account_var.get(),
                    )
                except Exception as exc:
                    messagebox.showerror("保存失败", str(exc), parent=dialog)
                    return
                if account_type == "payer":
                    self.service.toggle_account_favorite(account_type, account_var.get().strip(), favorite_var.get())
                self.status_var.set(f"已修改{label}账户")
                editing_account_var.set(account_var.get().strip())
                edit_mode_var.set(f"修改成功：{account_var.get().strip()}")
                messagebox.showinfo("修改成功", f"{label}账户已保存修改。", parent=dialog)
            else:
                self._save_account(account_type, name_var.get(), bank_var.get(), account_var.get())
                if account_type == "payer":
                    self.service.toggle_account_favorite(account_type, account_var.get().strip(), favorite_var.get())
                edit_mode_var.set(f"新增成功：{account_var.get().strip()}")
                messagebox.showinfo("保存成功", f"{label}账户已新增保存。", parent=dialog)
            refresh_tree()

        def delete_selected() -> None:
            record = selected_record()
            if not record:
                messagebox.showinfo("提示", "请先选择一条账户信息。", parent=dialog)
                return
            if not messagebox.askyesno("确认删除", f"确认删除该{label}账户吗？", parent=dialog):
                return
            try:
                self.service.delete_account(account_type, str(record.get("account", "")))
            except Exception as exc:
                messagebox.showerror("删除失败", str(exc), parent=dialog)
                return
            refresh_tree()
            clear_form()
            self.status_var.set(f"已删除{label}账户")

        def import_payees() -> None:
            file_path = filedialog.askopenfilename(
                parent=dialog,
                title="选择收款人导入文件",
                filetypes=[("CSV 文件", "*.csv"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
            )
            if not file_path:
                return
            try:
                records = self._read_import_accounts(Path(file_path))
                imported = self.service.import_accounts("payee", records)
            except Exception as exc:
                messagebox.showerror("导入失败", str(exc), parent=dialog)
                return
            refresh_tree()
            self.status_var.set(f"已批量导入 {imported} 条收款人")
            messagebox.showinfo("导入完成", f"成功导入 {imported} 条收款人记录。", parent=dialog)

        def export_payee_template() -> None:
            file_path = filedialog.asksaveasfilename(
                parent=dialog,
                title="下载收款人导入模板",
                defaultextension=".csv",
                initialfile="收款人导入模板.csv",
                filetypes=[("CSV 文件", "*.csv")],
            )
            if not file_path:
                return
            try:
                self._write_import_template(Path(file_path))
            except Exception as exc:
                messagebox.showerror("下载失败", str(exc), parent=dialog)
                return
            self.status_var.set("已下载收款人导入模板")
            messagebox.showinfo("下载完成", "收款人导入模板已生成。", parent=dialog)

        ttk.Button(top_frame, text="查询", command=refresh_tree).grid(row=0, column=2, padx=4)
        ttk.Button(top_frame, text="关闭", command=dialog.destroy).grid(row=0, column=3, padx=(12, 0))

        ttk.Button(form_buttons, textvariable=save_button_text_var, command=save_current).grid(row=0, column=0, padx=4)
        ttk.Button(form_buttons, text="清空录入", command=clear_form).grid(row=0, column=1, padx=4)
        ttk.Button(form_buttons, text="删除选中", command=delete_selected).grid(row=0, column=2, padx=4)
        ttk.Button(form_buttons, text="带回主界面", command=write_back).grid(row=0, column=3, padx=4)
        if account_type == "payer":
            ttk.Button(form_buttons, text="切换常用", command=lambda: self._toggle_selected_account_favorite(tree, account_type, refresh_tree, favorite_var)).grid(row=0, column=4, padx=4)
        if account_type == "payee":
            ttk.Button(form_buttons, text="批量导入", command=import_payees).grid(row=0, column=4, padx=4)
            ttk.Button(form_buttons, text="下载导入模板", command=export_payee_template).grid(row=0, column=5, padx=4)

        tip = ttk.Label(dialog, text="双击列表可回填到当前界面", foreground="#666666")
        tip.grid(row=3, column=0, sticky="w", padx=12, pady=(0, 12))

        tree.bind("<Double-1>", write_back)
        tree.bind("<<TreeviewSelect>>", load_selected_to_form)
        search_entry.bind("<Return>", lambda _event: refresh_tree())
        refresh_tree()
        search_entry.focus_set()
        dialog.wait_window()

    def _fill_account_fields(self, account_type: str, record: dict) -> None:
        fields = self._account_fields(account_type)
        self.field_vars[fields["name"]].set(str(record.get("name", "")))
        self.field_vars[fields["bank"]].set(str(record.get("bank", "")))
        self.field_vars[fields["account"]].set(str(record.get("account", "")))

    def _toggle_selected_account_favorite(self, tree, account_type: str, refresh_tree, favorite_var: tk.BooleanVar) -> None:
        records = getattr(tree, "records", [])
        picked = tree.selection()
        if not picked or not records:
            messagebox.showinfo("提示", "请先选择一条付款人账户。")
            return
        record = records[int(picked[0])]
        account = str(record.get("account", "")).strip()
        new_value = not bool(record.get("is_favorite"))
        try:
            self.service.toggle_account_favorite(account_type, account, new_value)
        except Exception as exc:
            messagebox.showerror("设置失败", str(exc))
            return
        favorite_var.set(new_value)
        refresh_tree()
        self.status_var.set("已更新常用付款人")

    @staticmethod
    def _account_fields(account_type: str) -> dict[str, str]:
        if account_type == "payer":
            return {"name": "payer_name", "bank": "payer_bank", "account": "payer_account"}
        return {"name": "payee_name", "bank": "payee_bank", "account": "payee_account"}

    @staticmethod
    def _read_import_accounts(file_path: Path) -> list[dict[str, str]]:
        rows: list[dict[str, str]] = []
        content: str | None = None
        for encoding in ("utf-8-sig", "gbk", "utf-8"):
            try:
                content = file_path.read_text(encoding=encoding)
                break
            except Exception:
                continue
        if content is None:
            raise ValueError("无法读取导入文件，请使用 UTF-8 或 GBK 编码的 CSV 文件。")

        raw_rows = list(csv.reader(content.splitlines()))
        if not raw_rows:
            raise ValueError("导入文件没有内容。")
        header = [str(item).strip() for item in raw_rows[0]]
        if not header:
            raise ValueError("导入文件缺少表头，请提供 名称/开户行/账号 三列。")

        field_map = {
            "name": ("name", "名称", "收款人", "收款人名称"),
            "bank": ("bank", "开户行", "银行", "收款人开户行"),
            "account": ("account", "账号", "账户", "收款人账号"),
        }

        index_map: dict[str, int] = {}
        for target, aliases in field_map.items():
            for alias in aliases:
                if alias in header:
                    index_map[target] = header.index(alias)
                    break

        for line_no, item in enumerate(raw_rows[1:], start=2):
            values = [str(value).strip() for value in item]
            normalized = {
                "name": values[index_map["name"]] if "name" in index_map and len(values) > index_map["name"] else "",
                "bank": values[index_map["bank"]] if "bank" in index_map and len(values) > index_map["bank"] else "",
                "account": values[index_map["account"]] if "account" in index_map and len(values) > index_map["account"] else "",
            }
            if not any(normalized.values()) and len(values) >= 3:
                normalized = {"name": values[0], "bank": values[1], "account": values[2]}
            if not any(normalized.values()):
                continue
            if not all(normalized.values()):
                raise ValueError(f"导入文件第 {line_no} 行缺少 名称/开户行/账号。")
            rows.append(normalized)

        if not rows:
            raise ValueError("导入文件没有可用数据。")
        return rows

    @staticmethod
    def _write_import_template(file_path: Path | str) -> None:
        Path(file_path).write_text(
            "名称,开户行,账号\n示例收款人,中国银行郑州分行,6222000000000000000\n",
            encoding="utf-8-sig",
        )

    @staticmethod
    def _export_print_records_to_excel(file_path: Path, records: list[dict]) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "已打印记录"
        headers = [
            "状态",
            "操作时间",
            "操作类型",
            "模板",
            "操作员",
            "流水号",
            "完整票号",
            "唯一识别码",
            "收款人名称",
            "收款人账号",
            "收款人开户行",
            "付款人名称",
            "付款人账号",
            "付款人开户行",
            "金额",
            "日期",
            "备注",
            "电脑名称",
            "物理地址",
            "作废时间",
            "作废原因",
            "作废操作员",
        ]
        sheet.append(headers)
        for cell in sheet[1]:
            cell.font = Font(bold=True)

        for item in records:
            sheet.append(
                [
                    TicketPrinterApp._print_record_status_text(item),
                    item.get("operation_time", ""),
                    item.get("action_type", ""),
                    item.get("template", ""),
                    item.get("operator_name", ""),
                    item.get("serial_number", ""),
                    item.get("full_serial", ""),
                    item.get("security_code", ""),
                    item.get("payee_name", ""),
                    item.get("payee_account", ""),
                    item.get("payee_bank", ""),
                    item.get("payer_name", ""),
                    item.get("payer_account", ""),
                    item.get("payer_bank", ""),
                    item.get("amount", ""),
                    item.get("date", ""),
                    item.get("note", ""),
                    item.get("device_id", ""),
                    item.get("physical_address", ""),
                    item.get("voided_at", ""),
                    item.get("void_reason", ""),
                    item.get("void_operator", ""),
                ]
            )

        widths = [10, 24, 12, 18, 14, 14, 14, 16, 18, 20, 24, 18, 20, 24, 12, 14, 24, 18, 20, 20, 24, 14]
        for index, width in enumerate(widths, start=1):
            sheet.column_dimensions[chr(64 + index)].width = width

        workbook.save(file_path)

    def _refresh_table(self, select_index: int | None = None) -> None:
        for item_id in self.table.get_children():
            self.table.delete(item_id)
        for index, row in enumerate(self.rows):
            status_text = self._row_status_text(row)
            tags = ["even" if index % 2 == 0 else "odd"]
            if row.printed_count == 1:
                tags.append("status_done")
            elif row.printed_count > 1:
                tags.append("status_reprint")
            self.table.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    status_text,
                    row.payee_name,
                    row.payee_account,
                    row.payer_name,
                    f"{row.amount:.2f}",
                    row.date_text,
                    row.note,
                ),
                tags=tuple(tags),
            )
        if select_index is not None and 0 <= select_index < len(self.rows):
            self.table.selection_set(str(select_index))
            self.table.focus(str(select_index))
        self._refresh_summary_cards()

    def _selected_index(self) -> int | None:
        indices = self._selected_indices()
        if not indices:
            return None
        return indices[0]

    def _selected_indices(self) -> list[int]:
        selection = self.table.selection()
        return [int(item_id) for item_id in selection]

    def _selected_row(self) -> DataRow | None:
        index = self._selected_index()
        if index is None:
            return None
        return self.rows[index]

    def _selected_rows(self) -> list[DataRow]:
        return [self.rows[index] for index in self._selected_indices() if 0 <= index < len(self.rows)]

    def _on_table_select(self, _event=None) -> None:
        row = self._selected_row()
        if row is None:
            self._refresh_summary_cards()
            return
        values = {
            "payee_name": row.payee_name,
            "payee_account": row.payee_account,
            "payee_bank": row.payee_bank,
            "payer_name": row.payer_name,
            "payer_account": row.payer_account,
            "amount": f"{row.amount:.2f}",
            "note": row.note,
            "date_text": row.date_text,
            "payer_bank": row.payer_bank,
            "in_city": row.in_city,
            "in_county": row.in_county,
            "extra_1": row.extra_1,
            "extra_2": row.extra_2,
        }
        for key, value in values.items():
            self.field_vars[key].set(value)
        self._refresh_default_payer_flag()
        self._refresh_summary_cards()

    @staticmethod
    def _row_status_text(row: DataRow) -> str:
        if row.printed_count <= 0:
            return "未打印"
        if row.printed_count == 1:
            return "已打印"
        return f"已重打{row.printed_count - 1}次"

    @staticmethod
    def _is_print_record_voided(record: dict) -> bool:
        return str(record.get("record_status", "")).strip().lower() == "voided"

    @staticmethod
    def _print_record_status_text(record: dict) -> str:
        return "已作废" if TicketPrinterApp._is_print_record_voided(record) else "正常"

    def _refresh_summary_cards(self) -> None:
        row_count = len(self.rows)
        total_amount = sum((row.amount for row in self.rows), Decimal("0"))
        selected_rows = self._selected_rows()
        self.queue_count_var.set(f"待打印 {row_count} 笔")
        self.total_amount_var.set(f"合计 {total_amount:.2f} 元")
        if not selected_rows:
            self.selection_summary_var.set("当前未选中记录")
            return
        selected_total = sum((row.amount for row in selected_rows), Decimal("0"))
        self.selection_summary_var.set(f"已选 {len(selected_rows)} 笔 / {selected_total:.2f} 元")

    def _apply_runtime_theme(self) -> None:
        theme = self.current_theme
        if hasattr(self, "status_banner"):
            self.status_banner.configure(bg=theme["root_bg"])
        for border in getattr(self, "_entry_borders", []):
            border.configure(bg=theme["entry_border"])
        for entry in getattr(self, "_entry_widgets", []):
            entry.configure(
                bg=theme["entry_bg"],
                fg=theme["title_fg"],
                insertbackground=theme["title_fg"],
            )
        for widget_name in ("default_template_checkbutton", "default_payer_checkbutton"):
            widget = getattr(self, widget_name, None)
            if widget is not None:
                widget.configure(
                    bg=theme["panel_bg"],
                    fg=theme["text_fg"],
                    activebackground=theme["panel_bg"],
                    activeforeground=theme["title_fg"],
                    selectcolor=theme["panel_bg"],
                )
        if hasattr(self, "table"):
            self.table.tag_configure("odd", background=theme["table_odd_bg"])
            self.table.tag_configure("even", background=theme["table_even_bg"])
            self._refresh_table(self._selected_index())
        self._update_status_banner()

    def _update_status_banner(self, *_args) -> None:
        if not hasattr(self, "status_banner"):
            return
        status_text = self.status_var.get().strip()
        theme = self.current_theme
        if any(keyword in status_text for keyword in ("失败", "错误")):
            bg_color, fg_color = theme["root_bg"], theme["status_error"]
        elif any(keyword in status_text for keyword in ("完成", "已", "成功")):
            bg_color, fg_color = theme["root_bg"], theme["status_success"]
        elif "处理" in status_text:
            bg_color, fg_color = theme["root_bg"], theme["status_processing"]
        else:
            bg_color, fg_color = theme["root_bg"], theme["status_idle"]
        self.status_banner.configure(bg=bg_color, fg=fg_color)

    def _refresh_sequence_display(self) -> None:
        try:
            prefix, current = self.service.get_sequence_state()
            if prefix:
                self.sequence_display_var.set(f"当前票号：{self.service.format_serial(prefix, current)}")
            else:
                self.sequence_display_var.set("当前票号：未设置前置符")
        except Exception as exc:
            self.sequence_display_var.set("当前票号：读取失败")
            logging.exception("读取流水号失败：%s", exc)

    def _on_template_change(self, *_args) -> None:
        self.template_display_var.set(f"当前模板：{self.template_var.get()}")
        self._refresh_default_template_flag()

    def _selected_template_sheet(self) -> str:
        return self.template_label_to_sheet[self.template_var.get()]

    def _refresh_default_template_flag(self) -> None:
        try:
            default_template = self.service.get_default_template()
            current_template = self._selected_template_sheet()
            self._updating_default_template_flag = True
            self.default_template_var.set(current_template == default_template)
        finally:
            self._updating_default_template_flag = False

    def _refresh_default_payer_flag(self) -> None:
        payer = self.service.get_default_payer()
        current = {
            "name": self.field_vars["payer_name"].get().strip(),
            "bank": self.field_vars["payer_bank"].get().strip(),
            "account": self.field_vars["payer_account"].get().strip(),
        }
        try:
            self._updating_default_payer_flag = True
            self.default_payer_var.set(
                bool(current["name"] or current["bank"] or current["account"])
                and current == payer
            )
        finally:
            self._updating_default_payer_flag = False

    def _toggle_default_template(self) -> None:
        if self._updating_default_template_flag:
            return
        if not self.default_template_var.get():
            self._refresh_default_template_flag()
            return
        template_name = self._selected_template_sheet()
        self.service.set_default_template(template_name)
        self.template_display_var.set(f"当前模板：{self.template_var.get()}（默认）")
        self.status_var.set("已保存默认模板")

    def _toggle_default_payer(self) -> None:
        if self._updating_default_payer_flag:
            return
        payer_name = self.field_vars["payer_name"].get().strip()
        payer_bank = self.field_vars["payer_bank"].get().strip()
        payer_account = self.field_vars["payer_account"].get().strip()
        if self.default_payer_var.get():
            if not payer_name or not payer_bank or not payer_account:
                self.default_payer_var.set(False)
                messagebox.showinfo("提示", "请先完整填写付款人名称、账号和开户行，再设置默认付款人。")
                return
            self.service.set_default_payer(payer_name, payer_bank, payer_account)
            self.status_var.set("已保存默认付款人")
        else:
            self.service.clear_default_payer()
            self.status_var.set("已取消默认付款人")
        self._refresh_default_payer_flag()

    def _on_theme_change(self, *_args) -> None:
        if self._updating_theme_flag:
            return
        theme_label = self.theme_var.get().strip()
        theme_name = self.theme_label_to_name.get(theme_label)
        if not theme_name or theme_name == self.current_theme_name:
            return
        self.current_theme_name = theme_name
        self.current_theme = THEME_PRESETS[theme_name]
        self.service.set_theme_name(theme_name)
        self._configure_styles()
        self.status_var.set(f"已切换主题：{theme_label}")

    def _save_operator_name(self) -> None:
        operator_name = self.operator_var.get().strip()
        self.service.set_operator_name(operator_name)
        self.operator_var.set(operator_name)
        self.status_var.set("已保存操作员")

    def _select_all_rows(self, _event=None):
        items = self.table.get_children()
        if items:
            self.table.selection_set(items)
            self.table.focus(items[0])
            self._on_table_select()
        return "break"

    def _query_print_records(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("已打印记录")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("1260x460")
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)

        search_var = tk.StringVar()
        top_frame = ttk.Frame(dialog, padding=(12, 12, 12, 8))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)
        ttk.Label(top_frame, text="记录查询").grid(row=0, column=0, padx=(0, 8))
        search_outer, search_entry = self._create_text_entry(top_frame, search_var)
        search_outer.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        columns = ("status", "time", "action", "template", "operator", "serial", "security", "payee", "payer", "amount")
        tree = ttk.Treeview(dialog, columns=columns, show="headings", height=14)
        headings = {
            "status": "状态",
            "time": "操作时间",
            "action": "操作类型",
            "template": "模板",
            "operator": "操作员",
            "serial": "流水号",
            "security": "唯一识别码",
            "payee": "收款人",
            "payer": "付款人",
            "amount": "金额",
        }
        widths = {
            "status": 80,
            "time": 150,
            "action": 80,
            "template": 160,
            "operator": 100,
            "serial": 120,
            "security": 100,
            "payee": 140,
            "payer": 140,
            "amount": 90,
        }
        for column, heading in headings.items():
            tree.heading(column, text=heading)
            tree.column(column, width=widths[column], anchor="w")
        tree.tag_configure("voided", foreground="#9b3d3d")
        tree.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns", pady=(0, 12))
        tree.configure(yscrollcommand=scrollbar.set)

        def current_records() -> list[dict]:
            return self.service.find_print_records(search_var.get().strip())

        def refresh_tree() -> None:
            records = current_records()
            tree.records = records  # type: ignore[attr-defined]
            tree.delete(*tree.get_children())
            for index, item in enumerate(records):
                tags = ("voided",) if self._is_print_record_voided(item) else ()
                tree.insert(
                    "",
                    "end",
                    iid=str(index),
                    values=(
                        self._print_record_status_text(item),
                        item.get("operation_time", ""),
                        item.get("action_type", ""),
                        item.get("template", ""),
                        item.get("operator_name", ""),
                        item.get("serial_number", ""),
                        item.get("security_code", ""),
                        item.get("payee_name", ""),
                        item.get("payer_name", ""),
                        item.get("amount", ""),
                    ),
                    tags=tags,
                )

        def selected_record() -> dict | None:
            picked = tree.selection()
            records = getattr(tree, "records", [])
            if not picked or not records:
                return None
            return records[int(picked[0])]

        def void_selected_record() -> None:
            record = selected_record()
            if not record:
                messagebox.showinfo("提示", "请先选择要作废的记录。", parent=dialog)
                return
            if self._is_print_record_voided(record):
                messagebox.showinfo("提示", "该记录已经作废。", parent=dialog)
                return
            serial_number = str(record.get("serial_number", "")).strip()
            security_code = str(record.get("security_code", "")).strip()
            if not messagebox.askyesno(
                "确认作废",
                f"确定作废这条记录吗？\n\n流水号：{serial_number}\n唯一识别码：{security_code}\n\n作废后记录会保留，但会标记为已作废。",
                parent=dialog,
            ):
                return
            reason = simpledialog.askstring(
                "作废原因",
                "请输入作废原因（例如：打印偏移、卡纸、内容错误、重新开票）：",
                parent=dialog,
            )
            if reason is None:
                return
            try:
                self.service.void_print_record(record, reason)
            except Exception as exc:
                messagebox.showerror("作废失败", str(exc), parent=dialog)
                return
            refresh_tree()
            self.status_var.set(f"已作废流水号：{serial_number}")
            messagebox.showinfo("作废成功", "该记录已标记为作废。", parent=dialog)

        def export_records() -> None:
            records = current_records()
            if not records:
                messagebox.showinfo("提示", "当前没有可导出的记录。", parent=dialog)
                return
            default_name = f"已打印记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            file_path = filedialog.asksaveasfilename(
                parent=dialog,
                title="导出已打印记录",
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel 工作簿", "*.xlsx")],
            )
            if not file_path:
                return
            try:
                self._export_print_records_to_excel(Path(file_path), records)
            except Exception as exc:
                messagebox.showerror("导出失败", str(exc), parent=dialog)
                return
            self.status_var.set(f"已导出 {len(records)} 条打印记录")
            messagebox.showinfo("导出成功", f"已导出 {len(records)} 条记录。", parent=dialog)

        ttk.Button(top_frame, text="查询", command=refresh_tree).grid(row=0, column=2, padx=4)
        ttk.Button(top_frame, text="作废选中", command=void_selected_record).grid(row=0, column=3, padx=4)
        ttk.Button(top_frame, text="导出Excel", command=export_records).grid(row=0, column=4, padx=4)
        ttk.Button(top_frame, text="关闭", command=dialog.destroy).grid(row=0, column=5, padx=(12, 0))

        tip = ttk.Label(dialog, text="可按流水号、唯一识别码、付款人、收款人等关键字查询", foreground="#666666")
        tip.grid(row=2, column=0, sticky="w", padx=12, pady=(0, 12))

        search_entry.bind("<Return>", lambda _event: refresh_tree())
        refresh_tree()
        search_entry.focus_set()
        dialog.wait_window()

    def _verify_print_record_dialog(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("核验票据")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("760x500")
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)

        keyword_var = tk.StringVar()
        top_frame = ttk.Frame(dialog, padding=(12, 12, 12, 8))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)
        ttk.Label(top_frame, text="流水号/唯一识别码").grid(row=0, column=0, padx=(0, 8))
        keyword_outer, keyword_entry = self._create_text_entry(top_frame, keyword_var)
        keyword_outer.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        result_frame = ttk.LabelFrame(dialog, text="核验结果", padding=12)
        result_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        result_frame.columnconfigure(1, weight=1)

        result_vars = {
            "status": tk.StringVar(value="请输入流水号或唯一识别码后点击核验"),
            "record_status": tk.StringVar(value=""),
            "serial": tk.StringVar(value=""),
            "security": tk.StringVar(value=""),
            "template": tk.StringVar(value=""),
            "operator": tk.StringVar(value=""),
            "action": tk.StringVar(value=""),
            "time": tk.StringVar(value=""),
            "payee": tk.StringVar(value=""),
            "payer": tk.StringVar(value=""),
            "amount": tk.StringVar(value=""),
            "device": tk.StringVar(value=""),
            "mac": tk.StringVar(value=""),
            "void_reason": tk.StringVar(value=""),
        }

        fields = [
            ("核验状态", "status"),
            ("记录状态", "record_status"),
            ("流水号", "serial"),
            ("唯一识别码", "security"),
            ("模板", "template"),
            ("操作员", "operator"),
            ("操作类型", "action"),
            ("打印时间", "time"),
            ("收款人", "payee"),
            ("付款人", "payer"),
            ("金额", "amount"),
            ("电脑名称", "device"),
            ("物理地址", "mac"),
            ("作废原因", "void_reason"),
        ]
        for row_index, (label, key) in enumerate(fields):
            ttk.Label(result_frame, text=label).grid(row=row_index, column=0, sticky="ne", padx=(0, 8), pady=4)
            ttk.Label(
                result_frame,
                textvariable=result_vars[key],
                foreground="#1f5f99" if key == "status" else None,
                wraplength=520,
                justify="left",
            ).grid(row=row_index, column=1, sticky="nw", pady=4)

        def do_verify(_event=None) -> None:
            keyword = keyword_var.get().strip()
            record = self.service.verify_print_record(keyword)
            if not record:
                result_vars["status"].set("未找到匹配记录")
                for key in ("record_status", "serial", "security", "template", "operator", "action", "time", "payee", "payer", "amount", "device", "mac", "void_reason"):
                    result_vars[key].set("")
                return
            if self._is_print_record_voided(record):
                result_vars["status"].set("已找到票据记录（该记录已作废）")
            else:
                result_vars["status"].set("已找到票据记录")
            result_vars["record_status"].set(self._print_record_status_text(record))
            result_vars["serial"].set(str(record.get("serial_number", "")))
            result_vars["security"].set(str(record.get("security_code", "")))
            result_vars["template"].set(str(record.get("template", "")))
            result_vars["operator"].set(str(record.get("operator_name", "")))
            result_vars["action"].set(str(record.get("action_type", "")))
            result_vars["time"].set(str(record.get("operation_time", "")))
            result_vars["payee"].set(str(record.get("payee_name", "")))
            result_vars["payer"].set(str(record.get("payer_name", "")))
            result_vars["amount"].set(str(record.get("amount", "")))
            result_vars["device"].set(str(record.get("device_id", "")))
            result_vars["mac"].set(str(record.get("physical_address", "")))
            result_vars["void_reason"].set(str(record.get("void_reason", "")))

        ttk.Button(top_frame, text="核验", command=do_verify).grid(row=0, column=2, padx=4)
        ttk.Button(top_frame, text="关闭", command=dialog.destroy).grid(row=0, column=3, padx=(12, 0))

        tip = ttk.Label(dialog, text="可输入流水号、完整票号或唯一识别码进行核验", foreground="#666666")
        tip.grid(row=2, column=0, sticky="w", padx=12, pady=(0, 12))

        keyword_entry.bind("<Return>", do_verify)
        keyword_entry.focus_set()
        dialog.wait_window()

    def _run_action(self, success_message: str, action) -> None:
        try:
            logging.info("开始执行：%s", success_message)
            self.status_var.set("处理中...")
            self.root.update_idletasks()
            action()
            if self.status_var.get() == "处理中...":
                self.status_var.set(success_message)
            logging.info("执行完成：%s", success_message)
        except Exception as exc:
            self.status_var.set("执行失败")
            logging.exception("执行失败：%s", success_message)
            messagebox.showerror("执行失败", str(exc))


def main() -> None:
    workbook_path = app_dir() / DEFAULT_WORKBOOK_NAME
    if not workbook_path.exists():
        raise FileNotFoundError(f"当前目录未找到模板文件：{DEFAULT_WORKBOOK_NAME}")

    root = tk.Tk()
    TicketPrinterApp(root, workbook_path)
    root.mainloop()


if __name__ == "__main__":
    main()
