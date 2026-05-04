#!/usr/bin/env python3
"""
First-run / reconfiguration wizard for AracTakip and TeklifTakip executables.

Each wizard renders a Tkinter form prefilled from the existing config (if any),
then writes a clean config file back. The Araç Takip wizard preserves CALC/SET
rule blocks unchanged so that domain rules survive UI edits.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


ARAC_DEFAULT_RULES = """\
# EMAIL SENDING RULES
# PLAKANO ETIKETHUCRESI DEGERHUCRESI KURAL
CALC: D8 B13 E13 <=15
CALC: D8 B14 E14 <=15
CALC: D8 B15 E15 <=15
CALC: D8 B16 E16 <=15
CALC: D8 B17 E17 <=15
CALC: D8 B22 E22 <=500

CALC: I8 G13 J13 <=15
CALC: I8 G14 J14 <=15
CALC: I8 G15 J15 <=15
CALC: I8 G16 J16 <=15
CALC: I8 G17 J17 <=15
CALC: I8 G22 J22 <=500

CALC: N8 L13 O13 <=15
CALC: N8 L14 O14 <=15
CALC: N8 L15 O15 <=15
CALC: N8 L16 O16 <=15
CALC: N8 L17 O17 <=15
CALC: N8 L22 O22 <=500

CALC: D31 B36 E36 <=15
CALC: D31 B37 E37 <=15
CALC: D31 B38 E38 <=15
CALC: D31 B39 E39 <=15
CALC: D31 B40 E40 <=15
CALC: D31 B45 E45 <=500

CALC: I31 G36 J36 <=15
CALC: I31 G37 J37 <=15
CALC: I31 G38 J38 <=15
CALC: I31 G39 J39 <=15
CALC: I31 G40 J40 <=15
CALC: I31 G45 J45 <=500

CALC: N31 L36 O36 <=15
CALC: N31 L37 O37 <=15
CALC: N31 L38 O38 <=15
CALC: N31 L39 O39 <=15
CALC: N31 L40 O40 <=15
CALC: N31 L45 O45 <=500

CALC: D55 B60 E60 <=15
CALC: D55 B61 E61 <=15
CALC: D55 B62 E62 <=15
CALC: D55 B63 E63 <=15
CALC: D55 B64 E64 <=15
CALC: D55 B69 E69 <=500

CALC: I55 G60 J60 <=15
CALC: I55 G61 J61 <=15
CALC: I55 G62 J62 <=15
CALC: I55 G63 J63 <=15
CALC: I55 G64 J64 <=15
CALC: I55 G69 J69 <=500

CALC: N55 L60 O60 <=15
CALC: N55 L61 O61 <=15
CALC: N55 L62 O62 <=15
CALC: N55 L63 O63 <=15
CALC: N55 L64 O64 <=15
CALC: N55 L69 O69 <=500

CALC: D79 B84 E84 <=15
CALC: D79 B85 E85 <=15
CALC: D79 B86 E86 <=15
CALC: D79 B87 E87 <=15
CALC: D79 B88 E88 <=15
CALC: D79 B93 E93 <=500

CALC: I79 G84 J84 <=15
CALC: I79 G85 J85 <=15
CALC: I79 G86 J86 <=15
CALC: I79 G87 J87 <=15
CALC: I79 G88 J88 <=15
CALC: I79 G93 J93 <=500

CALC: N79 L84 O84 <=15
CALC: N79 L85 O85 <=15
CALC: N79 L86 O86 <=15
CALC: N79 L87 O87 <=15
CALC: N79 L88 O88 <=15
CALC: N79 L93 O93 <=500

# KM INCREASE CELLS
SET: C22 50
SET: H22 300
SET: M22 75
SET: C45 75
SET: H45 60
SET: M45 60
SET: C69 60
SET: H69 50
SET: M69 50
SET: C93 50
SET: H93 60
SET: M93 60
"""


def _read_existing(path: Path) -> Tuple[Dict[str, str], List[str]]:
    """Return (key/value dict for simple keys, raw lines list) from an existing config."""
    if not path.exists():
        return {}, []
    encodings = ("utf-8-sig", "utf-8", "cp1254", "latin-1")
    for encoding in encodings:
        try:
            text = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        return {}, []

    raw_lines = text.splitlines()
    values: Dict[str, str] = {}
    for raw in raw_lines:
        line = raw.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        key, value = line.split(":", 1)
        values[key.strip().upper()] = value.strip()
    return values, raw_lines


def _extract_rule_block(raw_lines: List[str]) -> str:
    """Return the CALC/SET rule lines and their preceding comments, untouched."""
    block: List[str] = []
    for raw in raw_lines:
        stripped = raw.strip()
        if stripped.startswith("CALC:") or stripped.startswith("SET:"):
            block.append(raw)
    if not block:
        return ""
    return "\n".join(block) + "\n"


class _BaseWizard:
    title = "Setup"
    width = 560
    height = 520

    def __init__(self, config_path: Path) -> None:
        self.config_path = config_path
        self.saved = False
        self.root = tk.Tk()
        self.root.title(self.title)
        self.root.geometry(f"{self.width}x{self.height}")
        self.root.minsize(self.width, self.height)
        self.root.configure(padx=18, pady=18)
        self.entries: Dict[str, tk.Variable] = {}
        self.text_widgets: Dict[str, tk.Text] = {}

    def _add_label(self, parent: tk.Widget, text: str, row: int) -> None:
        ttk.Label(parent, text=text).grid(row=row, column=0, sticky="w", pady=(8, 2))

    def _add_entry(
        self,
        parent: tk.Widget,
        key: str,
        row: int,
        default: str = "",
        show: Optional[str] = None,
    ) -> None:
        var = tk.StringVar(value=default)
        entry = ttk.Entry(parent, textvariable=var, width=60, show=show or "")
        entry.grid(row=row, column=0, columnspan=3, sticky="we")
        self.entries[key] = var

    def _add_browse(
        self, parent: tk.Widget, key: str, row: int, default: str = ""
    ) -> None:
        var = tk.StringVar(value=default)
        entry = ttk.Entry(parent, textvariable=var, width=50)
        entry.grid(row=row, column=0, columnspan=2, sticky="we")
        ttk.Button(
            parent,
            text="Sec...",
            command=lambda: self._browse_into(var),
        ).grid(row=row, column=2, sticky="e", padx=(8, 0))
        self.entries[key] = var

    def _browse_into(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            title="Excel dosyasini sec",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("All", "*.*")],
        )
        if path:
            var.set(Path(path).name)

    def _add_text(self, parent: tk.Widget, key: str, row: int, default: str, height: int) -> None:
        widget = tk.Text(parent, height=height, width=60, wrap="word")
        widget.grid(row=row, column=0, columnspan=3, sticky="we")
        widget.insert("1.0", default)
        self.text_widgets[key] = widget

    def _read_text(self, key: str) -> str:
        return self.text_widgets[key].get("1.0", "end").strip()

    def _grid_columns(self, parent: tk.Widget) -> None:
        parent.columnconfigure(0, weight=1)

    def _add_buttons(self, parent: tk.Widget, row: int) -> None:
        bar = ttk.Frame(parent)
        bar.grid(row=row, column=0, columnspan=3, sticky="e", pady=(18, 0))
        ttk.Button(bar, text="Iptal", command=self._cancel).pack(side="right", padx=(8, 0))
        ttk.Button(bar, text="Kaydet", command=self._save).pack(side="right")

    def _cancel(self) -> None:
        self.saved = False
        self.root.destroy()

    def _save(self) -> None:
        try:
            content = self._build_config()
        except ValueError as exc:
            messagebox.showerror("Eksik bilgi", str(exc))
            return
        try:
            self.config_path.write_text(content, encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Kaydetme hatasi", f"{self.config_path}\n\n{exc}")
            return
        messagebox.showinfo("Tamam", f"Ayarlar kaydedildi:\n{self.config_path}")
        self.saved = True
        self.root.destroy()

    def _build_config(self) -> str:
        raise NotImplementedError

    def run(self) -> bool:
        self.root.mainloop()
        return self.saved


class _AracTakipWizard(_BaseWizard):
    title = "Arac Takip Ayarlar Sihirbazi"
    width = 600
    height = 560

    def __init__(self, config_path: Path) -> None:
        super().__init__(config_path)
        existing, raw_lines = _read_existing(config_path)
        self._existing_rules = _extract_rule_block(raw_lines) or ARAC_DEFAULT_RULES

        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True)
        self._grid_columns(frame)

        row = 0
        ttk.Label(
            frame,
            text="Arac Takip Email Bildirim Ayarlari",
            font=("TkDefaultFont", 12, "bold"),
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 12))
        row += 1

        self._add_label(frame, "Excel dosya adi (ayni klasorde):", row); row += 1
        self._add_browse(frame, "FILENAME", row, existing.get("FILENAME", "Araç Bakım Takibi.xlsx")); row += 1

        self._add_label(frame, "Gmail kullanici adi:", row); row += 1
        self._add_entry(frame, "GMAILUSERNAME", row, existing.get("GMAILUSERNAME", "")); row += 1

        self._add_label(frame, "Gmail uygulama sifresi (App Password):", row); row += 1
        self._add_entry(frame, "GMAILPASSWORD", row, existing.get("GMAILPASSWORD", ""), show="*"); row += 1

        self._add_label(frame, "Alici e-postalar (virgul ile ayrilmis):", row); row += 1
        self._add_text(
            frame,
            "EMAILRECIPIENTS",
            row,
            existing.get("EMAILRECIPIENTS", ""),
            height=4,
        )
        row += 1

        ttk.Label(
            frame,
            text="Not: CALC ve SET kurallari korunacak (ayni klasordeki TXT'den).",
            foreground="#555",
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(12, 0))
        row += 1

        self._add_buttons(frame, row)

    def _build_config(self) -> str:
        filename = self.entries["FILENAME"].get().strip()
        username = self.entries["GMAILUSERNAME"].get().strip()
        password = self.entries["GMAILPASSWORD"].get().strip()
        recipients = self._read_text("EMAILRECIPIENTS").replace("\n", ", ")

        missing = [
            name
            for name, value in (
                ("Excel dosyasi", filename),
                ("Gmail kullanici adi", username),
                ("Gmail sifresi", password),
                ("Alici e-postalar", recipients),
            )
            if not value
        ]
        if missing:
            raise ValueError("Asagidaki alanlar bos olamaz:\n - " + "\n - ".join(missing))

        return (
            "# Arac Bakim Takip ayarlari (sihirbazdan kaydedildi)\n"
            f"FILENAME: {filename}\n\n"
            "# Email ayarlari\n"
            f"GMAILUSERNAME: {username}\n"
            f"GMAILPASSWORD: {password}\n"
            f"EMAILRECIPIENTS: {recipients}\n\n"
            f"{self._existing_rules}"
        )


class _TeklifTakipWizard(_BaseWizard):
    title = "Teklif Takip Ayarlar Sihirbazi"
    width = 620
    height = 720

    def __init__(self, config_path: Path) -> None:
        super().__init__(config_path)
        existing, _ = _read_existing(config_path)

        canvas = tk.Canvas(self.root, borderwidth=0, highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        self._grid_columns(frame)

        row = 0
        ttk.Label(
            frame,
            text="Teklif Takip Hatirlatma Ayarlari",
            font=("TkDefaultFont", 12, "bold"),
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 12))
        row += 1

        self._add_label(frame, "Excel dosya adi (ayni klasorde):", row); row += 1
        self._add_browse(frame, "FILENAME", row, existing.get("FILENAME", "TeklifTakip_Master.xlsx")); row += 1

        self._add_label(frame, "Sayfa adi (bos = ilk sayfa):", row); row += 1
        self._add_entry(frame, "SHEETNAME", row, existing.get("SHEETNAME", "")); row += 1

        self._add_label(frame, "SMTP host:", row); row += 1
        self._add_entry(frame, "SMTP_HOST", row, existing.get("SMTP_HOST", "smtp.gmail.com")); row += 1

        self._add_label(frame, "SMTP port:", row); row += 1
        self._add_entry(frame, "SMTP_PORT", row, existing.get("SMTP_PORT", "465")); row += 1

        self._add_label(frame, "SMTP kullanici (Gmail):", row); row += 1
        self._add_entry(frame, "SMTP_USERNAME", row, existing.get("SMTP_USERNAME", "")); row += 1

        self._add_label(frame, "SMTP sifresi (App Password):", row); row += 1
        self._add_entry(frame, "SMTP_PASSWORD", row, existing.get("SMTP_PASSWORD", ""), show="*"); row += 1

        self._add_label(frame, "Ortak mail alicilar (gun >= 2):", row); row += 1
        self._add_text(frame, "COMMON_EMAIL", row, existing.get("COMMON_EMAIL", ""), height=3); row += 1

        self._add_label(frame, "Patron mail alicilar (gun >= 4):", row); row += 1
        self._add_text(frame, "BOSS_EMAIL", row, existing.get("BOSS_EMAIL", ""), height=3); row += 1

        self._add_label(frame, "CC alicilar (opsiyonel):", row); row += 1
        self._add_text(frame, "CC_EMAILS", row, existing.get("CC_EMAILS", ""), height=2); row += 1

        self._add_label(frame, "Ortak hatirlatma esigi (gun):", row); row += 1
        self._add_entry(frame, "REMINDER_COMMON_DAYS", row, existing.get("REMINDER_COMMON_DAYS", "2")); row += 1

        self._add_label(frame, "Patron hatirlatma esigi (gun):", row); row += 1
        self._add_entry(frame, "REMINDER_BOSS_DAYS", row, existing.get("REMINDER_BOSS_DAYS", "4")); row += 1

        self._add_label(frame, "Konu basligi oneki:", row); row += 1
        self._add_entry(frame, "SUBJECT_PREFIX", row, existing.get("SUBJECT_PREFIX", "[Teklif Takip]")); row += 1

        self._existing_columns = {
            "DATE_COLUMN": existing.get("DATE_COLUMN", "TARIH"),
            "STATUS_COLUMN": existing.get("STATUS_COLUMN", "SON DURUMU"),
            "COMPANY_COLUMN": existing.get("COMPANY_COLUMN", "SIRKET ADI"),
            "CONTACT_COLUMN": existing.get("CONTACT_COLUMN", "ILGILI"),
            "FLAG_COMMON_COLUMN": existing.get("FLAG_COMMON_COLUMN", "MAIL1"),
            "FLAG_COMMON_AT_COLUMN": existing.get("FLAG_COMMON_AT_COLUMN", ""),
            "FLAG_BOSS_COLUMN": existing.get("FLAG_BOSS_COLUMN", "MAIL2"),
            "FLAG_BOSS_AT_COLUMN": existing.get("FLAG_BOSS_AT_COLUMN", ""),
        }

        ttk.Label(
            frame,
            text="Not: Excel sutun isimleri varsayilan olarak korunur (TARIH, SIRKET ADI vs.).",
            foreground="#555",
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(12, 0))
        row += 1

        self._add_buttons(frame, row)

    def _build_config(self) -> str:
        def get(key: str) -> str:
            return self.entries[key].get().strip()

        def get_text(key: str) -> str:
            return ", ".join(part.strip() for part in re.split(r"[\n,;]", self._read_text(key)) if part.strip())

        filename = get("FILENAME")
        username = get("SMTP_USERNAME")
        password = get("SMTP_PASSWORD")
        common = get_text("COMMON_EMAIL")
        boss = get_text("BOSS_EMAIL")

        missing = [
            label
            for label, value in (
                ("Excel dosyasi", filename),
                ("SMTP kullanici", username),
                ("SMTP sifresi", password),
                ("Ortak mail alicilar", common),
                ("Patron mail alicilar", boss),
            )
            if not value
        ]
        if missing:
            raise ValueError("Asagidaki alanlar bos olamaz:\n - " + "\n - ".join(missing))

        try:
            common_days = int(get("REMINDER_COMMON_DAYS"))
            boss_days = int(get("REMINDER_BOSS_DAYS"))
        except ValueError:
            raise ValueError("Hatirlatma esikleri tam sayi olmali.")
        if boss_days < common_days:
            raise ValueError("Patron esigi ortak esigi geride birakamaz (boss_days >= common_days olmali).")

        cols = self._existing_columns

        lines = [
            "# Teklif takip ayarlari (sihirbazdan kaydedildi)",
            f"FILENAME: {filename}",
            f"SHEETNAME: {get('SHEETNAME')}",
            "",
            "# SMTP",
            f"SMTP_HOST: {get('SMTP_HOST') or 'smtp.gmail.com'}",
            f"SMTP_PORT: {get('SMTP_PORT') or '465'}",
            f"SMTP_USERNAME: {username}",
            f"SMTP_PASSWORD: {password}",
            "",
            "# Alicilar",
            f"COMMON_EMAIL: {common}",
            f"BOSS_EMAIL: {boss}",
            f"CC_EMAILS: {get_text('CC_EMAILS')}",
            "",
            "# Esikler",
            f"REMINDER_COMMON_DAYS: {common_days}",
            f"REMINDER_BOSS_DAYS: {boss_days}",
            "",
            "# Sutun adlari",
            f"DATE_COLUMN: {cols['DATE_COLUMN']}",
            f"STATUS_COLUMN: {cols['STATUS_COLUMN']}",
            f"COMPANY_COLUMN: {cols['COMPANY_COLUMN']}",
            f"CONTACT_COLUMN: {cols['CONTACT_COLUMN']}",
            f"FLAG_COMMON_COLUMN: {cols['FLAG_COMMON_COLUMN']}",
            f"FLAG_COMMON_AT_COLUMN: {cols['FLAG_COMMON_AT_COLUMN']}",
            f"FLAG_BOSS_COLUMN: {cols['FLAG_BOSS_COLUMN']}",
            f"FLAG_BOSS_AT_COLUMN: {cols['FLAG_BOSS_AT_COLUMN']}",
            "",
            f"SUBJECT_PREFIX: {get('SUBJECT_PREFIX') or '[Teklif Takip]'}",
            "",
        ]
        return "\n".join(lines)


def run_arac_takip_wizard(config_path: Path) -> bool:
    return _AracTakipWizard(config_path).run()


def run_teklif_takip_wizard(config_path: Path) -> bool:
    return _TeklifTakipWizard(config_path).run()
