"""
DRE Fechamento Financeiro - Importador de Lançamentos
Versão 1.0
"""

import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
import re as _re
from engine import DREEngine, detect_source_type, DREImportError, CREDENTIALS_FILE

CONFIG_FILE = Path(__file__).parent / "config.json"
DOTENV_FILE = Path(__file__).parent / ".env"


def _load_config() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_config(data: dict):
    CONFIG_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _read_dotenv() -> dict:
    """Parse .env without any external package."""
    result = {}
    try:
        for line in DOTENV_FILE.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            result[key.strip()] = val.strip().strip('"').strip("'")
    except Exception:
        pass
    return result

# ─── Color palette ───────────────────────────────────────────────
BG         = "#F0F2F5"
CARD       = "#FFFFFF"
PRIMARY    = "#2563EB"
PRIMARY_DK = "#1D4ED8"
SUCCESS    = "#16A34A"
DANGER     = "#DC2626"
WARNING    = "#D97706"
TEXT       = "#111827"
SUBTEXT    = "#6B7280"
BORDER     = "#E5E7EB"
ROW_ALT    = "#F9FAFB"
HEADER_BG  = "#EFF6FF"


class FileRow(tk.Frame):
    """One row representing a source file to import."""

    def __init__(self, parent, filepath: Path, on_remove, **kwargs):
        super().__init__(parent, bg=CARD, **kwargs)
        self.filepath = filepath
        self.on_remove = on_remove

        self.source_type = detect_source_type(filepath)
        self.sheet_var = tk.StringVar(value="DESPESAS")

        # ── layout ──
        self.columnconfigure(1, weight=1)

        # icon + filename
        icon = "📄"
        tk.Label(self, text=icon, bg=CARD, font=("Segoe UI", 13)).grid(
            row=0, column=0, padx=(10, 6), pady=8
        )

        name_frame = tk.Frame(self, bg=CARD)
        name_frame.grid(row=0, column=1, sticky="ew", pady=8)

        tk.Label(
            name_frame,
            text=filepath.name,
            bg=CARD,
            fg=TEXT,
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).pack(fill="x")

        badge_color = {
            "cartao":   "#DBEAFE",
            "caixa":    "#DCFCE7",
            "bradesco": "#FEF3C7",
            "ai":       "#F3E8FF",
        }.get(self.source_type, "#F3E8FF")

        badge_text_color = {
            "cartao":   "#1D4ED8",
            "caixa":    "#15803D",
            "bradesco": "#92400E",
            "ai":       "#7C3AED",
        }.get(self.source_type, "#7C3AED")

        badge_label = {
            "cartao":   "💳 Cartão",
            "caixa":    "🏦 Caixa PJ",
            "bradesco": "🏛 Bradesco",
            "ai":       "🤖 Análise IA",
        }.get(self.source_type, "🤖 Análise IA")

        badge = tk.Label(
            name_frame,
            text=badge_label,
            bg=badge_color,
            fg=badge_text_color,
            font=("Segoe UI", 8, "bold"),
            padx=6,
            pady=1,
        )
        badge.pack(anchor="w")

        # target sheet selector — cartao/caixa auto-route by sign; bradesco by Tipo; ai auto-detects
        selector_frame = tk.Frame(self, bg=CARD)
        selector_frame.grid(row=0, column=2, padx=10, pady=8)

        if self.source_type in ("cartao", "caixa"):
            tk.Label(
                selector_frame,
                text="Auto (sinal do valor)",
                bg="#F0FDF4", fg=SUCCESS,
                font=("Segoe UI", 8, "bold"),
                padx=6, pady=3,
            ).pack()
        elif self.source_type == "ai":
            tk.Label(
                selector_frame,
                text="Auto (IA detecta)",
                bg="#F3E8FF", fg="#7C3AED",
                font=("Segoe UI", 8, "bold"),
                padx=6, pady=3,
            ).pack()
        elif self.source_type == "bradesco":
            tk.Label(selector_frame, text="Destino:", bg=CARD, fg=SUBTEXT,
                     font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
            combo = ttk.Combobox(
                selector_frame,
                textvariable=self.sheet_var,
                values=["Auto (Tipo)", "DESPESAS", "RECEITAS"],
                state="readonly",
                width=14,
                font=("Segoe UI", 9),
            )
            combo.pack(side="left")
            self.sheet_var.set("Auto (Tipo)")
        else:
            tk.Label(selector_frame, text="Destino:", bg=CARD, fg=SUBTEXT,
                     font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
            combo = ttk.Combobox(
                selector_frame,
                textvariable=self.sheet_var,
                values=["DESPESAS", "RECEITAS"],
                state="readonly",
                width=14,
                font=("Segoe UI", 9),
            )
            combo.pack(side="left")

        # remove button
        tk.Button(
            self,
            text="✕",
            command=lambda: on_remove(self),
            bg=CARD,
            fg=DANGER,
            bd=0,
            font=("Segoe UI", 12),
            cursor="hand2",
            activebackground=CARD,
            activeforeground=DANGER,
        ).grid(row=0, column=3, padx=10)

        # bottom divider
        ttk.Separator(self, orient="horizontal").grid(
            row=1, column=0, columnspan=4, sticky="ew"
        )


class PreviewTable(tk.Frame):
    """Scrollable table showing rows that will be imported."""

    COLS = ("Destino", "Data de Emissão", "Histórico da Despesa/Receita", "Classificação", "Valor (R$)", "Arquivo")

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=CARD, **kwargs)

        # header
        hdr = tk.Frame(self, bg=HEADER_BG)
        hdr.pack(fill="x")
        col_weights = [1, 1, 4, 1, 1, 2]
        col_minsizes = [70, 90, 200, 100, 110, 120]
        for i, (col, w, ms) in enumerate(zip(self.COLS, col_weights, col_minsizes)):
            tk.Label(
                hdr,
                text=col,
                bg=HEADER_BG,
                fg=TEXT,
                font=("Segoe UI", 9, "bold"),
                anchor="w",
                padx=8,
                pady=6,
            ).grid(row=0, column=i, sticky="ew")
            hdr.columnconfigure(i, weight=w, minsize=ms)

        # scrollable body
        container = tk.Frame(self, bg=CARD)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=CARD, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.body = tk.Frame(canvas, bg=CARD)

        self.body.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        canvas.create_window((0, 0), window=self.body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Make body stretch to canvas width
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfig(canvas.find_withtag("all")[0], width=e.width) if canvas.find_withtag("all") else None,
        )

        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))
        self.body.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

    def load(self, rows: list[dict], categories: dict = None):
        self.combos = []
        for widget in self.body.winfo_children():
            widget.destroy()

        cat_dict = categories or {}

        for i, row in enumerate(rows):
            bg = ROW_ALT if i % 2 == 0 else CARD
            sheet = row["sheet"]
            sheet_color = SUCCESS if sheet == "RECEITAS" else "#2563EB"

            tk.Label(
                self.body, text=sheet, bg=bg, fg=sheet_color,
                font=("Segoe UI", 8, "bold"), anchor="w", padx=8, pady=5,
            ).grid(row=i, column=0, sticky="ew")

            tk.Label(
                self.body, text=str(row["data"]), bg=bg, fg=TEXT,
                font=("Segoe UI", 9), anchor="w", padx=8,
            ).grid(row=i, column=1, sticky="ew")

            tk.Label(
                self.body, text=row["historico"], bg=bg, fg=TEXT,
                font=("Segoe UI", 9), anchor="w", padx=8,
                wraplength=250,
            ).grid(row=i, column=2, sticky="ew")

            # Save original category to track manual changes
            if "_original_cat" not in row:
                row["_original_cat"] = row.get("classificacao", "")

            classificacao = row.get("classificacao", "")
            
            # Use Combobox instead of Label
            sheet_cats = cat_dict.get(sheet, [])
            combo = ttk.Combobox(
                self.body,
                values=[""] + sheet_cats,
                font=("Segoe UI", 8),
                width=15,
            )
            combo.set(classificacao)
            combo.grid(row=i, column=3, sticky="ew", padx=8, pady=2)
            
            self.combos.append((row, combo))

            val = row["valor"]
            val_color = DANGER if isinstance(val, (int, float)) and val < 0 else SUCCESS
            tk.Label(
                self.body,
                text=f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(val, (int, float)) else str(val),
                bg=bg, fg=val_color,
                font=("Segoe UI", 9, "bold"), anchor="e", padx=8,
            ).grid(row=i, column=4, sticky="ew")

            source = row.get("source", "")
            if len(source) > 22:
                source = source[:19] + "..."
            tk.Label(
                self.body, text=source, bg=bg, fg=SUBTEXT,
                font=("Segoe UI", 8), anchor="w", padx=8,
            ).grid(row=i, column=5, sticky="ew")

        self.body.columnconfigure(0, weight=1, minsize=70)
        self.body.columnconfigure(1, weight=1, minsize=90)
        self.body.columnconfigure(2, weight=4, minsize=200)
        self.body.columnconfigure(3, weight=1, minsize=100)
        self.body.columnconfigure(4, weight=1, minsize=110)
        self.body.columnconfigure(5, weight=2, minsize=120)

    def commit_edits(self):
        """Update row dicts with the latest combobox values."""
        for row, combo in getattr(self, "combos", []):
            row["classificacao"] = combo.get().strip()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DRE — Importador de Lançamentos")
        self.geometry("860x700")
        self.minsize(760, 560)
        self.configure(bg=BG)
        self.resizable(True, True)

        self.file_rows: list[FileRow] = []
        self.preview_data: list[dict] = []

        cfg = _load_config()
        dotenv = _read_dotenv()
        self._api_key: str = dotenv.get("OPENAI_API", "")
        self._dre_id: str = cfg.get("dre_spreadsheet_id", "")

        # Target email for Google Sheets sharing
        self._target_email = ""
        try:
            if CREDENTIALS_FILE.exists():
                creds = json.loads(CREDENTIALS_FILE.read_text(encoding="utf-8"))
                self._target_email = creds.get("client_email", "")
        except Exception:
            pass

        self._build_ui()

    # ─────────────────────────────── UI BUILD ────────────────────────────────

    def _build_ui(self):
        # ── Top bar ──
        topbar = tk.Frame(self, bg=PRIMARY, height=54)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)
        tk.Label(
            topbar,
            text="  📊  DRE — Importador de Lançamentos",
            bg=PRIMARY, fg="white",
            font=("Segoe UI", 13, "bold"),
        ).pack(side="left", padx=12, pady=10)

        # ── Main body (two columns) ──
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=12)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(1, weight=1)

        # ── LEFT PANEL ──
        left = tk.Frame(body, bg=BG)
        left.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 8))
        left.rowconfigure(2, weight=1)

        # ── DRE target (Google Sheets — editable) ──
        tk.Label(left, text="1. Planilha DRE (destino)", bg=BG, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        dre_card = tk.Frame(left, bg=CARD, relief="flat", bd=1,
                            highlightbackground=BORDER, highlightthickness=1)
        dre_card.pack(fill="x")

        tk.Label(
            dre_card,
            text="Cole o link da planilha Google Sheets:",
            bg=CARD, fg=SUBTEXT,
            font=("Segoe UI", 8),
            anchor="w", padx=10,
        ).pack(fill="x", pady=(8, 2))

        dre_input_frame = tk.Frame(dre_card, bg=CARD)
        dre_input_frame.pack(fill="x", padx=8, pady=(0, 8))

        self._dre_url_var = tk.StringVar(value=self._dre_id)
        dre_entry = tk.Entry(
            dre_input_frame,
            textvariable=self._dre_url_var,
            font=("Segoe UI", 9),
            bg=CARD, fg=TEXT,
            relief="flat",
        )
        dre_entry.pack(side="left", fill="x", expand=True, padx=(2, 4), pady=4)

        tk.Button(
            dre_input_frame, text="Salvar",
            command=self._save_dre_target,
            bg=PRIMARY, fg="white", bd=0,
            font=("Segoe UI", 8, "bold"),
            padx=8, pady=4, cursor="hand2",
            activebackground=PRIMARY_DK, activeforeground="white",
        ).pack(side="right", pady=4)

        # Source files
        tk.Label(left, text="2. Arquivos de lançamentos", bg=BG, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(14, 4))

        files_card = tk.Frame(left, bg=CARD, relief="flat", bd=1,
                              highlightbackground=BORDER, highlightthickness=1)
        files_card.pack(fill="both", expand=True)

        tk.Button(
            files_card, text="➕  Adicionar arquivos",
            command=self._add_files,
            bg="#F0FDF4", fg=SUCCESS, bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=10, pady=7, cursor="hand2",
            activebackground="#DCFCE7", activeforeground=SUCCESS,
        ).pack(fill="x", padx=8, pady=8)

        # scrollable file list
        self.file_list_frame = tk.Frame(files_card, bg=CARD)
        self.file_list_frame.pack(fill="both", expand=True, padx=4)

        self._empty_label = tk.Label(
            self.file_list_frame,
            text="Nenhum arquivo adicionado",
            bg=CARD, fg=SUBTEXT, font=("Segoe UI", 9, "italic"),
        )
        self._empty_label.pack(pady=16)

        # ── RIGHT PANEL ──
        right = tk.Frame(body, bg=BG)
        right.grid(row=0, column=1, rowspan=2, sticky="nsew")
        right.rowconfigure(1, weight=1)

        preview_header = tk.Frame(right, bg=BG)
        preview_header.pack(fill="x", pady=(0, 6))

        tk.Label(preview_header, text="3. Pré-visualização", bg=BG, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        self.count_label = tk.Label(
            preview_header, text="", bg=BG, fg=SUBTEXT,
            font=("Segoe UI", 9),
        )
        self.count_label.pack(side="left", padx=8)

        tk.Button(
            preview_header, text="🔄  Atualizar prévia",
            command=self._refresh_preview,
            bg=BG, fg=PRIMARY, bd=0,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            activebackground=BG, activeforeground=PRIMARY_DK,
        ).pack(side="right")

        preview_card = tk.Frame(right, bg=CARD, relief="flat", bd=1,
                                highlightbackground=BORDER, highlightthickness=1)
        preview_card.pack(fill="both", expand=True)

        self.preview_table = PreviewTable(preview_card)
        self.preview_table.pack(fill="both", expand=True)

        # ── Bottom action bar ──
        actionbar = tk.Frame(self, bg=CARD, height=60,
                             relief="flat", bd=1,
                             highlightbackground=BORDER, highlightthickness=1)
        actionbar.pack(fill="x", side="bottom")
        actionbar.pack_propagate(False)

        self.status_label = tk.Label(
            actionbar, text="Pronto.", bg=CARD, fg=SUBTEXT,
            font=("Segoe UI", 9), anchor="w",
        )
        self.status_label.pack(side="left", padx=16)

        self.import_btn = tk.Button(
            actionbar, text="⬆  Importar para o DRE",
            command=self._start_import,
            bg=SUCCESS, fg="white", bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=20, pady=10, cursor="hand2",
            activebackground="#15803D", activeforeground="white",
        )
        self.import_btn.pack(side="right", padx=16, pady=8)

    # ─────────────────────────────── ACTIONS ─────────────────────────────────

    def _save_dre_target(self):
        raw = self._dre_url_var.get().strip()

        # Check if user pasted an xlsx/xlsm file path
        if raw.lower().endswith((".xlsx", ".xlsm", ".xls")):
            messagebox.showwarning(
                "Formato inválido",
                "O destino deve ser uma planilha Google Sheets, não um arquivo Excel local.\n\n"
                "Faça upload do arquivo para o Google Drive e cole o link aqui.",
            )
            return

        # Extract spreadsheet ID from a Google Sheets URL
        m = _re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", raw)
        dre_id = m.group(1) if m else raw  # allow pasting raw ID too

        if not dre_id:
            self._set_status("Informe o link ou ID da planilha DRE.", color=DANGER)
            return

        self._dre_id = dre_id
        self._dre_url_var.set(dre_id)
        cfg = _load_config()
        cfg["dre_spreadsheet_id"] = dre_id
        _save_config(cfg)
        self._set_status("Planilha DRE salva.")

        if self._target_email:
            self._show_sharing_prompt()

    def _show_sharing_prompt(self):
        """Custom popup to remind user to share the spreadsheet."""
        win = tk.Toplevel(self)
        win.title("Compartilhar Planilha")
        win.geometry("460x280")
        win.resizable(False, False)
        win.configure(bg=CARD)
        win.transient(self)
        win.grab_set()

        # center popup
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (win.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{x}+{y}")

        content = tk.Frame(win, bg=CARD, padx=20, pady=20)
        content.pack(fill="both", expand=True)

        tk.Label(
            content, text="📋  Ação necessária!",
            bg=CARD, fg=PRIMARY, font=("Segoe UI", 12, "bold")
        ).pack(pady=(0, 10))

        tk.Label(
            content,
            text="Para que o programa consiga enviar os dados, você precisa\n"
                 "compartilhar a sua nova planilha com o e-mail abaixo:",
            bg=CARD, fg=TEXT, font=("Segoe UI", 10),
            justify="center",
        ).pack(pady=(0, 15))

        email_frame = tk.Frame(content, bg="#F3F4F6", padx=10, pady=10)
        email_frame.pack(fill="x", pady=(0, 15))

        email_label = tk.Label(
            email_frame, text=self._target_email,
            bg="#F3F4F6", fg=TEXT, font=("Consolas", 10, "bold"),
            wraplength=400,
        )
        email_label.pack()

        def _copy():
            self.clipboard_clear()
            self.clipboard_append(self._target_email)
            copy_btn.config(text="Copiado!", state="disabled")
            self.after(2000, lambda: copy_btn.config(text="Copiar E-mail", state="normal"))

        btn_frame = tk.Frame(content, bg=CARD)
        btn_frame.pack()

        copy_btn = tk.Button(
            btn_frame, text="Copiar E-mail",
            command=_copy,
            bg=PRIMARY, fg="white", bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=16, pady=8, cursor="hand2",
            activebackground=PRIMARY_DK, activeforeground="white"
        )
        copy_btn.pack(side="left", padx=5)

        tk.Button(
            btn_frame, text="Entendi",
            command=win.destroy,
            bg=BG, fg=TEXT, bd=0,
            font=("Segoe UI", 9),
            padx=16, pady=8, cursor="hand2",
            activebackground=BORDER
        ).pack(side="left", padx=5)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Selecionar arquivos de lançamentos",
            filetypes=[
                ("Todos os suportados", "*.xlsx *.xlsm *.pdf"),
                ("Excel files", "*.xlsx *.xlsm"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*"),
            ],
        )
        for p in paths:
            fp = Path(p)
            # avoid duplicates
            if any(r.filepath == fp for r in self.file_rows):
                continue
            row = FileRow(self.file_list_frame, fp, self._remove_file)
            row.pack(fill="x", pady=1)
            self.file_rows.append(row)

        self._update_empty_label()

    def _remove_file(self, row: FileRow):
        row.destroy()
        self.file_rows.remove(row)
        self._update_empty_label()

    def _update_empty_label(self):
        if self.file_rows:
            self._empty_label.pack_forget()
        else:
            self._empty_label.pack(pady=16)

    def _refresh_preview(self):
        if not self.file_rows:
            messagebox.showinfo("Aviso", "Adicione pelo menos um arquivo de lançamentos.")
            return

        # Collect info before spawning thread (Tk vars must be read on main thread)
        files_info = []
        for row in self.file_rows:
            files_info.append((row.filepath, row.source_type, row.sheet_var.get()))
        api_key = self._api_key

        self._set_status("⏳  Analisando arquivos... aguarde.", color=WARNING)
        self.count_label.config(text="(processando...)")

        dre_id = self._dre_id

        def run():
            try:
                import warnings
                from ai_mapper import AITruncationWarning
                engine = DREEngine(dre_id=dre_id)
                all_rows = []
                trunc_warnings = []
                for filepath, source_type, sheet_target in files_info:
                    with warnings.catch_warnings(record=True) as caught:
                        warnings.simplefilter("always", AITruncationWarning)
                        rows = engine.extract_rows(
                            filepath, source_type, sheet_target, api_key=api_key
                        )
                    for w in caught:
                        if issubclass(w.category, AITruncationWarning):
                            trunc_warnings.append(f"{filepath.name}: {w.message}")
                    all_rows.extend(rows)

                categories = {}
                if all_rows:
                    self.after(0, lambda: self._set_status("🏷️  Classificando lançamentos com IA...", color=PRIMARY))
                    categories = engine.fetch_categories()
                    memory = engine.fetch_memory_rules()
                    from ai_mapper import classify_transactions
                    all_rows = classify_transactions(all_rows, categories, memory, api_key)

                self.after(0, lambda r=all_rows, c=categories: self._on_preview_success(r, c))
                if trunc_warnings:
                    warn_msg = "Atenção — conteúdo truncado:\n\n" + "\n".join(trunc_warnings)
                    self.after(0, lambda: messagebox.showwarning("Aviso de truncamento", warn_msg))
            except Exception as e:
                msg = str(e)
                self.after(0, lambda: self._on_preview_error(msg))

        threading.Thread(target=run, daemon=True).start()

    def _on_preview_success(self, rows: list[dict], categories: dict = None):
        self.preview_data = rows
        self.preview_table.load(self.preview_data, categories)
        n = len(self.preview_data)
        self.count_label.config(text=f"({n} lançamento{'s' if n != 1 else ''})")
        self._set_status(f"{n} lançamentos prontos para importar.")

    def _on_preview_error(self, msg: str):
        self.preview_data = []
        self.count_label.config(text="")
        self._set_status(f"❌  Erro: {msg}", color=DANGER)
        messagebox.showerror("Erro na pré-visualização", msg)

    def _start_import(self):
        if not self.preview_data:
            messagebox.showwarning("Atenção", "Clique em 'Atualizar prévia' antes de importar.")
            return

        confirm = messagebox.askyesno(
            "Confirmar importação",
            f"Importar {len(self.preview_data)} lançamentos para a planilha DRE no Google Drive?\n\nEsta ação não pode ser desfeita.",
        )
        if not confirm:
            return

        self.import_btn.config(state="disabled", text="⏳  Importando...")
        self._set_status("Importando...")

        # Commit user edits from the comboboxes into self.preview_data
        self.preview_table.commit_edits()

        # Learn from manual corrections
        rules_to_save = []
        for r in self.preview_data:
            new_cat = r.get("classificacao", "")
            old_cat = r.get("_original_cat", "")
            # Only save rule if user provided a valid category that is different from AI's initial guess
            if new_cat and new_cat != old_cat:
                rules_to_save.append((r["historico"], new_cat, r["sheet"]))

        dre_id = self._dre_id

        def run():
            try:
                engine = DREEngine(dre_id=dre_id)
                inserted = engine.write_to_dre(self.preview_data)
                
                # Save learned rules (run sequentially since write_to_dre finished successfully)
                if rules_to_save:
                    self.after(0, lambda: self._set_status("🧠  Salvando aprendizados da IA...", color=PRIMARY))
                    engine.save_memory_rules(rules_to_save)
                    
                self.after(0, lambda: self._on_import_success(inserted))
            except Exception as e:
                msg = str(e)
                print(f"[IMPORT ERROR] {msg}")
                self.after(0, lambda: self._on_import_error(msg))

        threading.Thread(target=run, daemon=True).start()

    def _on_import_success(self, inserted: int):
        self.import_btn.config(state="normal", text="⬆  Importar para o DRE")
        self._set_status(f"✅  {inserted} lançamentos importados com sucesso!", color=SUCCESS)
        messagebox.showinfo(
            "Sucesso!",
            f"{inserted} lançamentos foram inseridos na planilha DRE no Google Drive.",
        )

    def _on_import_error(self, msg: str):
        self.import_btn.config(state="normal", text="⬆  Importar para o DRE")
        self._set_status(f"❌  Erro: {msg}", color=DANGER)
        messagebox.showerror("Erro na importação", msg)

    def _set_status(self, msg: str, color: str = SUBTEXT):
        self.status_label.config(text=msg, fg=color)


if __name__ == "__main__":
    app = App()
    app.mainloop()