"""
DRE Fechamento Financeiro - Importador de Lançamentos
Versão 1.0
"""

import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
from engine import DREEngine, detect_source_type, ImportError

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

        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

    def load(self, rows: list[dict]):
        for widget in self.body.winfo_children():
            widget.destroy()

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

            classificacao = row.get("classificacao", "")
            tk.Label(
                self.body, text=classificacao, bg=bg, fg=PRIMARY if classificacao else SUBTEXT,
                font=("Segoe UI", 8, "bold" if classificacao else "normal"), anchor="w", padx=8,
            ).grid(row=i, column=3, sticky="ew")

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
        # .env OPENAI_KEY takes priority over saved config
        self._api_key: str = dotenv.get("OPENAI_API") or cfg.get("openai_api_key", "")

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

        # ── OpenAI API Key ──
        tk.Label(left, text="OpenAI API Key", bg=BG, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        api_card = tk.Frame(left, bg=CARD, relief="flat", bd=1,
                            highlightbackground=BORDER, highlightthickness=1)
        api_card.pack(fill="x", pady=(0, 12))

        self._api_key_var = tk.StringVar(value=self._api_key)
        api_entry = tk.Entry(
            api_card,
            textvariable=self._api_key_var,
            font=("Segoe UI", 9),
            bg=CARD, fg=TEXT,
            relief="flat",
            show="*",
        )
        api_entry.pack(side="left", fill="x", expand=True, padx=(10, 4), pady=6)

        self._show_key_var = tk.BooleanVar(value=False)

        def _toggle_show():
            api_entry.config(show="" if self._show_key_var.get() else "*")

        tk.Checkbutton(
            api_card, text="mostrar", variable=self._show_key_var,
            command=_toggle_show,
            bg=CARD, fg=SUBTEXT, font=("Segoe UI", 8),
            activebackground=CARD, bd=0, cursor="hand2",
        ).pack(side="left", padx=(0, 4))

        tk.Button(
            api_card, text="Salvar",
            command=self._save_api_key,
            bg=PRIMARY, fg="white", bd=0,
            font=("Segoe UI", 8, "bold"),
            padx=8, pady=4, cursor="hand2",
            activebackground=PRIMARY_DK, activeforeground="white",
        ).pack(side="right", padx=(0, 8), pady=4)

        # ── DRE target (Google Sheets — fixed) ──
        tk.Label(left, text="1. Planilha DRE (destino)", bg=BG, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        dre_card = tk.Frame(left, bg=CARD, relief="flat", bd=1,
                            highlightbackground=BORDER, highlightthickness=1)
        dre_card.pack(fill="x")

        tk.Label(
            dre_card,
            text="📄  Google Sheets — DRE Fechamento",
            bg="#EFF6FF", fg=PRIMARY,
            font=("Segoe UI", 9, "bold"),
            anchor="w", padx=10, pady=10,
        ).pack(fill="x")

        tk.Label(
            dre_card,
            text="Os lançamentos serão enviados diretamente ao Google Drive.",
            bg=CARD, fg=SUBTEXT,
            font=("Segoe UI", 8),
            anchor="w", padx=10,
        ).pack(fill="x", pady=(0, 8))

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

    def _save_api_key(self):
        key = self._api_key_var.get().strip()
        self._api_key = key
        cfg = _load_config()
        cfg["openai_api_key"] = key
        _save_config(cfg)
        self._set_status("API Key salva." if key else "API Key removida.")

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
        api_key = self._api_key_var.get().strip() or self._api_key

        self._set_status("⏳  Analisando arquivos... aguarde.", color=WARNING)
        self.count_label.config(text="(processando...)")

        def run():
            try:
                engine = DREEngine()
                all_rows = []
                for filepath, source_type, sheet_target in files_info:
                    rows = engine.extract_rows(
                        filepath, source_type, sheet_target, api_key=api_key
                    )
                    all_rows.extend(rows)
                
                if all_rows:
                    self.after(0, lambda: self._set_status("🏷️  Classificando lançamentos com IA...", color=PRIMARY))
                    categories = engine.fetch_categories()
                    from ai_mapper import classify_transactions
                    all_rows = classify_transactions(all_rows, categories, api_key)

                self.after(0, lambda r=all_rows: self._on_preview_success(r))
            except Exception as e:
                msg = str(e)
                self.after(0, lambda: self._on_preview_error(msg))

        threading.Thread(target=run, daemon=True).start()

    def _on_preview_success(self, rows: list[dict]):
        self.preview_data = rows
        self.preview_table.load(self.preview_data)
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

        def run():
            try:
                engine = DREEngine()
                inserted = engine.write_to_dre(self.preview_data)
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