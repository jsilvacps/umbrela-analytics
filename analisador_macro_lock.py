import os
import re
import sys
import threading
import subprocess
import uuid
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "Umbrela Analytics"
APP_VERSION = "8.34"
GITHUB_REPO = "jsilvacps/umbrela-analytics"   # repositório público de releases
DEFAULT_EXCLUDE = "OFFLINE, [LOCK.OFFLINE], [LOCK.LOGIN], LOGIN, LOGOUT"
LOGO_FILE = "logo_umbrela_analytics.png"
EXCEL_MAX_ROWS = 1_048_576
EXCEL_DATA_ROWS = EXCEL_MAX_ROWS - 1

# ── Regras de auto-classificação ─────────────────────────────────────────────
# Cada regra é (regex_pattern, status). Testado contra name.lower().
# Primeira regra que casar vence.
AUTO_CLASSIFY_RULES = [
    # ── PRODUTIVO ──────────────────────────────────────────────────────────
    # Governo federal/estadual/municipal
    (r'\.gov\.br', "productive"),
    (r'\.gov\b',   "productive"),
    (r'\bgov\.br\b', "productive"),
    (r'fazenda\.', "productive"),
    (r'receita\.', "productive"),
    (r'serpro\.',  "productive"),
    (r'dataprev',  "productive"),
    (r'\btrf\d*\b', "productive"),
    (r'\bstj\b',   "productive"),
    (r'\bstf\b',   "productive"),
    (r'\btse\b',   "productive"),
    (r'\bcnj\b',   "productive"),
    (r'\binss\b',  "productive"),
    (r'sefaz',     "productive"),
    (r'nfe\.',     "productive"),
    (r'nfse\.',    "productive"),
    (r'e-cac',     "productive"),
    (r'siafi',     "productive"),
    (r'siape',     "productive"),
    (r'comprasnet', "productive"),
    (r'licitacoes',  "productive"),
    (r'licitações',  "productive"),
    (r'prefeitura',  "productive"),
    (r'tribunal',    "productive"),
    # Microsoft 365 / infraestrutura
    (r'microsoft',       "productive"),
    (r'office\s*365',    "productive"),
    (r'sharepoint',      "productive"),
    (r'outlook',         "productive"),
    (r'\bteams\b',       "productive"),
    (r'onedrive',        "productive"),
    (r'azure',           "productive"),
    (r'live\.com',       "productive"),
    (r'microsoftonline', "productive"),
    (r'office\.com',     "productive"),
    (r'windows update',  "productive"),
    (r'microsoft store', "productive"),
    (r'intune',          "productive"),
    (r'defender',        "productive"),
    # Chrome / navegador
    (r'^google chrome$', "productive"),
    (r'^chrome$',        "productive"),
    (r'^chromium$',      "productive"),
    # Ferramentas de trabalho comuns
    (r'\bzoom\b',  "productive"),
    (r'\bwebex\b', "productive"),
    (r'\bslack\b', "productive"),
    (r'\bnotion\b', "productive"),
    (r'\bjira\b',  "productive"),
    (r'\bconfluence\b', "productive"),
    (r'\btrello\b', "productive"),
    (r'anydesk',   "productive"),
    (r'teamviewer', "productive"),
    (r'vnc',       "productive"),
    # ── IMPRODUTIVO ────────────────────────────────────────────────────────
    # Google pessoal
    (r'gmail',           "unproductive"),
    (r'google.*mail',    "unproductive"),
    (r'photos\.google',  "unproductive"),
    (r'google.*photo',   "unproductive"),
    (r'docs\.google',    "unproductive"),
    (r'sheets\.google',  "unproductive"),
    (r'slides\.google',  "unproductive"),
    (r'google.*search',  "unproductive"),
    (r'pesquisa.*google', "unproductive"),
    # Redes sociais
    (r'facebook',    "unproductive"),
    (r'instagram',   "unproductive"),
    (r'tiktok',      "unproductive"),
    (r'\btwitter\b', "unproductive"),
    (r'\bx\.com\b',  "unproductive"),
    (r'whatsapp',    "unproductive"),
    (r'telegram',    "unproductive"),
    # Streaming / entretenimento
    (r'youtube',   "unproductive"),
    (r'netflix',   "unproductive"),
    (r'spotify',   "unproductive"),
    (r'twitch',    "unproductive"),
    (r'globoplay', "unproductive"),
    (r'prime video', "unproductive"),
    (r'disneyplus', "unproductive"),
    # Jogos
    (r'steam',      "unproductive"),
    (r'epicgames',  "unproductive"),
    (r'riot games', "unproductive"),
    (r'blizzard',   "unproductive"),
    (r'ubisoft',    "unproductive"),
    (r'valorant',   "unproductive"),
    (r'fortnite',   "unproductive"),
    (r'\bleague of legends\b', "unproductive"),
    # Compras / lazer
    (r'mercadolivre', "unproductive"),
    (r'americanas',   "unproductive"),
    (r'magazineluiza', "unproductive"),
    (r'shopee',       "unproductive"),
    (r'aliexpress',   "unproductive"),
    (r'amazon\.com\.br', "unproductive"),
    (r'olx\.com\.br',    "unproductive"),
]

MACROLOCK_COLUMNS = [
    "DataHora", "Equipamento", "Usuario", "Status", "Aplicativo",
    "CampoExtra", "JanelaEvento", "DuracaoSegundos", "TempoFormatado",
]

EXPORT_COLUMNS = [
    "Data", "Hora", "DiaSemana", "Periodo", "Equipamento", "Usuario",
    "Status", "Aplicativo", "CampoExtra", "JanelaEvento", "DuracaoSegundos", "Duracao",
]

REPORTS = {
    "fora_horario": "Trabalho fora do horário - 18:00 até 06:59",
    "fim_semana": "Trabalho no final de semana - sábado e domingo",
    "fora_horario_fim_semana": "Fora do horário OU final de semana",
    "por_usuario_equipamento": "Resumo por usuário/equipamento",
    "intervalo_horario": "Filtro por intervalo de horário específico",
    "blacklist": "Acessos a termos da Blacklist",
    "ociosidade_diario": "Ociosidade diária por usuário (Seg–Sex 08–12h e 13–17h)",
    "ociosidade_mensal": "Ociosidade mensal por usuário (Seg–Sex 08–12h e 13–17h)",
}

# Segundos de horário comercial esperado por dia útil (08–12h + 13–17h = 8h)
BUSINESS_SECONDS_PER_DAY = 28_800


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def sql_path(path: Path) -> str:
    return str(path).replace("'", "''")


def safe_name(value: str) -> str:
    value = str(value or "").strip()
    value = re.sub(r'[\\/:*?"<>|]+', "_", value)
    value = re.sub(r"\s+", "_", value)
    return value[:90] or "relatorio"


class MacroLockAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1120x760")
        self.root.minsize(1020, 680)
        self.root.state("zoomed")

        # Ícone da janela (taskbar + título)
        _ico = app_dir() / "logo_umbrela_analytics.ico"
        if not _ico.exists():
            _ico = Path(__file__).parent / "logo_umbrela_analytics.ico"
        try:
            self.root.iconbitmap(str(_ico))
        except Exception:
            pass

        self.file_path = tk.StringVar()   # mantido para compatibilidade
        self._multi_files: list[str] = []  # lista de arquivos para processamento múltiplo

        # ── Integração Makrolock ──────────────────────────────────────────
        self._mk_session = None          # requests.Session autenticado
        self._mk_config_file = app_dir() / "mk_config.json"
        _cfg = self._mk_load_config()
        self._mk_email      = tk.StringVar(value=_cfg.get("email", ""))
        self._mk_save_pass_var = tk.BooleanVar(value=bool(_cfg.get("save_pass", False)))
        _saved_pass = self._mk_load_saved_pass() if _cfg.get("save_pass") else ""
        self._mk_pass       = tk.StringVar(value=_saved_pass)
        self._mk_status  = tk.StringVar(value="Não conectado")
        self._mk_logged  = False
        self._mk_pass_md5: str = ""   # hash em memória para re-login automático
        self._mk_base_url = "https://dmsempresarial.makrolock.com.br"
        self._mk_group_computers: dict[str, list[str]] = {}  # group_id → [computer names]
        self._mk_id_to_name: dict[str, str] = {}  # numeric id → computer name
        self._mk_computer_vars: dict[str, tk.BooleanVar] = {}  # computer_name → BooleanVar
        self._mk_date_start = tk.StringVar(value="")
        self._mk_date_end   = tk.StringVar(value="")
        self._mk_report_type = tk.StringVar(value="sessions")
        self._mk_busy    = False
        self.output_path = tk.StringVar(value=r"C:\Users\TI\Desktop\indicadores")
        self.output_filename = tk.StringVar(value="")
        # Controle de Horário
        self._mk_horario_monitor_ativo = False
        self._mk_horario_after_id = None
        self._mk_horario_rows: dict[str, dict] = {}  # cname → dados da linha na tabela
        self.exclude_terms = tk.StringVar(value=DEFAULT_EXCLUDE)
        self.status_text = tk.StringVar(value="Selecione o CSV exportado do Macro Lock.")
        self.progress_value = tk.DoubleVar(value=0)

        self.night_start = tk.StringVar(value="18:00")
        self.night_end = tk.StringVar(value="07:00")
        self.report_type = tk.StringVar(value="fora_horario")
        self.report_format    = tk.StringVar(value="xlsx")
        self.report_dashboard = tk.BooleanVar(value=True)
        self.report_user = tk.StringVar(value="")
        self.report_equipment = tk.StringVar(value="")
        self.base_export_format = tk.StringVar(value="csv_xlsx")
        self.report_date_start = tk.StringVar(value="")
        self.report_date_end = tk.StringVar(value="")

        self.metric_total = tk.StringVar(value="0")
        self.metric_night = tk.StringVar(value="0")
        self.metric_weekend = tk.StringVar(value="0")
        self.metric_users = tk.StringVar(value="0")
        self.metric_equips = tk.StringVar(value="0")
        self.metric_period = tk.StringVar(value="Sem base carregada")

        # rename_maps: {"Equipamento": {"nome_errado": "nome_correto"}, "Usuario": {...}}
        self.rename_maps = {"Equipamento": {}, "Usuario": {}}

        # blacklist: lista de termos bloqueados a verificar nos relatórios
        self.blacklist_terms = []
        self._blacklist_file = app_dir() / "blacklist.json"
        self._load_blacklist()

        self._build_ui()
        # Verifica atualização em background após a UI estar pronta
        self.root.after(3000, lambda: threading.Thread(
            target=self._check_update, daemon=True).start())

    # ══════════════════════════════════════════════════════════════════════
    #  AUTO-UPDATE — verifica última release no GitHub
    # ══════════════════════════════════════════════════════════════════════
    def _check_update(self):
        """Consulta GitHub Releases e oferece atualização se houver versão nova."""
        try:
            import urllib.request, json as _json, tempfile, shutil
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            req = urllib.request.Request(url, headers={"User-Agent": "UmbrelaAnalytics"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = _json.loads(resp.read())

            tag = data.get("tag_name", "").lstrip("v")
            body = data.get("body", "")
            assets = data.get("assets", [])

            # Compara versões (ex: "7.1" > "7.0")
            def _ver(s):
                try: return tuple(int(x) for x in s.split("."))
                except Exception: return (0,)

            if _ver(tag) <= _ver(APP_VERSION):
                return  # já está atualizado

            # Localiza o .exe nos assets
            exe_asset = next(
                (a for a in assets if a["name"].lower().endswith(".exe")
                 and "setup" not in a["name"].lower()), None)
            if not exe_asset:
                return

            dl_url  = exe_asset["browser_download_url"]
            dl_size = exe_asset.get("size", 0)
            size_mb = f"{dl_size/1_048_576:.1f} MB" if dl_size else ""

            msg = (
                f"Nova versão disponível: v{tag}\n\n"
                f"{body[:300].strip()}\n\n" if body else
                f"Nova versão disponível: v{tag}\n\n"
            ) + f"Tamanho: {size_mb}\n\nDeseja baixar e atualizar agora?"

            self.root.after(0, lambda: self._offer_update(tag, dl_url, msg))

        except Exception:
            pass  # falha silenciosa — sem internet ou repo privado

    def _offer_update(self, tag: str, dl_url: str, msg: str):
        if not messagebox.askyesno(
            f"Atualização disponível — v{tag}",
            msg, icon="question"
        ):
            return
        threading.Thread(target=self._do_update, args=(dl_url,), daemon=True).start()

    def _do_update(self, dl_url: str):
        """Baixa o novo EXE e substitui o atual, depois reinicia."""
        import urllib.request, tempfile, os, sys
        try:
            self.root.after(0, lambda: self.set_status("Baixando atualização...", 10))
            tmp = tempfile.mktemp(suffix=".exe")

            def _report(n, bs, total):
                if total > 0:
                    pct = min(90, int(n * bs / total * 90))
                    self.root.after(0, lambda p=pct: self.set_status(
                        f"Baixando... {p}%", p))

            urllib.request.urlretrieve(dl_url, tmp, reporthook=_report)
            self.root.after(0, lambda: self.set_status("Instalando...", 95))

            current = sys.executable if getattr(sys, "frozen", False) else None
            if current and os.path.exists(current):
                # Windows: agenda substituição via script bat executado após encerrar
                bat = tempfile.mktemp(suffix=".bat")
                with open(bat, "w") as f:
                    f.write(
                        f'@echo off\n'
                        f'timeout /t 2 /nobreak >nul\n'
                        f'move /y "{tmp}" "{current}"\n'
                        f'start "" "{current}"\n'
                        f'del "%~f0"\n'
                    )
                import subprocess
                subprocess.Popen(["cmd", "/c", bat],
                                 creationflags=subprocess.CREATE_NO_WINDOW)
                self.root.after(0, self.root.destroy)
            else:
                # Modo dev: apenas informa onde foi salvo
                self.root.after(0, lambda: messagebox.showinfo(
                    "Atualização baixada",
                    f"Novo executável salvo em:\n{tmp}\n\nSubstitua manualmente."))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(
                "Erro na atualização", str(e)))
            self.root.after(0, lambda: self.set_status("Erro ao atualizar.", 0))

    def _load_logo(self, max_w=330):
        """Carrega a logo PNG (subsample simples para header)."""
        try:
            logo_path = app_dir() / LOGO_FILE
            if logo_path.exists():
                img = tk.PhotoImage(file=str(logo_path))
                if img.width() > max_w:
                    factor = max(1, int(img.width() / max_w))
                    img = img.subsample(factor, factor)
                self._logo_img = img
                return img
        except Exception:
            pass
        self._logo_img = None
        return None

    def _load_logo_hd(self, target_w=480):
        """Carrega a logo em alta definição usando PIL para redimensionar."""
        try:
            from PIL import Image, ImageTk
            logo_path = app_dir() / LOGO_FILE
            if logo_path.exists():
                pil_img = Image.open(str(logo_path))
                orig_w, orig_h = pil_img.size
                target_h = int(orig_h * target_w / orig_w)
                pil_img = pil_img.resize((target_w, target_h), Image.LANCZOS)
                img = ImageTk.PhotoImage(pil_img)
                self._logo_hd_img = img
                return img
        except Exception:
            pass
        return self._load_logo(max_w=target_w)

    def _bind_time_mask(self, entry_widget, str_var, next_widget=None):
        """Aplica máscara HH:MM com : automático."""
        _busy = [False]; _prev = [""]
        def _on_write(*_):
            if _busy[0]: return
            _busy[0] = True
            try:
                raw = str_var.get(); growing = len(raw) > len(_prev[0])
                digits = "".join(c for c in raw if c.isdigit())[:4]
                if len(digits) <= 2:
                    masked = digits + (":" if len(digits) == 2 and growing else "")
                else:
                    masked = digits[:2] + ":" + digits[2:4]
                if masked != raw:
                    _prev[0] = masked; str_var.set(masked)
                    entry_widget.icursor(len(masked))
                else:
                    _prev[0] = raw
                if next_widget and len(masked) == 5 and growing:
                    entry_widget.after(0, next_widget.focus_set)
            finally:
                _busy[0] = False
        str_var.trace_add("write", _on_write)

    def _bind_date_mask(self, entry_widget, str_var, next_widget=None):
        """Aplica máscara DD/MM/AAAA com / automático e cursor correto."""
        _busy  = [False]
        _prev  = [""]

        def _fmt(digits, growing):
            n = len(digits)
            if n <= 1:
                return digits
            if n == 2:
                # Auto-adiciona / ao avançar; mantém se já estava lá
                return digits + "/" if growing else digits
            if n <= 4:
                return digits[:2] + "/" + digits[2:]
            if n == 4 and growing:
                return digits[:2] + "/" + digits[2:4] + "/"
            return digits[:2] + "/" + digits[2:4] + "/" + digits[4:8]

        def _on_write(*_):
            if _busy[0]:
                return
            _busy[0] = True
            try:
                raw     = str_var.get()
                growing = len(raw) > len(_prev[0])
                digits  = "".join(c for c in raw if c.isdigit())[:8]
                # Detecta n=4 depois de formatar para adicionar segundo /
                masked  = _fmt(digits, growing)
                if len(digits) == 4 and growing and not masked.endswith("/"):
                    masked += "/"
                if masked != raw:
                    _prev[0] = masked
                    str_var.set(masked)
                    entry_widget.icursor(len(masked))
                else:
                    _prev[0] = raw
                # Avança foco automaticamente quando data completa (DD/MM/AAAA = 10 chars)
                if next_widget and len(masked) == 10 and growing:
                    entry_widget.after(0, next_widget.focus_set)
            finally:
                _busy[0] = False

        str_var.trace_add("write", _on_write)

    def _build_ui(self):
        # ── Paleta de cores premium ───────────────────────────────────────
        self.COLORS = {
            # ── Fundo dark navy ────────────────────────────────────────────
            "bg":       "#070D1C",   # fundo global — quase preto azulado
            "sidebar":  "#080E1F",   # sidebar levemente mais escura
            "sidehov":  "#0D1830",   # hover na sidebar
            "sidesel":  "#1035A0",   # selecionado — azul vivo
            "panel":    "#0C1628",   # painel secundário
            "card":     "#0C1628",   # cards
            "card2":    "#0F1D35",   # cards secundários
            "border":   "#162040",   # borda suave
            "border2":  "#1E2E56",   # borda mais visível
            # ── Texto ─────────────────────────────────────────────────────
            "text":     "#FFFFFF",   # texto principal — branco puro
            "muted":    "#7EB8D4",   # texto secundário — azul claro
            "muted2":   "#A8C8E8",   # labels e dicas — azul claro brilhante
            # ── Accent / azul ─────────────────────────────────────────────
            "accent":   "#2563EB",
            "accent2":  "#1D4ED8",
            "accentbg": "#0D1F4A",
            # ── Status colors ─────────────────────────────────────────────
            "ok":       "#10B981",
            "okbg":     "#062918",
            "warning":  "#F59E0B",
            "danger":   "#EF4444",
            # ── Paleta de tiles ───────────────────────────────────────────
            "purple":   "#8B5CF6",
            "purplebg": "#1A0A40",
            "teal":     "#06B6D4",
            "tealbg":   "#031C2E",
            "orange":   "#F97316",
            "orangebg": "#2A1200",
            "pink":     "#EC4899",
            "pinkbg":   "#2A0A1A",
        }
        c = self.COLORS
        self.root.configure(bg=c["bg"])

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # ── Frames ──────────────────────────────────────────────────────
        for name, bg in (
            ("TFrame",       c["bg"]),
            ("Card.TFrame",  c["card"]),
            ("Panel.TFrame", c["panel"]),
            ("Bg.TFrame",    c["bg"]),
        ):
            style.configure(name, background=bg, relief="flat", borderwidth=0)

        # ── Labels ──────────────────────────────────────────────────────
        style.configure("TLabel",
            background=c["bg"],   foreground=c["text"], font=("Segoe UI", 10))
        style.configure("Card.TLabel",
            background=c["card"], foreground=c["text"], font=("Segoe UI", 10))
        style.configure("Muted.TLabel",
            background=c["card"], foreground=c["muted"], font=("Segoe UI", 9))
        style.configure("Metric.TLabel",
            background=c["card"], foreground=c["teal"], font=("Segoe UI", 24, "bold"))
        style.configure("Panel.TLabel",
            background=c["panel"], foreground=c["text"], font=("Segoe UI", 10))

        # ── Botões ──────────────────────────────────────────────────────
        style.configure("TButton",
            font=("Segoe UI", 9), padding=(10, 6),
            background=c["card2"], foreground=c["muted2"],
            relief="flat", borderwidth=1)
        style.map("TButton",
            background=[("active", c["border2"]), ("pressed", c["border"])],
            foreground=[("active", c["text"])])

        style.configure("Accent.TButton",
            font=("Segoe UI", 10, "bold"), padding=(14, 8),
            background=c["teal"], foreground="#ffffff",
            relief="flat", borderwidth=0)
        style.map("Accent.TButton",
            background=[("active", "#0891B2"), ("pressed", "#0E7490"),
                        ("disabled", "#1E4A5A")])

        style.configure("Ok.TButton",
            font=("Segoe UI", 10, "bold"), padding=(14, 8),
            background=c["ok"], foreground="#ffffff",
            relief="flat", borderwidth=0)
        style.map("Ok.TButton",
            background=[("active", "#059669"), ("pressed", "#047857")])

        style.configure("Ghost.TButton",
            font=("Segoe UI", 9), padding=(10, 6),
            background=c["card"], foreground=c["muted2"],
            relief="flat", borderwidth=1)
        style.map("Ghost.TButton",
            background=[("active", c["card2"])],
            foreground=[("active", c["text"])])

        style.configure("Danger.TButton",
            font=("Segoe UI", 9), padding=(10, 6),
            background=c["card"], foreground=c["danger"],
            relief="flat", borderwidth=1)
        style.map("Danger.TButton",
            background=[("active", "#2D0A0A")])

        # ── Inputs ──────────────────────────────────────────────────────
        style.configure("TEntry",
            padding=(10, 7),
            fieldbackground=c["card2"], foreground=c["text"],
            insertcolor=c["teal"],
            bordercolor=c["border2"], lightcolor=c["border2"], darkcolor=c["border2"])
        style.configure("TCombobox",
            padding=(10, 7),
            fieldbackground=c["card2"], foreground=c["text"],
            bordercolor=c["border2"])
        style.map("TCombobox",
            fieldbackground=[("readonly", c["card2"])],
            selectbackground=[("readonly", c["accentbg"])],
            selectforeground=[("readonly", c["text"])])

        style.configure("TRadiobutton",
            background=c["card"], foreground=c["text"], font=("Segoe UI", 10))

        style.configure("TCheckbutton",
            background=c["card"], foreground=c["text"], font=("Segoe UI", 10))

        # ── Progressbar ─────────────────────────────────────────────────
        style.configure("Horizontal.TProgressbar",
            troughcolor=c["border"], background=c["teal"],
            bordercolor=c["border"], lightcolor=c["teal"],
            darkcolor=c["teal"], thickness=4)

        # ── Treeview ────────────────────────────────────────────────────
        style.configure("Treeview",
            background=c["card"], fieldbackground=c["card"],
            foreground=c["text"], font=("Segoe UI", 10),
            rowheight=32, borderwidth=0)
        style.configure("Treeview.Heading",
            background=c["card2"], foreground=c["muted2"],
            font=("Segoe UI", 9, "bold"), relief="flat",
            bordercolor=c["border"])
        style.map("Treeview",
            background=[("selected", c["accentbg"])],
            foreground=[("selected", c["teal"])])

        # ═══════════════════════════════════════════════════════════════
        # Raiz: shell → header + body
        # ═══════════════════════════════════════════════════════════════
        shell = tk.Frame(self.root, bg=c["bg"])
        shell.pack(fill="both", expand=True)

        # ── Header premium ────────────────────────────────────────────
        header = tk.Frame(shell, bg=c["sidebar"], height=58)
        header.pack(fill="x")
        header.pack_propagate(False)

        # Linha topo: teal fino e elegante
        tk.Frame(header, bg=c["teal"], height=2).pack(fill="x", side="top")

        # Logo no header — max_w=80 para caber nos 58px de altura do header
        logo = self._load_logo(max_w=80)
        if logo is not None:
            lbl_hdr_logo = tk.Label(header, image=logo, bg=c["sidebar"], borderwidth=0)
            lbl_hdr_logo.image = logo  # mantém referência
            lbl_hdr_logo.pack(side="left", padx=(18, 8), pady=6)

        # Brand
        brand_frame = tk.Frame(header, bg=c["sidebar"])
        brand_frame.pack(side="left", pady=0)
        tk.Label(brand_frame, text=APP_NAME, bg=c["sidebar"], fg="#FFFFFF",
                 font=("Segoe UI", 14, "bold")).pack(side="left")
        tk.Label(brand_frame, text=f"v{APP_VERSION}",
                 bg="#1E3A6E", fg="#93C5FD",
                 font=("Segoe UI", 8, "bold"), padx=7, pady=3).pack(
                     side="left", padx=(10, 0))

        # Separador + subtítulo
        tk.Frame(header, bg="#1E3A6E", width=1).pack(
            side="left", fill="y", pady=14, padx=18)
        sub_frame = tk.Frame(header, bg=c["sidebar"])
        sub_frame.pack(side="left")
        self._hdr_title_lbl = tk.Label(sub_frame, text="Início", bg=c["sidebar"],
                 fg="#B8D8F8", font=("Segoe UI", 9))
        self._hdr_title_lbl.pack(anchor="w")
        self._hdr_sub_lbl = tk.Label(sub_frame, text="Umbrela Analytics", bg=c["sidebar"],
                 fg="#6B8EB0", font=("Segoe UI", 8))
        self._hdr_sub_lbl.pack(anchor="w")

        # Lado direito: indicador de status do sistema
        right_hdr = tk.Frame(header, bg=c["sidebar"])
        right_hdr.pack(side="right", padx=22)

        self._hdr_status_dot = tk.Label(right_hdr, text="●", bg=c["sidebar"],
                                        fg=c["muted"], font=("Segoe UI", 10))
        self._hdr_status_dot.pack(side="right", padx=(4, 0))
        self._hdr_status_lbl = tk.Label(right_hdr, text="Sem base",
                                        bg=c["sidebar"], fg=c["muted"],
                                        font=("Segoe UI", 9))
        self._hdr_status_lbl.pack(side="right")

        # Avatar / usuário
        tk.Frame(header, bg="#1E3A6E", width=1).pack(
            side="right", fill="y", pady=14, padx=14)
        user_frame = tk.Frame(header, bg=c["sidebar"])
        user_frame.pack(side="right")
        avatar = tk.Label(user_frame, text="JS", bg=c["accent"],
                          fg="white", font=("Segoe UI", 9, "bold"),
                          width=3, pady=4)
        avatar.pack(side="right", padx=(6, 0))
        u_info = tk.Frame(user_frame, bg=c["sidebar"])
        u_info.pack(side="right")
        tk.Label(u_info, text="Jean Silva", bg=c["sidebar"],
                 fg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(anchor="e")
        tk.Label(u_info, text="Analista de TI",
                 bg=c["sidebar"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(anchor="e")

        # ── Corpo: sidebar + conteúdo ────────────────────────────────────
        body = tk.Frame(shell, bg=c["bg"])
        body.pack(fill="both", expand=True)

        # ── Sidebar ─────────────────────────────────────────────────────
        sidebar = tk.Frame(body, bg=c["sidebar"], width=200)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Divisor topo sidebar
        tk.Frame(sidebar, bg="#0A1020", height=1).pack(fill="x")

        # Seção de navegação
        tk.Label(sidebar, text="MENU", bg=c["sidebar"], fg="#4A7AAA",
                 font=("Segoe UI", 7, "bold")).pack(
                     anchor="w", padx=20, pady=(16, 4))

        # Área de conteúdo
        content_area = tk.Frame(body, bg=c["bg"])
        content_area.pack(side="left", fill="both", expand=True)

        # ── Barra inferior: status + log ─────────────────────────────────
        bottom = tk.Frame(content_area, bg=c["bg"])
        bottom.pack(side="bottom", fill="x")

        tk.Frame(bottom, bg=c["border"], height=1).pack(fill="x")

        # Status row
        status_row = tk.Frame(bottom, bg=c["sidebar"], pady=5)
        status_row.pack(fill="x", padx=0)
        tk.Frame(status_row, bg=c["border"], height=1).pack(fill="x")
        inner_status = tk.Frame(status_row, bg=c["sidebar"])
        inner_status.pack(fill="x", padx=16, pady=(5, 4))
        ttk.Progressbar(inner_status, variable=self.progress_value, maximum=100,
                        style="Horizontal.TProgressbar", length=200).pack(side="left")
        tk.Label(inner_status, textvariable=self.status_text,
                 bg=c["sidebar"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 0))

        # Log oculto — mantido para write_log(), sem exibir na UI
        self.log = tk.Text(bottom, height=1, state="normal",
                           bg=c["sidebar"], fg=c["sidebar"],
                           relief="flat", borderwidth=0, font=("Consolas", 8))
        # Tag para colorir mensagens de log
        self.log.tag_configure("ok",    foreground="#34D399")
        self.log.tag_configure("warn",  foreground="#FCD34D")
        self.log.tag_configure("error", foreground="#F87171")
        self.log.tag_configure("info",  foreground="#94A3B8")

        # Páginas
        self._pages = {}
        self._nav_items = {}
        self._cur_page = None

        nav_list = [
            ("inicio",     "⌂", "Início"),
            ("makrolock",  "⬡", "Makrolock"),
            ("processar",  "⚙", "Processar"),
            ("relatorios", "≡", "Relatórios"),
        ]
        for key, icon, label in nav_list:
            page = tk.Frame(content_area, bg=c["bg"], padx=24, pady=20)
            self._pages[key] = page

        _page_titles = {
            "inicio":     ("Início",     "Umbrela Analytics"),
            "makrolock":  ("Makrolock",  "Monitoramento de acessos"),
            "processar":  ("Processar",  "Importar e tratar dados"),
            "relatorios": ("Relatórios", "Análise e exportação"),
        }

        def _show(key):
            if self._cur_page:
                self._pages[self._cur_page].pack_forget()
            self._pages[key].pack(fill="both", expand=True)
            self._cur_page = key
            # Atualiza breadcrumb no header
            title, sub = _page_titles.get(key, (key.capitalize(), ""))
            if hasattr(self, "_hdr_title_lbl"):
                self._hdr_title_lbl.configure(text=title)
            if hasattr(self, "_hdr_sub_lbl"):
                self._hdr_sub_lbl.configure(text=sub)
            for k, widgets in self._nav_items.items():
                sel = (k == key)
                bg_n  = c["sidesel"] if sel else c["sidebar"]
                fg_n  = "#FFFFFF"   if sel else "#8AAED4"
                ind_c = c["accent"] if sel else c["sidebar"]
                for w, attr in widgets:
                    if attr == "ind":
                        w.configure(bg=ind_c)
                    elif attr == "bg":
                        w.configure(bg=bg_n)
                    elif attr == "fg":
                        w.configure(bg=bg_n, fg=fg_n)

        self._show_page = _show

        for key, icon, label in nav_list:
            outer_n = tk.Frame(sidebar, bg=c["sidebar"], cursor="hand2")
            outer_n.pack(fill="x", pady=1)

            ind = tk.Frame(outer_n, bg=c["sidebar"], width=3)
            ind.pack(side="left", fill="y")

            inner_n = tk.Frame(outer_n, bg=c["sidebar"], pady=10, padx=12)
            inner_n.pack(side="left", fill="both", expand=True)

            icon_lbl = tk.Label(inner_n, text=icon, bg=c["sidebar"], fg="#8AAED4",
                                font=("Segoe UI", 11), width=2, cursor="hand2")
            icon_lbl.pack(side="left")
            text_lbl = tk.Label(inner_n, text=label, bg=c["sidebar"], fg="#8AAED4",
                                font=("Segoe UI", 10), cursor="hand2")
            text_lbl.pack(side="left", padx=(8, 0))

            wlist = [
                (ind,      "ind"),
                (outer_n,  "bg"),
                (inner_n,  "bg"),
                (icon_lbl, "fg"),
                (text_lbl, "fg"),
            ]
            self._nav_items[key] = wlist

            cb = lambda e, k=key: _show(k)
            for w, _ in wlist:
                w.bind("<Button-1>", cb)
            # Hover
            def _hover_in(e, wl=wlist):
                if self._cur_page != [k for k, ww in self._nav_items.items() if ww is wl][0]:
                    for w, attr in wl:
                        if attr in ("bg",): w.configure(bg=c["sidehov"])
                        if attr == "fg": w.configure(bg=c["sidehov"])
            def _hover_out(e, wl=wlist):
                k_this = [k for k, ww in self._nav_items.items() if ww is wl][0]
                if self._cur_page != k_this:
                    for w, attr in wl:
                        if attr in ("bg",): w.configure(bg=c["sidebar"])
                        if attr == "fg": w.configure(bg=c["sidebar"])
            for w, _ in wlist:
                w.bind("<Enter>", _hover_in)
                w.bind("<Leave>", _hover_out)

        # Separador + versão no rodapé da sidebar
        tk.Label(sidebar, text=f"v{APP_VERSION} · © 2026 Jean Silva",
                 bg=c["sidebar"], fg="#2D5A8A",
                 font=("Segoe UI", 7)).pack(side="bottom", pady=(4, 8))

        # Linha de brilho sutil no rodapé da sidebar
        tk.Frame(sidebar, bg=c["border2"], height=1).pack(side="bottom", fill="x")
        tk.Frame(sidebar, bg=c["teal"],    height=1).pack(side="bottom", fill="x")

        # Constrói páginas
        tab_inicio    = self._pages["inicio"]
        tab_makrolock = self._pages["makrolock"]
        tab_process   = self._pages["processar"]
        tab_reports   = self._pages["relatorios"]

        self._build_inicio_tab(tab_inicio)
        self._build_makrolock_tab(tab_makrolock)
        self._build_process_tab(tab_process)
        self._build_reports_tab(tab_reports)

        _show("inicio")

    def _card(self, parent, padding=(16, 14), accent=False, color=None):
        """Card dark com borda sutil. accent=True adiciona faixa colorida à esquerda."""
        c = self.COLORS
        col = color or c["teal"]
        outer = tk.Frame(parent, bg=c["border"],
                         highlightthickness=1, highlightbackground=c["border2"])
        if accent:
            strip = tk.Frame(outer, bg=col, width=4)
            strip.pack(side="left", fill="y")
        inner = ttk.Frame(outer, style="Card.TFrame", padding=padding)
        inner.pack(side="left", fill="both", expand=True)
        return outer, inner

    def _section_label(self, parent, text, bg=None):
        bg = bg or self.COLORS["card"]
        tk.Label(parent, text=text, bg=bg, fg=self.COLORS["muted"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(10, 4))

    def _metric_card(self, parent, title, variable, subtitle=None,
                     icon="", color=None, bg_color=None):
        """Tile KPI estilo dark: ícone em círculo colorido, número grande, glow sutil."""
        c = self.COLORS
        color    = color    or c["teal"]
        bg_color = bg_color or c["tealbg"]

        # Borda externa (simula brilho)
        wrap = tk.Frame(parent, bg=color, padx=1, pady=1)
        wrap.pack_propagate(True)

        # Card principal dark
        card = tk.Frame(wrap, bg=c["card"])
        card.pack(fill="both", expand=True)

        body = tk.Frame(card, bg=c["card"], padx=18, pady=16)
        body.pack(fill="both", expand=True)

        # Linha topo: ícone + título
        top_row = tk.Frame(body, bg=c["card"])
        top_row.pack(fill="x", pady=(0, 10))

        if icon:
            ic_frm = tk.Frame(top_row, bg=bg_color, width=36, height=36)
            ic_frm.pack_propagate(False)
            ic_frm.pack(side="left")
            tk.Label(ic_frm, text=icon, bg=bg_color, fg=color,
                     font=("Segoe UI", 14)).place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(top_row, text=title.upper(), bg=c["card"],
                 fg=c["muted2"], font=("Segoe UI", 7, "bold"),
                 wraplength=120).pack(
                     side="left", padx=(10 if icon else 0, 0))

        # Valor numérico grande
        tk.Label(body, textvariable=variable, bg=c["card"],
                 fg="#FFFFFF", font=("Segoe UI", 28, "bold")).pack(anchor="w")

        if subtitle:
            tk.Label(body, text=subtitle, bg=c["card"],
                     fg=c["muted"], font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        # Faixa de cor na base do card (glow simulado)
        tk.Frame(card, bg=bg_color, height=3).pack(fill="x", side="bottom")

        return wrap

    def _build_inicio_tab(self, parent):
        """Tela inicial — logo centralizada, slogan e rodapé de marca."""
        c = self.COLORS

        # Remove o padding interno da página para usar todo o espaço
        parent.configure(padx=0, pady=0)

        # Coluna central para centralizar tudo verticalmente
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        center = tk.Frame(parent, bg=c["bg"])
        center.grid(row=0, column=0, sticky="nsew")
        center.columnconfigure(0, weight=1)
        center.rowconfigure(0, weight=1)

        inner = tk.Frame(center, bg=c["bg"])
        inner.grid(row=0, column=0)

        # ── Logo HD centralizada ───────────────────────────────────────
        logo_img = self._load_logo_hd(target_w=480)
        if logo_img is not None:
            lbl = tk.Label(inner, image=logo_img, bg=c["bg"], borderwidth=0)
            lbl.image = logo_img
            lbl.pack(anchor="center", pady=(0, 20))
        else:
            tk.Label(inner, text="Umbrela Analytics",
                     bg=c["bg"], fg=c["text"],
                     font=("Segoe UI", 40, "bold")).pack(anchor="center", pady=(0, 20))

        # ── Linha decorativa ──────────────────────────────────────────
        line_frm = tk.Frame(inner, bg=c["bg"])
        line_frm.pack(anchor="center", pady=(0, 14))
        tk.Frame(line_frm, bg=c["teal"],   width=80, height=2).pack(side="left")
        tk.Frame(line_frm, bg=c["accent"], width=40, height=2).pack(side="left")
        tk.Frame(line_frm, bg=c["purple"], width=24, height=2).pack(side="left")

        # ── Slogan ────────────────────────────────────────────────────
        tk.Label(inner,
                 text="Monitoramento · Análise · Produtividade",
                 bg=c["bg"], fg=c["muted2"],
                 font=("Segoe UI", 13)).pack(anchor="center")

        # ── Rodapé de marca ───────────────────────────────────────────
        tk.Frame(inner, bg=c["border2"], height=1, width=400).pack(anchor="center", pady=24)
        tk.Label(inner,
                 text="Desenvolvido por Umbrela Tecnologia",
                 bg=c["bg"], fg=c["muted"],
                 font=("Segoe UI", 10)).pack(anchor="center")
        tk.Label(inner, text=f"v{APP_VERSION}",
                 bg=c["bg"], fg=c["border2"],
                 font=("Segoe UI", 8)).pack(anchor="center", pady=(6, 0))

    def refresh_dashboard(self):
        # Procura o banco mais recente dentro das subpastas de data
        base = Path(self.output_path.get())
        db_candidates = sorted(base.rglob("base_macro_lock.duckdb"), key=lambda p: p.stat().st_mtime, reverse=True)
        output_dir = db_candidates[0].parent if db_candidates else self._dated_output_dir()
        try:
            con = self._connect_db(output_dir)
            total, night, weekend, users, equips, min_dt, max_dt = con.execute("""
                SELECT
                    COUNT(*) AS total,
                    SUM(CASE WHEN CAST(DataHoraTs AS TIME) >= TIME '18:00:00' OR CAST(DataHoraTs AS TIME) < TIME '07:00:00' THEN 1 ELSE 0 END) AS night,
                    SUM(CASE WHEN CAST(strftime(DataHoraTs, '%w') AS INTEGER) IN (0, 6) THEN 1 ELSE 0 END) AS weekend,
                    COUNT(DISTINCT Usuario) AS users,
                    COUNT(DISTINCT Equipamento) AS equips,
                    MIN(DataHoraTs)::VARCHAR,
                    MAX(DataHoraTs)::VARCHAR
                FROM tratado
            """).fetchone()
            self.metric_total.set(f"{total:,}".replace(",", "."))
            self.metric_night.set(f"{night or 0:,}".replace(",", "."))
            self.metric_weekend.set(f"{weekend or 0:,}".replace(",", "."))
            self.metric_users.set(f"{users:,}".replace(",", "."))
            self.metric_equips.set(f"{equips:,}".replace(",", "."))
            self.metric_period.set(f"{min_dt or '-'} até {max_dt or '-'}")
            con.close()
            self.set_status("Dashboard atualizado.", 100)
        except Exception as e:
            self.write_log(str(e))
            messagebox.showwarning("Dashboard", "Ainda não encontrei uma base tratada. Processe o CSV primeiro.")

    # ══════════════════════════════════════════════════════════════════════
    #  ABA MAKROLOCK — login → modal de escolha → relatórios | classificação
    # ══════════════════════════════════════════════════════════════════════
    def _build_makrolock_tab(self, parent):
        c = self.COLORS

        # Dados internos de classificação (precisam existir antes dos frames)
        self._mk_cls_items: list[tuple[str, str]] = []
        self._mk_cls_filtered: list[int] = []
        self._mk_cls_active_terms: list[str] = []

        # Container principal — empilha os 3 sub-frames
        self._mk_container = tk.Frame(parent, bg=c["bg"])
        self._mk_container.pack(fill="both", expand=True)

        # Cria os frames
        self._mk_frame_login    = self._build_mk_login_frame(self._mk_container)
        self._mk_frame_choice   = self._build_mk_choice_frame(self._mk_container)
        self._mk_frame_reports  = self._build_mk_reports_frame(self._mk_container)
        self._mk_frame_classify = self._build_mk_classify_frame(self._mk_container)
        self._mk_frame_horario  = self._build_mk_horario_frame(self._mk_container)

        # Começa na tela de login
        self._mk_show_login_frame()

    # ── Navegação entre frames ────────────────────────────────────────────
    def _mk_hide_all(self):
        for f in (self._mk_frame_login, self._mk_frame_choice,
                  self._mk_frame_reports, self._mk_frame_classify,
                  self._mk_frame_horario):
            f.pack_forget()

    def _mk_show_login_frame(self):
        self._mk_hide_all()
        self._mk_frame_login.pack(fill="both", expand=True)

    def _mk_show_choice_frame(self):
        self._mk_hide_all()
        # Atualiza o label de boas-vindas se existir
        email = getattr(self, "_mk_email", None)
        if email and hasattr(self, "_mk_choice_welcome_var"):
            em = email.get().strip() or "usuário"
            self._mk_choice_welcome_var.set(f"Bem-vindo, {em}")
        self._mk_frame_choice.pack(fill="both", expand=True)

    def _mk_show_reports_frame(self):
        self._mk_hide_all()
        self._mk_frame_reports.pack(fill="both", expand=True)

    def _mk_show_classify_frame(self):
        self._mk_hide_all()
        self._mk_frame_classify.pack(fill="both", expand=True)

    def _mk_show_horario_frame(self):
        self._mk_hide_all()
        self._mk_frame_horario.pack(fill="both", expand=True)

    def _mk_show_choice_modal(self):
        """Mantido para compatibilidade — redireciona para o frame integrado."""
        self._mk_show_choice_frame()

    def _build_mk_choice_frame(self, parent):
        """Tela de escolha pós-login — dark, elegante, sem popup."""
        c = self.COLORS
        frame = tk.Frame(parent, bg=c["bg"])

        # Coluna central
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=0)
        frame.columnconfigure(2, weight=1)
        frame.rowconfigure(0, weight=1)

        center = tk.Frame(frame, bg=c["bg"])
        center.grid(row=0, column=1, sticky="n", pady=60)

        # Header
        self._mk_choice_welcome_var = tk.StringVar(value="Bem-vindo")
        tk.Label(center, textvariable=self._mk_choice_welcome_var,
                 bg=c["bg"], fg="#FFFFFF",
                 font=("Segoe UI", 20, "bold")).pack(anchor="center")
        tk.Label(center, text="O que você deseja fazer no Makrolock?",
                 bg=c["bg"], fg=c["muted"],
                 font=("Segoe UI", 10)).pack(anchor="center", pady=(4, 36))

        # Linha divisória sutil
        tk.Frame(center, bg=c["border2"], height=1, width=460).pack(pady=(0, 36))

        # Cards row
        cards = tk.Frame(center, bg=c["bg"])
        cards.pack()

        def _choice_card(parent, accent, icon, title, subtitle, action):
            NORMAL = c["card"]
            HOVER  = c["card2"]

            outer = tk.Frame(parent, bg=accent, padx=1, pady=1,
                             cursor="hand2")
            inner = tk.Frame(outer, bg=NORMAL, width=200, height=180)
            inner.pack_propagate(False)
            inner.pack()

            body = tk.Frame(inner, bg=NORMAL, padx=24, pady=24)
            body.pack(fill="both", expand=True)

            # Ícone
            ic_bg = accent + "22"  # transparência simulada com hex
            ic_frm = tk.Frame(body, bg=c["accentbg"], width=52, height=52)
            ic_frm.pack_propagate(False)
            ic_frm.pack(anchor="w", pady=(0, 14))
            tk.Label(ic_frm, text=icon, bg=c["accentbg"],
                     fg=accent, font=("Segoe UI", 22)).place(
                         relx=0.5, rely=0.5, anchor="center")

            lbl_title = tk.Label(body, text=title, bg=NORMAL,
                                 fg="#FFFFFF", font=("Segoe UI", 13, "bold"))
            lbl_title.pack(anchor="w")
            lbl_sub = tk.Label(body, text=subtitle, bg=NORMAL,
                               fg=c["muted"], font=("Segoe UI", 8),
                               wraplength=160, justify="left")
            lbl_sub.pack(anchor="w", pady=(4, 0))

            # Seta →
            arrow = tk.Label(body, text="→", bg=NORMAL,
                             fg=accent, font=("Segoe UI", 14))
            arrow.pack(anchor="e")

            def _all_children(w):
                yield w
                for ch in w.winfo_children():
                    yield from _all_children(ch)

            def on_enter(e):
                inner.configure(bg=HOVER)
                for w in _all_children(inner):
                    try:
                        cur = w.cget("bg")
                        if cur == NORMAL: w.configure(bg=HOVER)
                    except Exception: pass

            def on_leave(e):
                inner.configure(bg=NORMAL)
                for w in _all_children(inner):
                    try:
                        cur = w.cget("bg")
                        if cur == HOVER: w.configure(bg=NORMAL)
                    except Exception: pass

            def on_click(e): action()

            for w in _all_children(outer):
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)
                w.bind("<Button-1>", on_click)

            return outer

        _choice_card(cards,
                     accent=c["teal"],
                     icon="↓", title="Relatórios",
                     subtitle="Baixe relatórios de atividades,\naplicativos e produtividade.",
                     action=self._mk_show_reports_frame
                     ).pack(side="left", padx=(0, 20))

        _choice_card(cards,
                     accent=c["purple"],
                     icon="✦", title="Classificação",
                     subtitle="Classifique atividades não\natribuídas em massa.",
                     action=self._mk_show_classify_frame
                     ).pack(side="left", padx=(0, 20))

        _choice_card(cards,
                     accent="#F59E0B",
                     icon="⏰", title="Controle de\nHorário",
                     subtitle="Monitor em tempo real.\nAvisos fora do expediente.",
                     action=self._mk_show_horario_frame
                     ).pack(side="left")

        # Botão sair
        tk.Frame(center, bg=c["border"], height=1, width=460).pack(pady=(36, 16))
        tk.Label(center, text="Deseja usar outra conta?", bg=c["bg"],
                 fg=c["muted"], font=("Segoe UI", 8)).pack()
        tk.Button(center, text="Desconectar", bg=c["bg"],
                  fg=c["danger"], relief="flat",
                  font=("Segoe UI", 8, "bold"), cursor="hand2",
                  activebackground=c["bg"], activeforeground="#FF6B6B",
                  command=self._mk_logout).pack(pady=(4, 0))

        return frame

    # ── Frame Login ───────────────────────────────────────────────────────
    def _build_mk_login_frame(self, parent):
        c = self.COLORS
        frame = tk.Frame(parent, bg=c["bg"])

        # Coluna central — máx 480px
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=0)
        frame.columnconfigure(2, weight=1)
        frame.rowconfigure(0, weight=1)

        center = tk.Frame(frame, bg=c["bg"])
        center.grid(row=0, column=1, sticky="n", pady=52)

        # ── Brand (sem card branco — direto no fundo escuro) ───────────────
        brand_row = tk.Frame(center, bg=c["bg"])
        brand_row.pack(anchor="w", pady=(0, 28))
        tk.Frame(brand_row, bg=c["teal"], width=4, height=48).pack(side="left")
        brand_txt = tk.Frame(brand_row, bg=c["bg"])
        brand_txt.pack(side="left", padx=(14, 0))
        tk.Label(brand_txt, text="Makrolock", bg=c["bg"], fg=c["text"],
                 font=("Segoe UI", 22, "bold")).pack(anchor="w")
        tk.Label(brand_txt, text="Acesse sua conta para continuar",
                 bg=c["bg"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))

        # ── Card do formulário — fundo card escuro ─────────────────────────
        brand = tk.Frame(center, bg=c["card"],
                         highlightthickness=1, highlightbackground=c["border2"])
        brand.pack(fill="x")

        brand_inner = tk.Frame(brand, bg=c["card"], padx=36, pady=28)
        brand_inner.pack(fill="x")

        # ── Formulário ────────────────────────────────────────────────────
        form = tk.Frame(brand_inner, bg=c["card"])
        form.pack(fill="x")
        form.columnconfigure(0, weight=1)

        # Status badge (fundo danger escuro, texto claro)
        self._mk_badge = tk.Label(form, textvariable=self._mk_status,
                                  bg=c["card2"], fg=c["danger"],
                                  font=("Segoe UI", 8, "bold"), padx=10, pady=5,
                                  highlightthickness=1, highlightbackground="#3A1A1A")
        self._mk_badge.grid(row=0, column=0, sticky="ew", pady=(0, 18))

        tk.Label(form, text="E-MAIL", bg=c["card"], fg=c["muted2"],
                 font=("Segoe UI", 8, "bold")).grid(row=1, column=0, sticky="w", pady=(0, 5))
        email_entry = ttk.Entry(form, textvariable=self._mk_email, width=38)
        email_entry.grid(row=2, column=0, sticky="ew", pady=(0, 16))

        tk.Label(form, text="SENHA", bg=c["card"], fg=c["muted2"],
                 font=("Segoe UI", 8, "bold")).grid(row=3, column=0, sticky="w", pady=(0, 5))
        pass_entry = ttk.Entry(form, textvariable=self._mk_pass, width=38, show="•")
        pass_entry.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        pass_entry.bind("<Return>", lambda e: self._mk_login())

        tk.Checkbutton(form, text="Salvar senha", variable=self._mk_save_pass_var,
                       bg=c["card"], fg=c["muted"], selectcolor=c["card2"],
                       activebackground=c["card"], activeforeground=c["muted"],
                       font=("Segoe UI", 8), cursor="hand2").grid(row=5, column=0, sticky="w", pady=(0, 16))

        # Botão entrar — full width
        ttk.Button(form, text="  Entrar no Makrolock  ", style="Accent.TButton",
                   command=self._mk_login).grid(row=6, column=0, sticky="ew")

        # ── Configs avançadas — colapsável, dark ────────────────────────────
        cfg_toggle_frm = tk.Frame(brand_inner, bg=c["card"], pady=12)
        cfg_toggle_frm.pack(fill="x")
        cfg_lbl = tk.Label(cfg_toggle_frm, text="▸  Configurações avançadas",
                           bg=c["card"], fg=c["muted"],
                           font=("Segoe UI", 8), cursor="hand2")
        cfg_lbl.pack(anchor="w")

        cfg_body = tk.Frame(brand_inner, bg=c["card2"],
                            highlightthickness=1, highlightbackground=c["border"])
        self._mk_url_var = tk.StringVar(value=self._mk_base_url)

        def _toggle_cfg():
            if cfg_body.winfo_ismapped():
                cfg_body.pack_forget()
                cfg_lbl.configure(text="▸  Configurações avançadas")
            else:
                cfg_body.pack(fill="x", pady=(0, 8))
                cfg_lbl.configure(text="▾  Configurações avançadas")

        cfg_lbl.bind("<Button-1>", lambda e: _toggle_cfg())

        cfg_inner = tk.Frame(cfg_body, bg=c["card2"], padx=16, pady=12)
        cfg_inner.pack(fill="x")
        cfg_inner.columnconfigure(0, weight=1)
        tk.Label(cfg_inner, text="URL DO SISTEMA", bg=c["card2"], fg=c["muted2"],
                 font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Entry(cfg_inner, textvariable=self._mk_url_var).grid(row=1, column=0, sticky="ew")
        tk.Label(cfg_inner, text="Altere apenas se a URL do seu Makrolock for diferente.",
                 bg=c["card2"], fg=c["muted"],
                 font=("Segoe UI", 7)).grid(row=2, column=0, sticky="w", pady=(4, 0))

        # Log oculto — mantido apenas para compatibilidade com _mk_log()
        self._mk_log_box = tk.Text(frame, height=1, state="disabled",
                                   bg=c["bg"], fg=c["muted"],
                                   font=("Consolas", 8), relief="flat", bd=0)

        return frame

    # ── Frame Relatórios ──────────────────────────────────────────────────
    def _build_mk_reports_frame(self, parent):
        c = self.COLORS
        frame = tk.Frame(parent, bg=c["bg"])

        # Scroll
        canvas = tk.Canvas(frame, bg=c["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=c["bg"])
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))

        def _mw(ev): canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _mw))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # ── Topbar de navegação — dark ────────────────────────────────────
        topbar = tk.Frame(inner, bg=c["card"],
                          highlightthickness=1, highlightbackground=c["border2"])
        topbar.pack(fill="x", pady=(0, 20))
        nav_line = tk.Frame(topbar, height=3)
        nav_line.pack(fill="x")
        tk.Frame(nav_line, bg=c["teal"], height=3).pack(side="left", fill="y", expand=True)
        tk.Frame(nav_line, bg=c["accent"], height=3).pack(side="left", fill="y", expand=True)
        nav_inner = tk.Frame(topbar, bg=c["card"], padx=20, pady=12)
        nav_inner.pack(fill="x")

        # Faixa lateral colorida
        tk.Frame(nav_inner, bg=c["teal"], width=3, height=18).pack(side="left", padx=(0, 12))

        # Breadcrumb esquerda
        bc = tk.Frame(nav_inner, bg=c["card"])
        bc.pack(side="left")
        bc_mk = tk.Label(bc, text="Makrolock", bg=c["card"], fg=c["muted"],
                         font=("Segoe UI", 9), cursor="hand2")
        bc_mk.pack(side="left")
        bc_mk.bind("<Button-1>", lambda e: self._mk_show_choice_modal())
        tk.Label(bc, text=" › ", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Label(bc, text="Relatórios", bg=c["card"], fg=c["teal"],
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        # Botões direita
        btn_r = tk.Frame(nav_inner, bg=c["card"])
        btn_r.pack(side="right")
        ttk.Button(btn_r, text="⬅  Menu", style="Ghost.TButton",
                   command=self._mk_show_choice_modal).pack(side="left", padx=(0, 8))
        ttk.Button(btn_r, text="Desconectar", style="Danger.TButton",
                   command=self._mk_logout).pack(side="left")

        # Card relatórios
        outer_r, card_r = self._card(inner, accent=True, color=c["teal"])
        outer_r.pack(fill="x", pady=(0, 12))

        tk.Label(card_r, text="📥  BUSCAR RELATÓRIOS DO MAKROLOCK", bg=c["card"], fg=c["teal"],
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 8))

        date_row = ttk.Frame(card_r, style="Card.TFrame")
        date_row.pack(fill="x", pady=(0, 10))
        tk.Label(date_row, text="DE", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(side="left")
        _e_mk_ds = ttk.Entry(date_row, textvariable=self._mk_date_start, width=14)
        _e_mk_ds.pack(side="left", padx=(4, 0))
        tk.Label(date_row, text="ATÉ", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(side="left", padx=(12, 0))
        _e_mk_de = ttk.Entry(date_row, textvariable=self._mk_date_end, width=14)
        _e_mk_de.pack(side="left", padx=(4, 0))
        self._bind_date_mask(_e_mk_ds, self._mk_date_start, next_widget=_e_mk_de)
        self._bind_date_mask(_e_mk_de, self._mk_date_end)
        tk.Label(date_row, text="(DD/MM/AAAA)", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(side="left", padx=(8, 0))

        tk.Frame(card_r, bg=c["border"], height=1).pack(fill="x", pady=(0, 8))

        cols = ttk.Frame(card_r, style="Card.TFrame")
        cols.pack(fill="x", pady=(0, 8))
        cols.columnconfigure(0, weight=1)
        cols.columnconfigure(1, weight=1)

        left_col = tk.Frame(cols, bg=c["card"])
        left_col.grid(row=0, column=0, sticky="nw")
        tk.Label(left_col, text="TIPOS DE RELATÓRIO", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

        _MK_REPORT_TYPES = [
            ("activities",   "📋 Atividades"),
            ("apps",         "💻 Aplicativos"),
            ("internet",     "🌐 Internet"),
            ("devices",      "🖨 Dispositivos"),
            ("files",        "📁 Arquivos"),
            ("productivity", "📈 Produtividade"),
        ]
        self._mk_type_vars = {}
        self._mk_type_labels = {}
        for key, label in _MK_REPORT_TYPES:
            var = tk.BooleanVar(value=(key == "activities"))
            self._mk_type_vars[key] = var
            self._mk_type_labels[key] = label
            ttk.Checkbutton(left_col, text=label, variable=var).pack(anchor="w", pady=1)

        right_col = tk.Frame(cols, bg=c["card"])
        right_col.grid(row=0, column=1, sticky="nw", padx=(20, 0))
        tk.Label(right_col, text="GRUPOS / COMPUTADORES", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))
        self._mk_groups_frame = right_col
        self._mk_group_vars: dict[str, tk.BooleanVar] = {}
        self._mk_groups_lbl = tk.Label(right_col,
                                       text="Conecte-se para carregar os grupos.",
                                       bg=c["card"], fg=c["muted"], font=("Segoe UI", 9))
        self._mk_groups_lbl.pack(anchor="w")
        ttk.Button(right_col, text="🔄 Recarregar grupos", style="Ghost.TButton",
                   command=self._mk_load_groups).pack(anchor="w", pady=(6, 0))

        tk.Frame(card_r, bg=c["border"], height=1).pack(fill="x", pady=(0, 8))

        self._mk_prog_var = tk.DoubleVar(value=0)
        self._mk_prog_lbl = tk.StringVar(value="")
        self._mk_prog_bar = ttk.Progressbar(card_r, variable=self._mk_prog_var, maximum=100,
                                            style="Horizontal.TProgressbar")
        self._mk_prog_bar.pack(fill="x", pady=(0, 4))
        tk.Label(card_r, textvariable=self._mk_prog_lbl, bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 8))

        btn_row_r = ttk.Frame(card_r, style="Card.TFrame")
        btn_row_r.pack(anchor="w")
        self._mk_fetch_btn = ttk.Button(btn_row_r, text="📥  Baixar relatórios",
                                        style="Accent.TButton",
                                        command=self._mk_fetch_and_process)
        self._mk_fetch_btn.pack(side="left")
        tk.Label(btn_row_r,
                 text="  Os arquivos serão adicionados à fila — processe quando quiser na aba ⚙ Processar.",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 8)).pack(side="left")

        # Log visível
        tk.Frame(card_r, bg=c["border"], height=1).pack(fill="x", pady=(12, 6))
        log_hdr = tk.Frame(card_r, bg=c["card"])
        log_hdr.pack(fill="x")
        tk.Label(log_hdr, text="LOG", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 7, "bold")).pack(side="left")
        ttk.Button(log_hdr, text="Limpar", style="Ghost.TButton",
                   command=self._mk_log_clear).pack(side="right")

        log_frame = tk.Frame(card_r, bg=c["card2"])
        log_frame.pack(fill="both", expand=True, pady=(4, 0))
        self._mk_log_box = tk.Text(
            log_frame, height=8, state="disabled",
            bg=c["card2"], fg="#A8C8F0", font=("Consolas", 8),
            wrap="word", relief="flat", bd=0,
        )
        log_sb = ttk.Scrollbar(log_frame, orient="vertical", command=self._mk_log_box.yview)
        self._mk_log_box.configure(yscrollcommand=log_sb.set)
        log_sb.pack(side="right", fill="y")
        self._mk_log_box.pack(fill="both", expand=True, padx=6, pady=4)

        return frame

    # ── Frame Classificação ───────────────────────────────────────────────
    def _build_mk_classify_frame(self, parent):
        c = self.COLORS
        frame = tk.Frame(parent, bg=c["bg"])

        # Scroll
        canvas = tk.Canvas(frame, bg=c["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=c["bg"])
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))

        def _mw(ev): canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _mw))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # ── Topbar de navegação — dark ────────────────────────────────────
        topbar = tk.Frame(inner, bg=c["card"],
                          highlightthickness=1, highlightbackground=c["border2"])
        topbar.pack(fill="x", pady=(0, 20))
        nav_line2 = tk.Frame(topbar, height=3)
        nav_line2.pack(fill="x")
        tk.Frame(nav_line2, bg=c["purple"], height=3).pack(side="left", fill="y", expand=True)
        tk.Frame(nav_line2, bg=c["accent"],  height=3).pack(side="left", fill="y", expand=True)
        nav_inner = tk.Frame(topbar, bg=c["card"], padx=20, pady=12)
        nav_inner.pack(fill="x")

        tk.Frame(nav_inner, bg=c["purple"], width=3, height=18).pack(side="left", padx=(0, 12))

        bc = tk.Frame(nav_inner, bg=c["card"])
        bc.pack(side="left")
        bc_mk = tk.Label(bc, text="Makrolock", bg=c["card"], fg=c["muted"],
                         font=("Segoe UI", 9), cursor="hand2")
        bc_mk.pack(side="left")
        bc_mk.bind("<Button-1>", lambda e: self._mk_show_choice_modal())
        tk.Label(bc, text=" › ", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Label(bc, text="Classificação", bg=c["card"], fg=c["purple"],
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        btn_r = tk.Frame(nav_inner, bg=c["card"])
        btn_r.pack(side="right")
        ttk.Button(btn_r, text="⬅  Menu", style="Ghost.TButton",
                   command=self._mk_show_choice_modal).pack(side="left", padx=(0, 8))
        ttk.Button(btn_r, text="Desconectar", style="Danger.TButton",
                   command=self._mk_logout).pack(side="left")

        # Card classificação
        outer_cls, card_cls = self._card(inner, accent=True, color="#8B5CF6")
        outer_cls.pack(fill="x", pady=(0, 12))

        tk.Label(card_cls, text="🏷️  CLASSIFICAÇÃO EM MASSA (Produtividade)", bg=c["card"], fg="#8B5CF6",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 4))

        self._mk_cls_status = tk.StringVar(value="Clique em 'Buscar' para carregar os itens não classificados.")
        tk.Label(card_cls, textvariable=self._mk_cls_status, bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 8))

        filter_row = ttk.Frame(card_cls, style="Card.TFrame")
        filter_row.pack(fill="x", pady=(0, 4))
        filter_row.columnconfigure(1, weight=1)
        tk.Label(filter_row, text="🔎 Filtro:", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._mk_cls_filter_var = tk.StringVar()
        filter_entry = ttk.Entry(filter_row, textvariable=self._mk_cls_filter_var)
        filter_entry.grid(row=0, column=1, sticky="ew")
        self._mk_cls_filter_var.trace_add("write", lambda *_: self._mk_cls_apply_filter())

        _QUICK_FILTERS = [
            ("🏛 Governo",   ["gov.br", "jus.br", "leg.br", "mp.br", "mil.br", "def.br",
                               "tcu", "tst", "trf", "tjsp", "tjmg", "tjrj", "tre", "tse",
                               "receita", "serpro", "dataprev", "inss", "esocial", "cnes",
                               "sefaz", "detran", "denatran", "anatel", "anvisa", "ibge"]),
            ("🏦 Banco",     ["banco", "bradesco", "itau", "santander", "bb.com.br", "caixa",
                               "cef.gov", "nubank", "inter.", "sicoob", "sicredi", "banrisul",
                               "btg", "xp.com", "picpay", "mercadopago", "pagbank", "c6bank"]),
            ("💬 Social",    ["facebook", "instagram", "twitter", "x.com", "tiktok",
                               "linkedin", "whatsapp", "telegram", "discord", "reddit",
                               "snapchat", "pinterest", "kwai"]),
            ("▶ Streaming", ["youtube", "netflix", "spotify", "twitch", "prime",
                               "disney", "hbomax", "globoplay", "deezer", "soundcloud",
                               "crunchyroll", "paramount", "apple.tv"]),
            ("🎮 Jogos",     ["steam", "epicgames", "roblox", "battlenet", "riot",
                               "minecraft", "ubisoft", "ea.com", "origin", "gog.com"]),
            ("☁ Nuvem",     ["drive.google", "onedrive", "dropbox", "sharepoint",
                               "box.com", "mega.nz", "icloud"]),
        ]
        self._mk_quick_filters = _QUICK_FILTERS

        qf_row = ttk.Frame(card_cls, style="Card.TFrame")
        qf_row.pack(anchor="w", pady=(0, 8))
        for label, _ in _QUICK_FILTERS:
            ttk.Button(qf_row, text=label, style="Ghost.TButton",
                       command=lambda lbl=label: self._mk_cls_quick_filter(lbl)).pack(side="left", padx=(0, 4))
        ttk.Button(qf_row, text="✖ Limpar filtro", style="Ghost.TButton",
                   command=self._mk_cls_clear_filter).pack(side="left", padx=(4, 0))

        lb_cls_wrap = tk.Frame(card_cls, bg=c["card"],
                               highlightthickness=1, highlightbackground=c["border"])
        lb_cls_wrap.pack(fill="x", pady=(0, 6))
        sb_cls = ttk.Scrollbar(lb_cls_wrap, orient="vertical")
        self._mk_cls_lb = tk.Listbox(
            lb_cls_wrap, height=10, selectmode="multiple",
            bg=c["bg"], fg=c["text"], relief="flat", bd=0,
            highlightthickness=0,
            font=("Segoe UI", 9), yscrollcommand=sb_cls.set,
            exportselection=False,
            activestyle="dotbox",
        )
        sb_cls.config(command=self._mk_cls_lb.yview)
        self._mk_cls_lb.pack(side="left", fill="x", expand=True)
        sb_cls.pack(side="right", fill="y")

        def _lb_scroll(event):
            self._mk_cls_lb.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
        self._mk_cls_lb.bind("<MouseWheel>", _lb_scroll)
        self._mk_cls_lb.bind("<Enter>", lambda e: self._mk_cls_lb.bind_all("<MouseWheel>", _lb_scroll))
        self._mk_cls_lb.bind("<Leave>", lambda e: self._mk_cls_lb.unbind_all("<MouseWheel>"))

        sel_row = ttk.Frame(card_cls, style="Card.TFrame")
        sel_row.pack(fill="x", pady=(0, 6))
        ttk.Button(sel_row, text="Sel. todos", style="Ghost.TButton",
                   command=lambda: self._mk_cls_lb.select_set(0, "end")).pack(side="left")
        ttk.Button(sel_row, text="Desselecionar", style="Ghost.TButton",
                   command=lambda: self._mk_cls_lb.selection_clear(0, "end")).pack(side="left", padx=(6, 0))
        self._mk_cls_count_lbl = tk.StringVar(value="")
        tk.Label(sel_row, textvariable=self._mk_cls_count_lbl, bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(side="right")

        cls_btn_row = ttk.Frame(card_cls, style="Card.TFrame")
        cls_btn_row.pack(anchor="w", pady=(0, 6))
        self._mk_cls_btn_load = ttk.Button(cls_btn_row, text="🔍  Buscar não classificados",
                                           style="Accent.TButton", command=self._mk_classify_load)
        self._mk_cls_btn_load.pack(side="left")
        self._mk_cls_btn_auto = ttk.Button(cls_btn_row, text="⚡ Auto-classificar",
                                           style="Ghost.TButton", command=self._mk_auto_classify)
        self._mk_cls_btn_auto.pack(side="left", padx=(8, 0))
        tk.Label(cls_btn_row, text="  Marcar selecionados:", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(side="left", padx=(12, 4))
        self._mk_cls_btn_prod = ttk.Button(cls_btn_row, text="✅ Produtivo",
                                           style="Accent.TButton",
                                           command=lambda: self._mk_classify_apply("productive"))
        self._mk_cls_btn_prod.pack(side="left", padx=(0, 4))
        self._mk_cls_btn_unprod = ttk.Button(cls_btn_row, text="❌ Não Produtivo",
                                             style="Ghost.TButton",
                                             command=lambda: self._mk_classify_apply("unproductive"))
        self._mk_cls_btn_unprod.pack(side="left", padx=(0, 4))
        self._mk_cls_btn_neutral = ttk.Button(cls_btn_row, text="➖ Neutro",
                                              style="Ghost.TButton",
                                              command=lambda: self._mk_classify_apply("neutral"))
        self._mk_cls_btn_neutral.pack(side="left")

        self._mk_cls_prog_var = tk.DoubleVar(value=0)
        self._mk_cls_prog_lbl = tk.StringVar(value="")
        self._mk_cls_prog_bar = ttk.Progressbar(card_cls, variable=self._mk_cls_prog_var, maximum=100)
        self._mk_cls_prog_bar.pack(fill="x", pady=(6, 2))
        tk.Label(card_cls, textvariable=self._mk_cls_prog_lbl, bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(anchor="w")

        return frame

    # ── Frame Controle de Horário ─────────────────────────────────────────
    def _build_mk_horario_frame(self, parent):
        c = self.COLORS
        AMBER = "#F59E0B"
        frame = tk.Frame(parent, bg=c["bg"])

        # Scroll
        canvas = tk.Canvas(frame, bg=c["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=c["bg"])
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_resize(e):
            canvas.itemconfig(win_id, width=e.width)
        def _on_frame(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.bind("<Configure>", _on_resize)
        inner.bind("<Configure>", _on_frame)

        def _scroll(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _scroll)

        pad = tk.Frame(inner, bg=c["bg"], padx=32, pady=24)
        pad.pack(fill="both", expand=True)

        # Navegação
        nav = tk.Frame(pad, bg=c["bg"])
        nav.pack(fill="x", pady=(0, 16))
        ttk.Button(nav, text="⬅  Menu", style="Ghost.TButton",
                   command=self._mk_show_choice_modal).pack(side="left")
        ttk.Button(nav, text="Desconectar", style="Danger.TButton",
                   command=self._mk_logout).pack(side="right")

        # Título
        tk.Label(pad, text="⏰  CONTROLE DE HORÁRIO", bg=c["bg"], fg=AMBER,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(pad, text="Monitor em tempo real · Avisos automáticos fora do expediente",
                 bg=c["bg"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 16))

        # ── Card: Configurar expediente ──────────────────────────────────
        _, card_cfg = self._card(pad, accent=True, color=AMBER)
        _.pack(fill="x", pady=(0, 12))

        tk.Label(card_cfg, text="EXPEDIENTE", bg=c["card"], fg=AMBER,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 10))

        dias_row = tk.Frame(card_cfg, bg=c["card"])
        dias_row.pack(anchor="w", pady=(0, 8))
        tk.Label(dias_row, text="Dias:", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))

        self._mk_hor_dias = {}
        dias = [("Seg", "seg"), ("Ter", "ter"), ("Qua", "qua"),
                ("Qui", "qui"), ("Sex", "sex"), ("Sáb", "sab"), ("Dom", "dom")]
        for label, key in dias:
            var = tk.BooleanVar(value=(key not in ("sab", "dom")))
            self._mk_hor_dias[key] = var
            ttk.Checkbutton(dias_row, text=label, variable=var).pack(side="left", padx=2)

        hora_row = tk.Frame(card_cfg, bg=c["card"])
        hora_row.pack(anchor="w", pady=(0, 8))
        tk.Label(hora_row, text="Horário:", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        self._mk_hor_inicio = tk.StringVar(value="08:00")
        self._mk_hor_fim    = tk.StringVar(value="18:00")
        ttk.Entry(hora_row, textvariable=self._mk_hor_inicio, width=8).pack(side="left")
        tk.Label(hora_row, text="até", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=6)
        ttk.Entry(hora_row, textvariable=self._mk_hor_fim, width=8).pack(side="left")

        acao_row = tk.Frame(card_cfg, bg=c["card"])
        acao_row.pack(anchor="w", pady=(0, 8))
        tk.Label(acao_row, text="Ação:", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        self._mk_hor_acao = tk.StringVar(value="aviso")
        ttk.Radiobutton(acao_row, text="📢 Enviar aviso", variable=self._mk_hor_acao,
                        value="aviso").pack(side="left", padx=4)
        ttk.Radiobutton(acao_row, text="📝 Só registrar", variable=self._mk_hor_acao,
                        value="log").pack(side="left", padx=4)

        msg_row = tk.Frame(card_cfg, bg=c["card"])
        msg_row.pack(fill="x", pady=(0, 8))
        tk.Label(msg_row, text="Mensagem:", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(anchor="w")
        self._mk_hor_msg = tk.StringVar(
            value="Atenção: seu computador está ativo fora do horário de trabalho.")
        ttk.Entry(msg_row, textvariable=self._mk_hor_msg, width=60).pack(fill="x", pady=(4, 0))

        intervalo_row = tk.Frame(card_cfg, bg=c["card"])
        intervalo_row.pack(anchor="w")
        tk.Label(intervalo_row, text="Verificar a cada:", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        self._mk_hor_intervalo = tk.StringVar(value="5")
        ttk.Entry(intervalo_row, textvariable=self._mk_hor_intervalo, width=5).pack(side="left")
        tk.Label(intervalo_row, text="minutos", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(4, 0))

        # ── Card: Monitor em tempo real ──────────────────────────────────
        _, card_mon = self._card(pad, accent=True, color=c["teal"])
        _.pack(fill="x", pady=(0, 12))

        mon_hdr = tk.Frame(card_mon, bg=c["card"])
        mon_hdr.pack(fill="x", pady=(0, 10))
        tk.Label(mon_hdr, text="MONITOR EM TEMPO REAL", bg=c["card"], fg=c["teal"],
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        self._mk_hor_status_lbl = tk.Label(mon_hdr, text="● Parado",
                                            bg=c["card"], fg=c["muted"],
                                            font=("Segoe UI", 9))
        self._mk_hor_status_lbl.pack(side="right")

        btn_row = tk.Frame(card_mon, bg=c["card"])
        btn_row.pack(anchor="w", pady=(0, 10))
        self._mk_hor_btn_iniciar = ttk.Button(btn_row, text="▶ Iniciar monitor",
                                               style="Accent.TButton",
                                               command=self._mk_hor_toggle)
        self._mk_hor_btn_iniciar.pack(side="left", padx=(0, 8))
        ttk.Button(btn_row, text="🔄 Verificar agora",
                   style="Ghost.TButton",
                   command=self._mk_hor_verificar_agora).pack(side="left")

        # Tabela de estações
        cols_def = [("Usuário (Computador)", 220), ("Status", 90),
                    ("Online", 70), ("Situação", 110), ("Ação", 100)]
        tbl_frame = tk.Frame(card_mon, bg=c["card"])
        tbl_frame.pack(fill="x")

        self._mk_hor_tree = ttk.Treeview(tbl_frame,
                                          columns=[d[0] for d in cols_def],
                                          show="headings", height=12)
        for col, width in cols_def:
            self._mk_hor_tree.heading(col, text=col)
            self._mk_hor_tree.column(col, width=width, minwidth=60)

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical",
                             command=self._mk_hor_tree.yview)
        self._mk_hor_tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._mk_hor_tree.pack(fill="x")

        # Tags de cor para a tabela
        self._mk_hor_tree.tag_configure("fora",   foreground="#F87171")
        self._mk_hor_tree.tag_configure("dentro", foreground="#34D399")
        self._mk_hor_tree.tag_configure("off",    foreground=c["muted"])

        # Duplo clique → enviar mensagem manual
        self._mk_hor_tree.bind("<Double-1>", self._mk_hor_enviar_manual)

        # ── Card: Log de ocorrências ─────────────────────────────────────
        _, card_log = self._card(pad, accent=False)
        _.pack(fill="x", pady=(0, 12))

        log_hdr = tk.Frame(card_log, bg=c["card"])
        log_hdr.pack(fill="x", pady=(0, 6))
        tk.Label(log_hdr, text="LOG DE OCORRÊNCIAS", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Button(log_hdr, text="Limpar", style="Ghost.TButton",
                   command=lambda: self._mk_hor_log_txt.configure(state="normal") or
                   self._mk_hor_log_txt.delete("1.0", "end") or
                   self._mk_hor_log_txt.configure(state="disabled")
                   ).pack(side="right")

        self._mk_hor_log_txt = tk.Text(card_log, height=8, bg=c["card2"],
                                        fg=c["text"], font=("Consolas", 8),
                                        relief="flat", state="disabled",
                                        wrap="word")
        self._mk_hor_log_txt.pack(fill="x")

        return frame

    def _mk_hor_log(self, msg: str):
        import datetime as _dt
        ts = _dt.datetime.now().strftime("%H:%M:%S")
        txt = self._mk_hor_log_txt
        txt.configure(state="normal")
        txt.insert("end", f"[{ts}] {msg}\n")
        txt.see("end")
        txt.configure(state="disabled")

    def _mk_hor_toggle(self):
        if self._mk_horario_monitor_ativo:
            self._mk_horario_monitor_ativo = False
            if self._mk_horario_after_id:
                try:
                    self.root.after_cancel(self._mk_horario_after_id)
                except Exception:
                    pass
                self._mk_horario_after_id = None
            self._mk_hor_btn_iniciar.configure(text="▶ Iniciar monitor")
            self._mk_hor_status_lbl.configure(text="● Parado", fg=self.COLORS["muted"])
            self._mk_hor_log("Monitor parado.")
        else:
            self._mk_horario_monitor_ativo = True
            self._mk_hor_btn_iniciar.configure(text="⏹ Parar monitor")
            self._mk_hor_status_lbl.configure(text="● Ativo", fg="#34D399")
            self._mk_hor_log("Monitor iniciado.")
            self._mk_hor_ciclo()

    def _mk_hor_verificar_agora(self):
        import threading
        threading.Thread(target=self._mk_hor_do_check, daemon=True).start()

    def _mk_hor_ciclo(self):
        if not self._mk_horario_monitor_ativo:
            return
        import threading
        threading.Thread(target=self._mk_hor_do_check, daemon=True).start()
        try:
            mins = max(1, int(self._mk_hor_intervalo.get()))
        except ValueError:
            mins = 5
        self._mk_horario_after_id = self.root.after(mins * 60_000, self._mk_hor_ciclo)

    def _mk_hor_do_check(self):
        import datetime as _dt
        if not self._mk_logged or not self._mk_session:
            self.root.after(0, lambda: self._mk_hor_log("⚠ Não conectado ao Makrolock."))
            return
        try:
            rc = self._mk_session.get(f"{self._mk_base_url}/api/front/computers", timeout=15)
            if rc.status_code != 200:
                self.root.after(0, lambda: self._mk_hor_log(f"⚠ Erro ao buscar computadores: HTTP {rc.status_code}"))
                return
            cdata = rc.json()
            clist = cdata if isinstance(cdata, list) else (
                cdata.get("computers") or cdata.get("data") or [])
        except Exception as ex:
            self.root.after(0, lambda: self._mk_hor_log(f"⚠ Erro de rede: {ex}"))
            return

        now = _dt.datetime.now()
        dow_map = {0: "seg", 1: "ter", 2: "qua", 3: "qui", 4: "sex", 5: "sab", 6: "dom"}
        dia_key = dow_map[now.weekday()]
        dia_ok  = self._mk_hor_dias.get(dia_key, tk.BooleanVar(value=True)).get()

        try:
            h_ini = _dt.time.fromisoformat(self._mk_hor_inicio.get())
            h_fim = _dt.time.fromisoformat(self._mk_hor_fim.get())
        except ValueError:
            h_ini, h_fim = _dt.time(8, 0), _dt.time(18, 0)

        hora_atual = now.time().replace(second=0, microsecond=0)
        dentro_horario = dia_ok and h_ini <= hora_atual <= h_fim

        rows_data = []
        infratores = []
        for comp in clist:
            cname   = comp.get("name") or comp.get("alias") or str(comp.get("id", "?"))
            uname   = (comp.get("currentUser") or comp.get("user") or
                       comp.get("username") or "")
            online  = bool(comp.get("online", False))
            display = f"{uname} ({cname})" if uname else cname
            if not online:
                situacao = "Offline"
                tag = "off"
            elif dentro_horario:
                situacao = "✅ No expediente"
                tag = "dentro"
            else:
                situacao = "⚠ Fora do expediente"
                tag = "fora"
                if online:
                    infratores.append((cname, uname, display))

            rows_data.append((display, "Online" if online else "Offline",
                              "Sim" if online else "Não", situacao, tag, cname))

        def _update_ui():
            tree = self._mk_hor_tree
            tree.delete(*tree.get_children())
            for display, status, is_online, situacao, tag, cname in rows_data:
                tree.insert("", "end",
                            values=(display, status, is_online, situacao, "✉ Enviar aviso"),
                            tags=(tag,), iid=cname)
            if infratores:
                self._mk_hor_log(
                    f"{'⚠' if not dentro_horario else '✅'} "
                    f"{len(infratores)} estação(ões) ativa(s) fora do expediente: "
                    + ", ".join(d for _, _, d in infratores))
                if self._mk_hor_acao.get() == "aviso":
                    for cname, uname, display in infratores:
                        self._mk_hor_enviar_aviso(cname, uname)
            else:
                self._mk_hor_log(
                    f"✅ {'Dentro do expediente' if dentro_horario else 'Fora do expediente'} "
                    f"— nenhuma estação irregular.")

        self.root.after(0, _update_ui)

    def _mk_hor_enviar_aviso(self, cname: str, uname: str = ""):
        """Envia mensagem de aviso para uma estação via API Makrolock."""
        msg = self._mk_hor_msg.get().strip()
        if not msg:
            return
        self._mk_hor_log(f"📢 Enviando aviso para {cname}...")
        import threading
        threading.Thread(target=self._mk_hor_do_enviar,
                         args=(cname, msg), daemon=True).start()

    def _mk_hor_do_enviar(self, cname: str, msg: str):
        # TODO: substituir pela URL/método real capturado via F12 no portal Makrolock
        # Exemplo provável: POST /api/front/computers/{id}/message  body: {"message": msg}
        # ou: POST /api/front/message  body: {"computer": cname, "message": msg}
        try:
            # Tenta descobrir o ID numérico pelo nome
            cid = None
            for numeric_id, name in self._mk_id_to_name.items():
                if name == cname:
                    cid = numeric_id
                    break

            # ── Endpoint a confirmar ──────────────────────────────────────
            # Descomente e ajuste após capturar via F12:
            # url = f"{self._mk_base_url}/api/front/computers/{cid}/message"
            # payload = {"message": msg, "type": "alert"}
            # rc = self._mk_session.post(url, json=payload, timeout=10)
            # ok = rc.status_code in (200, 201, 204)
            # ──────────────────────────────────────────────────────────────

            # Por enquanto, simula o envio e loga que precisa do endpoint
            ok = False
            self.root.after(0, lambda: self._mk_hor_log(
                f"⏳ [{cname}] Endpoint de mensagem pendente — capture via F12 no portal."))
        except Exception as ex:
            self.root.after(0, lambda: self._mk_hor_log(f"❌ Erro ao enviar para {cname}: {ex}"))

    def _mk_hor_enviar_manual(self, event):
        """Duplo clique na tabela → envia mensagem manual para aquela estação."""
        tree = self._mk_hor_tree
        sel = tree.selection()
        if not sel:
            return
        cname = sel[0]  # iid é o cname
        msg = self._mk_hor_msg.get().strip()
        if not msg:
            return
        self._mk_hor_log(f"📢 Envio manual para {cname}...")
        import threading
        threading.Thread(target=self._mk_hor_do_enviar,
                         args=(cname, msg), daemon=True).start()

    # ── Métodos de integração Makrolock ───────────────────────────────────
    def _mk_load_config(self) -> dict:
        import json
        try:
            if self._mk_config_file.exists():
                return json.loads(self._mk_config_file.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}

    def _mk_load_email(self) -> str:
        return self._mk_load_config().get("email", "")

    def _mk_load_saved_pass(self) -> str:
        import base64
        try:
            v = self._mk_load_config().get("pass_b64", "")
            return base64.b64decode(v.encode()).decode() if v else ""
        except Exception:
            return ""

    def _mk_save_config(self, email: str, password: str = "", save_pass: bool = False):
        import json, base64
        data = self._mk_load_config()
        data["email"] = email
        if save_pass and password:
            data["pass_b64"] = base64.b64encode(password.encode()).decode()
            data["save_pass"] = True
        else:
            data.pop("pass_b64", None)
            data["save_pass"] = False
        try:
            self._mk_config_file.write_text(
                json.dumps(data, ensure_ascii=False), encoding="utf-8"
            )
        except Exception:
            pass

    def _mk_save_email(self, email: str):
        self._mk_save_config(email, save_pass=False)

    def _mk_log(self, msg: str, level: str = "info"):
        """Escreve no log da aba Makrolock. level: 'info' | 'ok' | 'warn' | 'error'"""
        from datetime import datetime as _dt
        colors = {"info": "#A8C8F0", "ok": "#6EE7B7", "warn": "#FDE68A", "error": "#FCA5A5"}
        tag = level
        ts = _dt.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        try:
            box = self._mk_log_box
            box.configure(state="normal")
            box.tag_configure(tag, foreground=colors.get(level, "#A8C8F0"))
            box.insert("end", line, tag)
            box.see("end")
            box.configure(state="disabled")
            self.root.update_idletasks()
        except Exception:
            pass
        self.write_log(f"[Makrolock] {msg}")

    def _mk_log_clear(self):
        try:
            self._mk_log_box.configure(state="normal")
            self._mk_log_box.delete("1.0", "end")
            self._mk_log_box.configure(state="disabled")
        except Exception:
            pass

    def _mk_login(self):
        threading.Thread(target=self._mk_do_login, daemon=True).start()

    def _mk_do_login(self):
        try:
            import requests as _req
        except ImportError:
            messagebox.showerror("Dependência ausente",
                "Instale o requests:\npython -m pip install requests")
            return

        email = self._mk_email.get().strip()
        passwd = self._mk_pass.get().strip()
        if not email or not passwd:
            messagebox.showwarning("Login", "Preencha e-mail e senha.")
            return

        self._mk_status.set("Conectando...")
        self._mk_badge.configure(bg=self.COLORS["warning"])
        self.root.update_idletasks()

        import hashlib
        base = self._mk_url_var.get().rstrip("/")
        # A API do Makrolock exige senha em MD5
        passwd_md5 = hashlib.md5(passwd.encode("utf-8")).hexdigest()
        self._mk_log(f"POST {base}/api/auth/login  (user: {email})")
        self._mk_log(f"Senha MD5: {passwd_md5[:8]}...{passwd_md5[-4:]}")
        try:
            sess = _req.Session()
            sess.headers.update({
                "Accept":       "application/json",
                "Content-Type": "application/json",
                "Origin":       base,
                "Referer":      f"{base}/",
                "User-Agent":   "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            })

            # Busca geo (IP público) para incluir no payload — opcional, mas o browser envia
            geo = {}
            try:
                geo_resp = _req.get("https://ipapi.co/json/", timeout=5)
                if geo_resp.status_code == 200:
                    gd = geo_resp.json()
                    geo = {
                        "ip":           gd.get("ip", ""),
                        "city":         gd.get("city", ""),
                        "country":      gd.get("country_name", ""),
                        "country_code": gd.get("country_code", ""),
                    }
                    self._mk_log(f"Geo: {geo.get('city')}, {geo.get('country')}")
            except Exception:
                self._mk_log("Geo não obtido (sem internet pública) — continuando sem geo.")

            payload = {"email": email, "password": passwd_md5}
            if geo:
                payload["geo"] = geo

            resp = sess.post(f"{base}/api/auth/login", json=payload, timeout=15)
            self._mk_log(f"HTTP {resp.status_code}  |  {len(resp.content)} bytes")

            if resp.status_code == 200:
                # Auth via cookie auth_token (set-cookie automático no Session)
                cookie = sess.cookies.get("auth_token")
                self._mk_log(f"Cookie auth_token: {cookie[:12]}...{cookie[-6:] if cookie else 'N/A'}", "ok")
                self._mk_session = sess
                self._mk_base_url = base
                self._mk_logged = True
                self._mk_pass_md5 = passwd_md5
                self._mk_save_config(email, self._mk_pass.get().strip(), self._mk_save_pass_var.get())
                self._mk_status.set(f"✅  Conectado como {email}")
                self._mk_badge.configure(bg=self.COLORS["ok"])
                self._mk_log(f"Login bem-sucedido!", "ok")
                # Carrega grupos automaticamente e abre modal de escolha
                threading.Thread(target=self._mk_do_load_groups, daemon=True).start()
                self.root.after(0, self._mk_show_choice_modal)
            else:
                self._mk_status.set("❌  Login falhou")
                self._mk_badge.configure(bg=self.COLORS["danger"])
                self._mk_log(f"Falha: HTTP {resp.status_code}  corpo: {resp.text[:400]}", "error")
                messagebox.showerror("Login",
                    f"Login falhou (HTTP {resp.status_code}).\n\nDetalhe:\n{resp.text[:300]}")
        except Exception as e:
            self._mk_status.set("❌  Erro de conexão")
            self._mk_badge.configure(bg=self.COLORS["danger"])
            self._mk_log(f"Exceção: {e}", "error")
            messagebox.showerror("Erro de conexão", str(e))

    def _mk_load_groups(self):
        """Busca grupos de computadores da API e popula os checkboxes."""
        if not self._mk_logged or not self._mk_session:
            messagebox.showwarning("Makrolock", "Conecte-se primeiro.")
            return
        threading.Thread(target=self._mk_do_load_groups, daemon=True).start()

    def _mk_do_load_groups(self):
        try:
            resp = self._mk_session.get(f"{self._mk_base_url}/api/front/computergroups", timeout=10)
            self._mk_log(f"Grupos: HTTP {resp.status_code}")
            if resp.status_code != 200:
                self._mk_log(f"Grupos: erro {resp.text[:200]}", "error")
                return
            data = resp.json()
            groups = data if isinstance(data, list) else data.get("groups", data.get("data", []))
            self._mk_log(f"Grupos recebidos: {len(groups)}  raw: {str(groups)[:300]}")

            # Busca lista de computadores para mapear ID numérico → nome real
            id_to_name: dict[str, str] = {}
            cname_to_display: dict[str, str] = {}  # nome real → label para UI
            rc = self._mk_session.get(f"{self._mk_base_url}/api/front/computers", timeout=10)
            self._mk_log(f"computers: HTTP {rc.status_code}  {rc.headers.get('Content-Type','?')}")
            if rc.status_code == 200:
                try:
                    cdata = rc.json()
                    clist = cdata if isinstance(cdata, list) else (
                        cdata.get("data") or cdata.get("computers") or
                        cdata.get("items") or cdata.get("results") or []
                    )
                    for comp in (clist if isinstance(clist, list) else []):
                        cid = str(comp.get("id", comp.get("_id", "")))
                        cname = (comp.get("name") or comp.get("alias") or
                                 comp.get("hostname") or comp.get("label") or cid)
                        # Tenta pegar nome do usuário logado
                        uname = (comp.get("currentUser") or comp.get("user") or
                                 comp.get("username") or comp.get("lastUser") or
                                 comp.get("loggedUser") or comp.get("userLogin") or "")
                        display = f"{uname} ({cname})" if uname else cname
                        if cid:
                            id_to_name[cid] = cname       # nome real para API
                            cname_to_display[cname] = display  # label para UI
                    self._mk_log(f"Computadores mapeados: {len(id_to_name)}  ex={dict(list(id_to_name.items())[:3])}")
                except Exception as ex:
                    self._mk_log(f"computers parse erro: {ex}  raw={rc.text[:300]}", "warn")
            self._mk_id_to_name = id_to_name  # guarda para uso no fetch

            # Atualiza UI na thread principal
            def _update():
                c = self.COLORS
                for w in list(self._mk_groups_frame.winfo_children()):
                    w.destroy()
                self._mk_group_vars.clear()
                self._mk_computer_vars.clear()

                tk.Label(self._mk_groups_frame, text="GRUPOS / COMPUTADORES", bg=c["card"], fg=c["muted"],
                         font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

                if not groups:
                    tk.Label(self._mk_groups_frame, text="Nenhum grupo encontrado.",
                             bg=c["card"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w")
                else:
                    for g in groups:
                        gid  = str(g.get("id", g.get("_id", "")))
                        gname = g.get("name") or g.get("alias") or g.get("label") or gid
                        raw_ids = [str(x) for x in (g.get("computerIds") or [])]
                        named = [id_to_name.get(cid, cid) for cid in raw_ids]
                        self._mk_group_computers[gid] = named
                        self._mk_log(f"  {gname}: {named}")

                        # Variável de grupo (Selecionar todos do grupo)
                        g_var = tk.BooleanVar(value=True)
                        self._mk_group_vars[gid] = (g_var, gname)

                        # Frame do grupo (cabeçalho clicável + computadores)
                        g_frame = tk.Frame(self._mk_groups_frame, bg=c["card"])
                        g_frame.pack(fill="x", pady=(4, 0))

                        expanded = tk.BooleanVar(value=True)
                        children_frame = tk.Frame(g_frame, bg=c["card2"])

                        def _toggle(ef=children_frame, ev=expanded):
                            if ev.get():
                                ef.pack_forget()
                                ev.set(False)
                            else:
                                ef.pack(fill="x", padx=(14, 0))
                                ev.set(True)

                        # Cabeçalho do grupo
                        hdr = tk.Frame(g_frame, bg=c["card"])
                        hdr.pack(fill="x")
                        arrow_lbl = tk.Label(hdr, text="▼", bg=c["card"], fg=c["accent"],
                                             font=("Segoe UI", 8), cursor="hand2")
                        arrow_lbl.pack(side="left")

                        def _toggle_with_arrow(ef=children_frame, ev=expanded, al=arrow_lbl):
                            if ev.get():
                                ef.pack_forget()
                                ev.set(False)
                                al.configure(text="▶")
                            else:
                                ef.pack(fill="x", padx=(14, 0))
                                ev.set(True)
                                al.configure(text="▼")

                        arrow_lbl.bind("<Button-1>", lambda e, f=_toggle_with_arrow: f())

                        # Checkbox do grupo (seleciona/desmarca todos filhos)
                        def _on_group_toggle(gv=g_var, named_=named):
                            val = gv.get()
                            for cn in named_:
                                if cn in self._mk_computer_vars:
                                    self._mk_computer_vars[cn].set(val)

                        ttk.Checkbutton(hdr, text=gname, variable=g_var,
                                        command=lambda gv=g_var, n=named: _on_group_toggle(gv, n)
                                        ).pack(side="left", padx=(2, 0))

                        # Frame de computadores (expandido por padrão)
                        children_frame.pack(fill="x", padx=(14, 0))

                        for cname in named:
                            cv = tk.BooleanVar(value=True)
                            self._mk_computer_vars[cname] = cv

                            def _on_comp_toggle(gv=g_var, named_=named):
                                all_on = all(self._mk_computer_vars[cn].get() for cn in named_ if cn in self._mk_computer_vars)
                                gv.set(all_on)

                            clabel = cname_to_display.get(cname, cname)
                            ttk.Checkbutton(children_frame, text=clabel, variable=cv,
                                            command=lambda gv=g_var, n=named: _on_comp_toggle(gv, n)
                                            ).pack(anchor="w", pady=1)

                ttk.Button(self._mk_groups_frame, text="🔄 Recarregar grupos",
                           style="Ghost.TButton",
                           command=self._mk_load_groups).pack(anchor="w", pady=(8, 0))
                self._mk_log(f"{len(groups)} grupo(s) carregados.", "ok")

            self.root.after(0, _update)
        except Exception as e:
            self._mk_log(f"Erro ao carregar grupos: {e}", "error")

    def _mk_logout(self):
        if self._mk_session:
            try:
                self._mk_session.get(f"{self._mk_base_url}/api/auth/logout", timeout=5)
            except Exception:
                pass
        self._mk_session = None
        self._mk_logged = False
        self._mk_status.set("Não conectado")
        self._mk_badge.configure(bg=self.COLORS["danger"])
        self._mk_log("Sessão encerrada.", "warn")
        self.root.after(0, self._mk_show_login_frame)

    def _mk_fetch_and_process(self):
        if not self._mk_logged or not self._mk_session:
            messagebox.showwarning("Makrolock", "Conecte-se primeiro usando e-mail e senha.")
            return
        if self._mk_busy:
            return
        self._mk_busy = True
        self._mk_fetch_btn.configure(state="disabled")
        threading.Thread(target=self._mk_do_fetch, daemon=True).start()

    def _mk_do_fetch(self):
        import tempfile, traceback as _tb
        try:
            from datetime import datetime as _dt

            def _parse(s):
                s = s.strip()
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        return _dt.strptime(s, fmt)
                    except ValueError:
                        pass
                raise ValueError(f"Data inválida: '{s}' — use DD/MM/AAAA")

            dt_start = _parse(self._mk_date_start.get())
            dt_end   = _parse(self._mk_date_end.get())
            dt_end   = dt_end.replace(hour=23, minute=59, second=59)

            # Tipos selecionados
            selected_types = [(k, self._mk_type_labels.get(k, k)) for k, v in self._mk_type_vars.items() if v.get()]
            if not selected_types:
                raise ValueError("Selecione ao menos um tipo de relatório.")

            # Grupos selecionados (None = sem filtro de grupo = todos)
            selected_groups = []  # lista de (group_id, group_name)
            if self._mk_group_vars:
                for gid, (var, name) in self._mk_group_vars.items():
                    if var.get():
                        selected_groups.append((gid, name))
                # Só exige grupo se nenhum computador individual estiver selecionado
                any_computer_selected = any(cv.get() for cv in self._mk_computer_vars.values())
                if not selected_groups and not any_computer_selected:
                    raise ValueError("Selecione ao menos um grupo ou computador.")
            else:
                selected_groups = [(None, "Todos")]

            base = self._mk_base_url
            from_s = int(dt_start.timestamp())
            to_s   = int(dt_end.timestamp())
            # Mantém também versão ms caso o endpoint de export precise
            from_ms = from_s * 1000
            to_ms   = to_s   * 1000

            self._mk_log(f"Período: {dt_start.strftime('%d/%m/%Y')} → {dt_end.strftime('%d/%m/%Y')}")
            self._mk_log(f"from_s={from_s}  to_s={to_s}  from_ms={from_ms}  to_ms={to_ms}")
            self._mk_log(f"Tipos: {[t for t,_ in selected_types]}  |  Grupos: {[n for _,n in selected_groups]}")

            added_files = []
            PAGE_SIZE = 1000  # registros por página

            def _fetch_all_pages(type_key, computer=None):
                """Busca todas as páginas do endpoint JSON e retorna lista de registros."""
                ep = f"{base}/api/front/reports"
                base_params = {
                    "type":      type_key,
                    "from":      from_s,
                    "to":        to_s,
                    "size":      PAGE_SIZE,
                    "sortBy":    "timestamp",
                    "sortDir":   "asc",
                }
                if computer:
                    base_params["computer"] = computer
                all_records = []
                page = 1
                total_items = None
                actual_page_size = None
                while True:
                    params = {**base_params, "page": page}
                    _r = self._mk_session.get(ep, params=params, timeout=60)
                    if _r.status_code != 200:
                        self._mk_log(f"    p{page} HTTP {_r.status_code}: {_r.text[:100]}", "warn")
                        break
                    try:
                        jdata = _r.json()
                    except Exception:
                        self._mk_log(f"    p{page} JSON inválido: {_r.text[:100]}", "warn")
                        break
                    records = jdata.get("data", [])
                    if total_items is None:
                        total_items = jdata.get("total", 0)
                        actual_page_size = jdata.get("pageSize", len(records) or 25)
                        total_pages = max(1, -(-total_items // actual_page_size))
                        self._mk_log(f"    total={total_items}  pageSize={actual_page_size}  páginas={total_pages}")
                    all_records.extend(records)
                    # Para quando não há mais registros ou chegamos na última página
                    if not records or len(all_records) >= total_items:
                        break
                    page += 1
                return all_records

            import datetime as _dt
            _logged_sample = [False]

            def _record_to_csv_line(rec):
                """Converte registro JSON para linha CSV com separador ;"""
                if not _logged_sample[0]:
                    self._mk_log(f"  sample rec: {rec}")
                    _logged_sample[0] = True

                def _esc(v):
                    s = str(v) if v is not None else ""
                    if ";" in s or '"' in s or "\n" in s:
                        s = '"' + s.replace('"', '""') + '"'
                    return s

                # Converte timestamp Unix → "YYYY-MM-DD HH:MM:SS"
                ts_raw = rec.get("timestamp", rec.get("date", ""))
                try:
                    ts = _dt.datetime.fromtimestamp(int(ts_raw)).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    ts = str(ts_raw)

                comp = rec.get("computer",  rec.get("computerName", rec.get("hostname", "")))
                user = rec.get("user",      rec.get("username", rec.get("login", "")))
                stat = rec.get("status",    rec.get("statusName", ""))
                # Duração em segundos (campo inteiro)
                dur_raw = rec.get("duration", rec.get("time", rec.get("seconds", "")))
                dur = str(dur_raw) if dur_raw is not None else ""
                cat  = rec.get("category",  rec.get("categoryName", rec.get("tag", "")))
                act  = rec.get("text",      rec.get("activity", rec.get("title", rec.get("window", ""))))
                app  = rec.get("app",       rec.get("process", rec.get("application", "")))
                return ";".join(_esc(v) for v in [ts, comp, user, stat, dur, cat, act, app])

            # Recolhe computadores selecionados individualmente
            all_computer_ids = [cname for cname, cv in self._mk_computer_vars.items() if cv.get()]
            if not all_computer_ids:
                seen_cids: set = set()
                for gid, gname in selected_groups:
                    if gid and gid in self._mk_group_computers:
                        for cid in self._mk_group_computers[gid]:
                            if cid not in seen_cids:
                                seen_cids.add(cid)
                                all_computer_ids.append(cid)
            self._mk_log(f"Computadores: {all_computer_ids or ['(todos)']}")

            CSV_HEADER = "Data/Hora;Computador;Usuário;Status;Duração;Categoria;Atividade;App"

            total_types = len(selected_types)
            done = 0
            for type_key, type_lbl in selected_types:
                done += 1
                self._mk_prog_var.set(int(done / total_types * 90))
                self._mk_prog_lbl.set(f"[{done}/{total_types}] {type_lbl}")
                self._mk_log(f"\nBaixando: [{done}/{total_types}] {type_lbl}")

                all_data_lines = []
                targets = all_computer_ids if all_computer_ids else [None]
                for idx, cid in enumerate(targets):
                    self._mk_log(f"  [{idx+1}/{len(targets)}] computer={cid or 'todos'}")
                    records = _fetch_all_pages(type_key, computer=cid)
                    if records:
                        for rec in records:
                            all_data_lines.append(_record_to_csv_line(rec))
                        self._mk_log(f"    +{len(records)} registros", "ok")
                    else:
                        self._mk_log(f"    sem dados")

                if not all_data_lines:
                    # Fallback: sem filtro de computer
                    self._mk_log(f"  fallback sem filtro")
                    records = _fetch_all_pages(type_key, computer=None)
                    for rec in records:
                        all_data_lines.append(_record_to_csv_line(rec))

                if not all_data_lines:
                    self._mk_log(f"  ⚠ Nenhum dado encontrado para {type_lbl}.", "warn")
                    continue

                # Grava CSV combinado
                combined = CSV_HEADER + "\n" + "\n".join(all_data_lines) + "\n"
                tmp = tempfile.NamedTemporaryFile(
                    delete=False, suffix=".csv",
                    prefix=f"mk_{type_key}_",
                    mode="w", encoding="utf-8-sig",
                )
                tmp.write(combined)
                tmp.close()
                self._mk_log(f"  ✅ {len(all_data_lines)} linhas → {tmp.name}", "ok")
                added_files.append(tmp.name)

            # Adiciona todos à fila sem processar
            for f in added_files:
                if f not in self._multi_files:
                    self._multi_files.append(f)

            self._mk_prog_var.set(100)
            msg = f"✅  {len(added_files)} arquivo(s) adicionado(s) à fila."
            self._mk_prog_lbl.set(msg)
            self._mk_log(msg, "ok")
            self._mk_log("Vá para ⚙ Processar e clique em '▶ Processar' quando quiser.", "ok")

            n_files = len(added_files)

            def _finish_on_ui():
                try:
                    self._files_refresh()
                except Exception:
                    pass
                if n_files:
                    self._show_page("processar")
                    messagebox.showinfo(
                        "Relatórios baixados",
                        f"{n_files} arquivo(s) adicionado(s) à fila.\n\n"
                        "Clique em '▶ Processar' para consolidar os dados."
                    )

            self.root.after(0, _finish_on_ui)

        except Exception as e:
            _err = str(e)
            self._mk_prog_lbl.set(f"❌  {_err}")
            self._mk_log(f"ERRO: {_err}", "error")
            self._mk_log(_tb.format_exc(), "error")
            self.root.after(0, lambda: messagebox.showerror("Erro Makrolock", _err))
        finally:
            self._mk_busy = False
            self.root.after(0, lambda: self._mk_fetch_btn.configure(state="normal"))

    # ── Filtros da listbox de classificação ──────────────────────────────
    def _mk_cls_apply_filter(self, extra_terms: list[str] | None = None):
        q = self._mk_cls_filter_var.get().strip().lower()
        if extra_terms is not None:
            terms = extra_terms
            self._mk_cls_active_terms = extra_terms  # persiste o filtro de categoria
        elif q:
            terms = [q]
            self._mk_cls_active_terms = [q]
        else:
            terms = self._mk_cls_active_terms        # mantém filtro anterior
        self._mk_cls_lb.delete(0, "end")
        self._mk_cls_filtered = []
        for i, (name, cat) in enumerate(self._mk_cls_items):
            show = not terms or any(t in name.lower() or t in cat.lower() for t in terms)
            if show:
                self._mk_cls_filtered.append(i)
                display = f"{name}  [{cat}]" if cat else name
                self._mk_cls_lb.insert("end", display)
        total = len(self._mk_cls_items)
        showing = len(self._mk_cls_filtered)
        self._mk_cls_count_lbl.set(
            f"Exibindo {showing} de {total}" if showing < total else f"{total} itens"
        )

    def _mk_cls_quick_filter(self, label: str):
        for lbl, terms in self._mk_quick_filters:
            if lbl == label:
                self._mk_cls_filter_var.set("")
                self._mk_cls_active_terms = []        # reseta para aplicar os novos termos
                self._mk_cls_apply_filter(terms)
                return

    def _mk_cls_clear_filter(self):
        self._mk_cls_filter_var.set("")
        self._mk_cls_active_terms = []
        self._mk_cls_apply_filter([])

    # ── Classificação em Massa ────────────────────────────────────────────
    def _mk_classify_load(self):
        if not self._mk_logged or not self._mk_session:
            messagebox.showwarning("Makrolock", "Conecte-se primeiro.")
            return
        self._mk_cls_btn_load.configure(state="disabled")
        threading.Thread(target=self._mk_do_classify_load, daemon=True).start()

    def _mk_do_classify_load(self):
        try:
            base = self._mk_base_url
            self._mk_log("Buscando atividades não classificadas (/productivity/list)...")
            self.root.after(0, lambda: (
                self._mk_cls_prog_var.set(10),
                self._mk_cls_prog_lbl.set("Consultando servidor...")
            ))

            resp = self._mk_session.get(f"{base}/api/front/productivity/list", timeout=30)
            if resp.status_code == 401 and self._mk_reauth():
                resp = self._mk_session.get(f"{base}/api/front/productivity/list", timeout=30)
            self._mk_log(f"HTTP {resp.status_code}  |  {len(resp.content)} bytes")

            if resp.status_code != 200:
                self._mk_log(f"Erro ao buscar lista: HTTP {resp.status_code}", "warn")
                self.root.after(0, lambda: (
                    self._mk_cls_prog_var.set(0),
                    self._mk_cls_prog_lbl.set("")
                ))
                return

            data = resp.json()
            entries = data.get("entries", data) if isinstance(data, dict) else data
            if not isinstance(entries, list):
                self._mk_log("Formato inesperado na resposta.", "warn")
                return

            self._mk_log(f"Total de atividades no servidor: {len(entries)}")
            self.root.after(0, lambda: self._mk_cls_prog_var.set(60))

            # Separa os não classificados (status == "notAssigned")
            unclassified = []
            for item in entries:
                if not isinstance(item, dict):
                    continue
                st = item.get("status", "")
                if st == "notAssigned":
                    nm  = item.get("name", "")
                    cat = item.get("categoryName", "")
                    if nm:
                        unclassified.append((nm, cat))

            self._mk_log(
                f"Produtivos: {sum(1 for e in entries if e.get('status')=='productive')}  "
                f"Improdutivos: {sum(1 for e in entries if e.get('status')=='unproductive')}  "
                f"Não classificados: {len(unclassified)}", "ok"
            )

            self._mk_cls_items = unclassified

            def _update_ui():
                self._mk_cls_items = unclassified
                self._mk_cls_filter_var.set("")
                self._mk_cls_apply_filter()
                total = len(unclassified)
                self._mk_cls_status.set(
                    f"{total} site(s)/app(s) sem classificação — use o filtro ou botões de categoria." if total
                    else "✅  Nenhum item sem classificação encontrado."
                )
                self._mk_cls_prog_var.set(0)
                self._mk_cls_prog_lbl.set("")
            self.root.after(0, _update_ui)
        except Exception as e:
            self._mk_log(f"Erro ao buscar lista: {e}", "error")
        finally:
            self.root.after(0, lambda: self._mk_cls_btn_load.configure(state="normal"))

    def _mk_classify_apply(self, status: str):
        if not self._mk_logged or not self._mk_session:
            messagebox.showwarning("Makrolock", "Conecte-se primeiro.")
            return
        lb_sel = list(self._mk_cls_lb.curselection())
        if not lb_sel:
            messagebox.showinfo("Classificação", "Selecione ao menos um item na lista.")
            return
        # Mapeia posições do listbox filtrado → índices reais em _mk_cls_items
        real_indices = [self._mk_cls_filtered[i] for i in lb_sel if i < len(self._mk_cls_filtered)]
        if not real_indices:
            return
        label_map = {"productive": "Produtivo", "unproductive": "Não Produtivo", "neutral": "Neutro"}
        label = label_map.get(status, status)
        total = len(real_indices)
        if not messagebox.askyesno("Confirmar classificação",
                                   f"Classificar {total} item(s) como '{label}'?\n\nEssa ação não pode ser desfeita."):
            return
        for btn in (self._mk_cls_btn_load, self._mk_cls_btn_prod,
                    self._mk_cls_btn_unprod, self._mk_cls_btn_neutral):
            btn.configure(state="disabled")
        threading.Thread(target=self._mk_do_classify_apply,
                         args=(real_indices, status), daemon=True).start()

    def _mk_auto_classify(self):
        """Aplica AUTO_CLASSIFY_RULES a todos os itens carregados e classifica via API."""
        if not self._mk_logged or not self._mk_session:
            messagebox.showwarning("Makrolock", "Conecte-se primeiro.")
            return
        if not self._mk_cls_items:
            messagebox.showinfo("Auto-classificar", "Busque os itens primeiro.")
            return

        import re as _re
        compiled = [(re.compile(pat, re.IGNORECASE), status)
                    for pat, status in AUTO_CLASSIFY_RULES]

        prod_items, unprod_items, neutral_items = [], [], []
        for i, (name, cat) in enumerate(self._mk_cls_items):
            for pattern, status in compiled:
                if pattern.search(name):
                    if status == "productive":
                        prod_items.append(i)
                    elif status == "unproductive":
                        unprod_items.append(i)
                    else:
                        neutral_items.append(i)
                    break  # primeira regra vence

        total = len(prod_items) + len(unprod_items) + len(neutral_items)
        if total == 0:
            messagebox.showinfo("Auto-classificar", "Nenhum item correspondeu às regras automáticas.")
            return

        msg = (f"Foram identificados {total} itens para classificação automática:\n\n"
               f"  ✅ Produtivo:      {len(prod_items)}\n"
               f"  ❌ Não Produtivo:  {len(unprod_items)}\n"
               f"  ➖ Neutro:         {len(neutral_items)}\n\n"
               f"Deseja aplicar agora?")
        if not messagebox.askyesno("Auto-classificar", msg):
            return

        for btn in (self._mk_cls_btn_load, self._mk_cls_btn_prod,
                    self._mk_cls_btn_unprod, self._mk_cls_btn_neutral,
                    self._mk_cls_btn_auto):
            btn.configure(state="disabled")

        # Junta todos em lista única (name, status) para processar sequencialmente
        auto_list: list[tuple[int, str]] = (
            [(i, "productive")   for i in prod_items] +
            [(i, "unproductive") for i in unprod_items] +
            [(i, "neutral")      for i in neutral_items]
        )
        # Ordena do maior para o menor índice — a remoção sequencial não desloca os anteriores
        auto_list.sort(key=lambda x: x[0], reverse=True)

        def _run():
            base = self._mk_base_url
            hdrs = {"Content-Type": "application/json"}
            done = 0
            errors = 0
            total_auto = len(auto_list)
            removed_indices = []

            for step, (idx, status) in enumerate(auto_list):
                if idx >= len(self._mk_cls_items):
                    continue
                name, _ = self._mk_cls_items[idx]
                pct = int((step + 1) / total_auto * 100)
                lbl = f"Auto {step+1}/{total_auto}: {name[:55]}"
                self.root.after(0, lambda p=pct, l=lbl: (
                    self._mk_cls_prog_var.set(p),
                    self._mk_cls_prog_lbl.set(l)
                ))
                try:
                    resp = self._mk_session.put(
                        f"{base}/api/front/productivity/classify",
                        json={"name": name, "status": status},
                        headers=hdrs, timeout=10)
                    if resp.status_code == 401:
                        if self._mk_reauth():
                            resp = self._mk_session.put(
                                f"{base}/api/front/productivity/classify",
                                json={"name": name, "status": status},
                                headers=hdrs, timeout=10)
                        else:
                            self._mk_log("Re-autenticação falhou. Abortando auto-classify.", "error")
                            break
                    ok = resp.status_code == 200
                    for gid in self._mk_group_vars:
                        gr = self._mk_session.put(
                            f"{base}/api/front/productivity/classify-group",
                            json={"name": name, "groupId": int(gid), "status": status},
                            headers=hdrs, timeout=10)
                        if gr.status_code == 401 and self._mk_reauth():
                            self._mk_session.put(
                                f"{base}/api/front/productivity/classify-group",
                                json={"name": name, "groupId": int(gid), "status": status},
                                headers=hdrs, timeout=10)
                    if ok:
                        done += 1
                        removed_indices.append(idx)
                    else:
                        errors += 1
                        self._mk_log(f"  ⚠ '{name}': HTTP {resp.status_code}", "warn")
                except Exception as e:
                    errors += 1
                    self._mk_log(f"  ⚠ '{name}': {e}", "warn")

            msg = f"⚡ Auto-classificação: {done} item(s)"
            if errors:
                msg += f"  |  ⚠ {errors} erro(s)"
            self._mk_log(msg, "ok" if not errors else "warn")

            def _finish():
                self._mk_cls_prog_var.set(100)
                self._mk_cls_prog_lbl.set(msg)
                for i in sorted(removed_indices, reverse=True):
                    if i < len(self._mk_cls_items):
                        self._mk_cls_items.pop(i)
                self._mk_cls_apply_filter()
                self._mk_cls_status.set(
                    f"{msg} — {len(self._mk_cls_items)} restantes sem classificação.")
                for btn in (self._mk_cls_btn_load, self._mk_cls_btn_prod,
                            self._mk_cls_btn_unprod, self._mk_cls_btn_neutral,
                            self._mk_cls_btn_auto):
                    btn.configure(state="normal")
            self.root.after(0, _finish)

        threading.Thread(target=_run, daemon=True).start()

    def _mk_reauth(self) -> bool:
        """Re-autentica usando credenciais guardadas em memória. Retorna True se ok."""
        email = self._mk_email.get().strip()
        passwd_md5 = self._mk_pass_md5
        if not email or not passwd_md5:
            return False
        try:
            import requests as _req
            base = self._mk_base_url
            sess = _req.Session()
            sess.headers.update(self._mk_session.headers if self._mk_session else {})
            resp = sess.post(f"{base}/api/auth/login",
                             json={"email": email, "password": passwd_md5},
                             timeout=15)
            if resp.status_code == 200:
                self._mk_session = sess
                self._mk_logged = True
                self._mk_log("🔄 Re-autenticado automaticamente.", "ok")
                return True
        except Exception as e:
            self._mk_log(f"Re-autenticação falhou: {e}", "error")
        return False

    def _mk_do_classify_apply(self, indices: list, status: str, _silent: bool = False):
        base = self._mk_base_url
        total = len(indices)
        done = 0
        errors = 0
        try:
            for i, idx in enumerate(indices):
                if idx >= len(self._mk_cls_items):
                    continue
                name, _ = self._mk_cls_items[idx]
                pct = int((i + 1) / total * 100)
                lbl = f"Classificando {i+1}/{total}: {name[:60]}"
                self.root.after(0, lambda p=pct, l=lbl: (
                    self._mk_cls_prog_var.set(p),
                    self._mk_cls_prog_lbl.set(l)
                ))
                try:
                    hdrs = {"Content-Type": "application/json"}
                    # PUT classify (global) — com retry automático em 401
                    resp = self._mk_session.put(
                        f"{base}/api/front/productivity/classify",
                        json={"name": name, "status": status},
                        headers=hdrs, timeout=10,
                    )
                    if resp.status_code == 401:
                        self._mk_log("Sessão expirada — tentando re-autenticar...", "warn")
                        if self._mk_reauth():
                            resp = self._mk_session.put(
                                f"{base}/api/front/productivity/classify",
                                json={"name": name, "status": status},
                                headers=hdrs, timeout=10,
                            )
                        else:
                            self._mk_log("Re-autenticação falhou. Abortando.", "error")
                            errors += total - i
                            break
                    ok = resp.status_code == 200
                    # PUT classify-group para cada grupo
                    for gid in self._mk_group_vars:
                        gr = self._mk_session.put(
                            f"{base}/api/front/productivity/classify-group",
                            json={"name": name, "groupId": int(gid), "status": status},
                            headers=hdrs, timeout=10,
                        )
                        if gr.status_code == 401:
                            if self._mk_reauth():
                                self._mk_session.put(
                                    f"{base}/api/front/productivity/classify-group",
                                    json={"name": name, "groupId": int(gid), "status": status},
                                    headers=hdrs, timeout=10,
                                )
                    if ok:
                        done += 1
                    else:
                        errors += 1
                        self._mk_log(f"  ⚠ '{name}': HTTP {resp.status_code}  {resp.text[:100]}", "warn")
                except Exception as e:
                    errors += 1
                    self._mk_log(f"  ⚠ '{name}': {e}", "warn")

            msg = f"✅  {done} item(s) classificado(s)"
            if errors:
                msg += f"  |  ⚠ {errors} erro(s)"
            self._mk_log(msg, "ok" if not errors else "warn")

            def _finish():
                self._mk_cls_prog_var.set(100)
                self._mk_cls_prog_lbl.set(msg)
                self._mk_cls_status.set(f"{msg} — {len(self._mk_cls_items) - len(indices)} restantes.")
                # Remove itens classificados da lista interna
                for i in sorted(indices, reverse=True):
                    if i < len(self._mk_cls_items):
                        self._mk_cls_items.pop(i)
                # Reaplica filtro atual para atualizar listbox
                self._mk_cls_apply_filter()
            self.root.after(0, _finish)
        except Exception as e:
            self._mk_log(f"ERRO na classificação: {e}", "error")
        finally:
            if not _silent:
                def _re_enable():
                    for btn in (self._mk_cls_btn_load, self._mk_cls_btn_prod,
                                self._mk_cls_btn_unprod, self._mk_cls_btn_neutral,
                                self._mk_cls_btn_auto):
                        btn.configure(state="normal")
                self.root.after(0, _re_enable)

    def _build_process_tab(self, parent):
        c = self.COLORS

        top = tk.Frame(parent, bg=c["bg"])
        top.pack(fill="x", pady=(0, 16))
        tk.Label(top, text="⚙️  Processar", bg=c["bg"], fg=c["text"],
                 font=("Segoe UI", 15, "bold")).pack(side="left")

        outer, card = self._card(parent, accent=True)
        outer.pack(fill="x", pady=(0, 12))
        card.columnconfigure(0, weight=1)

        self._section_label(card, "ARQUIVOS CSV / XLSX  (selecione um ou vários)")

        # Listbox com scrollbar para múltiplos arquivos
        lb_wrap = tk.Frame(card, bg=c["card"])
        lb_wrap.pack(fill="x", pady=(0, 4))
        lb_wrap.columnconfigure(0, weight=1)

        lb_scroll = ttk.Scrollbar(lb_wrap, orient="vertical")
        self._files_lb = tk.Listbox(
            lb_wrap, height=5, selectmode="extended",
            bg=c["bg"], fg=c["text"], relief="flat", bd=0,
            highlightthickness=1, highlightbackground=c["border"],
            font=("Segoe UI", 9), yscrollcommand=lb_scroll.set,
        )
        lb_scroll.config(command=self._files_lb.yview)
        self._files_lb.grid(row=0, column=0, sticky="ew")
        lb_scroll.grid(row=0, column=1, sticky="ns")

        # Contador e botões de ação
        self._files_count_var = tk.StringVar(value="Nenhum arquivo selecionado")
        tk.Label(card, textvariable=self._files_count_var, bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 4))

        files_btn_row = ttk.Frame(card, style="Card.TFrame")
        files_btn_row.pack(anchor="w", pady=(0, 8))
        ttk.Button(files_btn_row, text="+ Adicionar arquivos", style="Accent.TButton",
                   command=self._files_add).pack(side="left")
        ttk.Button(files_btn_row, text="Remover selecionados", style="Ghost.TButton",
                   command=self._files_remove).pack(side="left", padx=(8, 0))
        ttk.Button(files_btn_row, text="Limpar lista", style="Ghost.TButton",
                   command=self._files_clear).pack(side="left", padx=(8, 0))

        self._section_label(card, "PASTA DE SAÍDA")
        row1 = ttk.Frame(card, style="Card.TFrame")
        row1.pack(fill="x", pady=(0, 4))
        row1.columnconfigure(0, weight=1)
        ttk.Entry(row1, textvariable=self.output_path).grid(row=0, column=0, sticky="ew")
        ttk.Button(row1, text="Escolher", command=self.select_output).grid(row=0, column=1, padx=(8, 0))

        self._section_label(card, "NOME DO ARQUIVO DE SAÍDA  (deixe vazio para gerar automático)")
        ttk.Entry(card, textvariable=self.output_filename).pack(fill="x", pady=(0, 4))

        self._section_label(card, "TERMOS/EVENTOS PARA EXCLUIR  (separados por vírgula)")
        ttk.Entry(card, textvariable=self.exclude_terms).pack(fill="x", pady=(0, 4))

        self._section_label(card, "SALVAR BASE TRATADA COMO")
        fmt_frame = ttk.Frame(card, style="Card.TFrame")
        fmt_frame.pack(anchor="w", pady=(0, 6))
        ttk.Radiobutton(fmt_frame, text="CSV", variable=self.base_export_format, value="csv").pack(side="left")
        ttk.Radiobutton(fmt_frame, text="Excel .xlsx", variable=self.base_export_format, value="xlsx").pack(side="left", padx=(16, 0))
        ttk.Radiobutton(fmt_frame, text="CSV + Excel", variable=self.base_export_format, value="csv_xlsx").pack(side="left", padx=(16, 0))

        action = ttk.Frame(parent)
        action.pack(fill="x", pady=(12, 0))
        ttk.Button(action, text="▶  Processar", style="Accent.TButton", command=self.start_processing).pack(side="left")
        ttk.Button(action, text="Abrir pasta de saída", style="Ghost.TButton", command=self.open_output_folder).pack(side="left", padx=(8, 0))

        # Separador — importar base já tratada
        tk.Frame(parent, bg=c["border"], height=1).pack(fill="x", pady=(18, 0))

        outer_imp, imp_card = self._card(parent, padding=(14, 12), accent=False)
        outer_imp.pack(fill="x", pady=(12, 0))
        imp_card.columnconfigure(0, weight=1)

        tk.Label(imp_card, text="⬆  IMPORTAR BASE JÁ TRATADA", bg=c["card"], fg=c["ok"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        tk.Label(imp_card,
                 text="Carregue diretamente um CSV ou XLSX que já foi processado antes — pula o processamento e vai direto para os Relatórios.",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 9), wraplength=700, justify="left").pack(anchor="w", pady=(4, 10))

        imp_row = ttk.Frame(imp_card, style="Card.TFrame")
        imp_row.pack(fill="x")
        imp_row.columnconfigure(0, weight=1)
        self._import_path = tk.StringVar()
        e_imp = ttk.Entry(imp_row, textvariable=self._import_path)
        e_imp.grid(row=0, column=0, sticky="ew")
        # Garante que o Entry mostre o final do caminho (texto mais relevante)
        self._import_path.trace_add("write", lambda *_: e_imp.xview_moveto(1.0))
        ttk.Button(imp_row, text="Selecionar", command=self._select_import_file).grid(row=0, column=1, padx=(8, 0))

        imp_action = ttk.Frame(parent)
        imp_action.pack(fill="x", pady=(10, 0))
        ttk.Button(imp_action, text="⬆  Importar Base Tratada", style="Ok.TButton",
                   command=self.start_import_treated).pack(side="left")

        # ── Barra de progresso dedicada ──────────────────────────────────
        tk.Frame(parent, bg=c["border"], height=1).pack(fill="x", pady=(18, 0))

        outer_prog, prog_card = self._card(parent, padding=(14, 12))
        outer_prog.pack(fill="x", pady=(12, 0))

        prog_top = ttk.Frame(prog_card, style="Card.TFrame")
        prog_top.pack(fill="x", pady=(0, 8))
        tk.Label(prog_top, text="PROGRESSO", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        self._proc_pct_var = tk.StringVar(value="")
        tk.Label(prog_top, textvariable=self._proc_pct_var, bg=c["card"], fg=c["accent"],
                 font=("Segoe UI", 8, "bold")).pack(side="right")

        self._proc_bar = ttk.Progressbar(prog_card, variable=self.progress_value, maximum=100,
                                         style="Horizontal.TProgressbar")
        self._proc_bar.pack(fill="x", pady=(0, 8))

        self._proc_step_var = tk.StringVar(value="Aguardando início...")
        tk.Label(prog_card, textvariable=self._proc_step_var, bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9), anchor="w").pack(fill="x")

        # Etapas como indicadores visuais
        steps_frame = ttk.Frame(prog_card, style="Card.TFrame")
        steps_frame.pack(fill="x", pady=(12, 0))

        self._step_labels = []
        steps = [
            "Lendo arquivo",
            "Tratando dados",
            "Aplicando filtros",
            "Exportando base",
            "Concluído",
        ]
        for i, step in enumerate(steps):
            col_frame = ttk.Frame(steps_frame, style="Card.TFrame")
            col_frame.pack(side="left", expand=True, fill="x")

            dot = tk.Canvas(col_frame, width=10, height=10,
                            bg=c["card"], highlightthickness=0)
            dot.pack(anchor="center", pady=(0, 4))
            dot.create_oval(1, 1, 9, 9, fill=c["border"], outline=c["border2"], tags="dot")

            lbl = tk.Label(col_frame, text=step, bg=c["card"], fg=c["muted"],
                           font=("Segoe UI", 8), wraplength=90, justify="center")
            lbl.pack(anchor="center")

            if i < len(steps) - 1:
                tk.Frame(col_frame, bg=c["border"], height=1).pack(
                    side="right", fill=None, expand=False)

            self._step_labels.append((dot, lbl))

        # Dicas
        outer2, tips = self._card(parent, padding=(14, 12))
        outer2.pack(fill="x", pady=(14, 0))
        tk.Label(tips, text="COMO USAR", bg=c["card"], fg=c["accent"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        dicas = [
            "Opção A — Processar do zero:  selecione o CSV/XLSX bruto do Macro Lock e clique em Processar.",
            "Opção B — Importar base tratada:  selecione um arquivo já tratado (ex: macro_lock_base_tratada_completa.csv) e clique em Importar.",
            "Após qualquer uma das opções, vá para Relatórios e gere os filtros sem precisar repetir o processo.",
        ]
        for d in dicas:
            tk.Label(tips, text=d, bg=c["card"], fg=c["muted"],
                     font=("Segoe UI", 9), wraplength=700, justify="left").pack(anchor="w", pady=2)

    def _build_names_tab(self, parent):
        c = self.COLORS
        top = ttk.Frame(parent)
        top.pack(fill="x", pady=(0, 10))
        tk.Label(top, text="✏️  Corrigir Nomes", bg=c["bg"], fg=c["text"],
                 font=("Segoe UI", 15, "bold")).pack(side="left")
        ttk.Button(top, text="Limpar correções", style="Ghost.TButton", command=self._clear_renames).pack(side="right")
        ttk.Button(top, text="↺  Carregar do CSV", command=self._load_names_from_csv).pack(side="right", padx=(0, 8))

        outer_hint, hint = self._card(parent, padding=(10, 8))
        outer_hint.pack(fill="x", pady=(0, 10))
        tk.Label(hint,
                 text="Clique duas vezes na coluna 'Novo Nome' para corrigir. Deixe em branco para manter o original.",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w")

        panes = ttk.Frame(parent)
        panes.pack(fill="both", expand=True)
        panes.columnconfigure(0, weight=1)
        panes.columnconfigure(1, weight=1)
        panes.rowconfigure(0, weight=1)

        for col_idx, col_name in enumerate(("Equipamento", "Usuario")):
            outer, frame = self._card(panes, padding=(10, 8))
            outer.grid(row=0, column=col_idx, sticky="nsew", padx=(0 if col_idx == 0 else 8, 0))
            frame.rowconfigure(1, weight=1)
            frame.columnconfigure(0, weight=1)
            tk.Label(frame, text=col_name.upper(), bg=c["card"], fg=c["muted"],
                     font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))

            tree_frame = ttk.Frame(frame, style="Card.TFrame")
            tree_frame.grid(row=1, column=0, sticky="nsew")
            tree_frame.rowconfigure(0, weight=1)
            tree_frame.columnconfigure(0, weight=1)

            tv = ttk.Treeview(tree_frame, columns=("original", "novo"), show="headings", height=14)
            tv.heading("original", text="Nome original")
            tv.heading("novo", text="✎  Novo Nome")
            tv.column("original", width=200, stretch=True)
            tv.column("novo", width=200, stretch=True)
            tv.grid(row=0, column=0, sticky="nsew")

            sb = ttk.Scrollbar(tree_frame, orient="vertical", command=tv.yview)
            sb.grid(row=0, column=1, sticky="ns")
            tv.configure(yscrollcommand=sb.set)

            tv.bind("<Double-1>", lambda e, t=tv, cn=col_name: self._on_name_double_click(e, t, cn))

            if col_name == "Equipamento":
                self._tv_equipamento = tv
            else:
                self._tv_usuario = tv

    def _load_names_from_csv(self):
        raw = self.file_path.get().strip()
        if not raw:
            messagebox.showwarning("CSV não selecionado", "Selecione o arquivo CSV na aba 'Processar' primeiro.")
            return
        csv_path = Path(raw)
        if not csv_path.is_file():
            messagebox.showwarning("CSV não encontrado", f"Arquivo não encontrado:\n{csv_path}")
            return
        self.set_status("Lendo valores únicos do CSV...", 10)
        threading.Thread(target=self._do_load_names, args=(csv_path,), daemon=True).start()

    def _do_load_names(self, csv_path: Path):
        try:
            import csv as csv_mod
            equips, users = set(), set()
            if csv_path.suffix.lower() in (".xlsx", ".xls"):
                import openpyxl
                wb = openpyxl.load_workbook(str(csv_path), read_only=True, data_only=True)
                ws = wb.active
                for row in ws.iter_rows(values_only=True):
                    if row and len(row) > 2:
                        eq = str(row[1] or "").strip()
                        us = str(row[2] or "").strip()
                        if eq: equips.add(eq)
                        if us: users.add(us)
                wb.close()
            else:
                encoding = self._detect_source_encoding(csv_path)
                delim = self._detect_delimiter(csv_path, encoding)
                with open(csv_path, "r", encoding=encoding, errors="replace", newline="") as f:
                    reader = csv_mod.reader(f, delimiter=delim)
                    for row in reader:
                        if len(row) > 2:
                            eq = str(row[1]).strip()
                            us = str(row[2]).strip()
                            if eq: equips.add(eq)
                            if us: users.add(us)
            self.root.after(0, self._populate_name_trees, sorted(equips), sorted(users))
            self.set_status(f"Valores carregados: {len(equips)} equipamentos, {len(users)} usuários.", 100)
        except Exception as e:
            self.write_log(f"Erro ao ler CSV para nomes: {e}")
            messagebox.showerror("Erro", str(e))

    def _populate_name_trees(self, equips, users):
        for tv, values, col_name in (
            (self._tv_equipamento, equips, "Equipamento"),
            (self._tv_usuario, users, "Usuario"),
        ):
            tv.delete(*tv.get_children())
            existing = self.rename_maps[col_name]
            for v in values:
                novo = existing.get(v, "")
                tv.insert("", "end", iid=v, values=(v, novo))

    def _on_name_double_click(self, event, tv, col_name):
        region = tv.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = tv.identify_column(event.x)
        if col != "#2":
            return
        row_id = tv.identify_row(event.y)
        if not row_id:
            return
        x, y, w, h = tv.bbox(row_id, "#2")
        original_val = tv.set(row_id, "original")
        current_novo = tv.set(row_id, "novo")

        entry_var = tk.StringVar(value=current_novo)
        entry = tk.Entry(tv, textvariable=entry_var, font=("Segoe UI", 10))
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.select_range(0, "end")

        def confirm(event=None):
            new_val = entry_var.get().strip()
            tv.set(row_id, "novo", new_val)
            if new_val:
                self.rename_maps[col_name][original_val] = new_val
            else:
                self.rename_maps[col_name].pop(original_val, None)
            entry.destroy()

        entry.bind("<Return>", confirm)
        entry.bind("<Tab>", confirm)
        entry.bind("<FocusOut>", confirm)
        entry.bind("<Escape>", lambda e: entry.destroy())

    def _clear_renames(self):
        self.rename_maps = {"Equipamento": {}, "Usuario": {}}
        for tv in (self._tv_equipamento, self._tv_usuario):
            for iid in tv.get_children():
                tv.set(iid, "novo", "")
        self.write_log("Correções de nomes limpas.")

    # ------------------------------------------------------------------
    # Blacklist
    # ------------------------------------------------------------------

    def _load_blacklist(self):
        import json
        try:
            if self._blacklist_file.exists():
                data = json.loads(self._blacklist_file.read_text(encoding="utf-8"))
                self.blacklist_terms = [str(t) for t in data if str(t).strip()]
        except Exception:
            self.blacklist_terms = []

    def _save_blacklist(self):
        import json
        try:
            self._blacklist_file.write_text(
                json.dumps(self.blacklist_terms, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as e:
            self.write_log(f"Aviso: não foi possível salvar blacklist: {e}")

    def _build_blacklist_tab(self, parent):
        c = self.COLORS

        top = ttk.Frame(parent)
        top.pack(fill="x", pady=(0, 10))
        tk.Label(top, text="🚫  Blacklist", bg=c["bg"], fg=c["text"],
                 font=("Segoe UI", 15, "bold")).pack(side="left")

        outer_hint, hint = self._card(parent, padding=(10, 8))
        outer_hint.pack(fill="x", pady=(0, 10))
        tk.Label(hint,
                 text="Cadastre termos para monitorar (Aplicativo / JanelaEvento). Use * como coringa: *youtube*, chrome.exe",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w")

        body = ttk.Frame(parent)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=0)
        body.rowconfigure(0, weight=1)

        # Lista de termos
        outer_list, list_frame = self._card(body, padding=(10, 8))
        outer_list.grid(row=0, column=0, sticky="nsew")
        list_frame.rowconfigure(1, weight=1)
        list_frame.columnconfigure(0, weight=1)
        tk.Label(list_frame, text="TERMOS CADASTRADOS", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))

        lb_frame = ttk.Frame(list_frame, style="Card.TFrame")
        lb_frame.grid(row=1, column=0, sticky="nsew")
        lb_frame.rowconfigure(0, weight=1)
        lb_frame.columnconfigure(0, weight=1)

        self._blacklist_lb = tk.Listbox(
            lb_frame,
            bg=c["card"], fg=c["text"], selectbackground=c["accent2"],
            selectforeground="#ffffff", font=("Consolas", 10),
            relief="flat", borderwidth=0, highlightthickness=0,
            activestyle="none",
        )
        self._blacklist_lb.grid(row=0, column=0, sticky="nsew")
        sb_bl = ttk.Scrollbar(lb_frame, orient="vertical", command=self._blacklist_lb.yview)
        sb_bl.grid(row=0, column=1, sticky="ns")
        self._blacklist_lb.configure(yscrollcommand=sb_bl.set)

        # Painel de ações à direita
        outer_act, actions = self._card(body, padding=(14, 12))
        outer_act.grid(row=0, column=1, sticky="n", padx=(10, 0))

        tk.Label(actions, text="ADICIONAR TERMO", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

        self._bl_entry_var = tk.StringVar()
        entry = ttk.Entry(actions, textvariable=self._bl_entry_var, width=26)
        entry.pack(fill="x", pady=(0, 6))
        entry.bind("<Return>", lambda e: self._bl_add())

        ttk.Button(actions, text="+ Adicionar", style="Accent.TButton", command=self._bl_add).pack(fill="x", pady=(0, 14))

        tk.Frame(actions, bg=c["border"], height=1).pack(fill="x", pady=(0, 12))

        ttk.Button(actions, text="Remover selecionado", style="Ghost.TButton", command=self._bl_remove).pack(fill="x", pady=(0, 6))
        ttk.Button(actions, text="Limpar tudo", style="Ghost.TButton", command=self._bl_clear).pack(fill="x", pady=(0, 6))

        tk.Frame(actions, bg=c["border"], height=1).pack(fill="x", pady=(10, 8))

        self._bl_count_var = tk.StringVar(value="0 termos")
        tk.Label(actions, textvariable=self._bl_count_var, bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9)).pack(anchor="w")

        # Popula lista com o que foi carregado do arquivo
        self._bl_refresh_listbox()

    def _bl_refresh_listbox(self):
        self._blacklist_lb.delete(0, "end")
        for term in self.blacklist_terms:
            self._blacklist_lb.insert("end", term)
        self._bl_count_var.set(f"{len(self.blacklist_terms)} termo(s) cadastrado(s)")

    def _bl_add(self):
        term = self._bl_entry_var.get().strip()
        if not term:
            return
        if term in self.blacklist_terms:
            messagebox.showinfo("Duplicado", f"O termo '{term}' já está na blacklist.")
            return
        self.blacklist_terms.append(term)
        self._save_blacklist()
        self._bl_refresh_listbox()
        self._bl_entry_var.set("")
        self.write_log(f"Blacklist: adicionado '{term}'")

    def _bl_remove(self):
        sel = self._blacklist_lb.curselection()
        if not sel:
            return
        idx = sel[0]
        term = self.blacklist_terms[idx]
        self.blacklist_terms.pop(idx)
        self._save_blacklist()
        self._bl_refresh_listbox()
        self.write_log(f"Blacklist: removido '{term}'")

    def _bl_clear(self):
        if not self.blacklist_terms:
            return
        if messagebox.askyesno("Limpar blacklist", "Remover todos os termos da blacklist?"):
            self.blacklist_terms.clear()
            self._save_blacklist()
            self._bl_refresh_listbox()
            self.write_log("Blacklist limpa.")

    def _build_reports_tab(self, parent):
        c = self.COLORS

        top_r = tk.Frame(parent, bg=c["bg"])
        top_r.pack(fill="x", pady=(0, 16))
        tk.Label(top_r, text="📋  Relatórios", bg=c["bg"], fg=c["text"],
                 font=("Segoe UI", 15, "bold")).pack(side="left")

        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)

        tab_padrao    = ttk.Frame(nb, style="Card.TFrame", padding=(16, 14))
        tab_especif   = ttk.Frame(nb, style="Card.TFrame", padding=(16, 14))
        nb.add(tab_padrao,  text="  Padrão  ")
        nb.add(tab_especif, text="  Específicos  ")

        self._build_reports_padrao(tab_padrao)
        self._build_reports_especificos(tab_especif)
        return  # _build_reports_padrao / _especificos constroem o restante

    def _build_reports_padrao(self, parent):
        c = self.COLORS

        REPORT_HINTS = {
            "fora_horario":             "Traz tudo FORA do horário comercial (ex: 18:00 → 07:00).",
            "fim_semana":               "Apenas sábado e domingo. Horário não é usado.",
            "fora_horario_fim_semana":  "Fora do horário OU final de semana.",
            "por_usuario_equipamento":  "Resumo por máquina/usuário. Horário não é usado.",
            "intervalo_horario":        "Eventos DENTRO do intervalo. Ex: 00:00 → 09:00 mostra quem ligou cedo.",
            "blacklist":                "Apenas eventos que bateram com a Blacklist (Aplicativo / JanelaEvento).",
            "ociosidade_diario":        "Resumo DIÁRIO de ociosidade por usuário. Seg–Sex, janelas 08–12h e 13–17h. Dias úteis com atividade registrada (colunas: Ativo e Ocioso em horas e %).",
            "ociosidade_mensal":        "Resumo MENSAL de ociosidade por usuário. Agrega por mês os dias úteis com atividade registrada. Colunas: Ativo e Ocioso em horas e %.",
        }
        self._report_hint_var = tk.StringVar(value=REPORT_HINTS.get(self.report_type.get(), ""))

        def _on_report_type_change(*_):
            self._report_hint_var.set(REPORT_HINTS.get(self.report_type.get(), ""))
        self.report_type.trace_add("write", _on_report_type_change)

        outer, card = self._card(parent)
        outer.pack(fill="x", pady=(0, 12))
        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=1)

        # Tipo
        self._section_label(card, "TIPO DE RELATÓRIO")
        combo = ttk.Combobox(card, textvariable=self.report_type, values=list(REPORTS.keys()), state="readonly")
        combo.pack(fill="x", pady=(0, 4))
        tk.Label(card, textvariable=self._report_hint_var, bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 9), wraplength=700, justify="left").pack(anchor="w", pady=(0, 12))

        tk.Frame(card, bg=c["border"], height=1).pack(fill="x", pady=(0, 12))

        # Horários e filtros em grid 2 colunas
        grid = ttk.Frame(card, style="Card.TFrame")
        grid.pack(fill="x")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        def lbl(parent, text, row, col, padx=(0,0)):
            tk.Label(parent, text=text, bg=c["card"], fg=c["muted"],
                     font=("Segoe UI", 8)).grid(row=row, column=col, sticky="w", padx=padx, pady=(0, 2))

        lbl(grid, "HORÁRIO INICIAL",  0, 0)
        lbl(grid, "HORÁRIO FINAL",    0, 1, (16, 0))
        _e_ns = ttk.Entry(grid, textvariable=self.night_start, width=10)
        _e_ns.grid(row=1, column=0, sticky="w", pady=(0, 10))
        _e_ne = ttk.Entry(grid, textvariable=self.night_end, width=10)
        _e_ne.grid(row=1, column=1, sticky="w", padx=(16, 0), pady=(0, 10))
        self._bind_time_mask(_e_ns, self.night_start, next_widget=_e_ne)
        self._bind_time_mask(_e_ne, self.night_end)

        lbl(grid, "USUÁRIO (opcional)",      2, 0)
        lbl(grid, "EQUIPAMENTO (opcional)",  2, 1, (16, 0))
        ttk.Entry(grid, textvariable=self.report_user,      width=28).grid(row=3, column=0, sticky="ew", pady=(0, 10))
        ttk.Entry(grid, textvariable=self.report_equipment, width=28).grid(row=3, column=1, sticky="ew", padx=(16, 0), pady=(0, 10))

        lbl(grid, "DATA INICIAL (opcional)", 4, 0)
        lbl(grid, "DATA FINAL (opcional)",   4, 1, (16, 0))
        _e_rds = ttk.Entry(grid, textvariable=self.report_date_start, width=18)
        _e_rds.grid(row=5, column=0, sticky="w", pady=(0, 4))
        _e_rde = ttk.Entry(grid, textvariable=self.report_date_end, width=18)
        _e_rde.grid(row=5, column=1, sticky="w", padx=(16, 0), pady=(0, 4))
        self._bind_date_mask(_e_rds, self.report_date_start, next_widget=_e_rde)
        self._bind_date_mask(_e_rde, self.report_date_end)
        tk.Label(grid, text="DD/MM/AAAA ou AAAA-MM-DD  —  vazio = toda a base",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 8)).grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 12))

        tk.Frame(card, bg=c["border"], height=1).pack(fill="x", pady=(0, 10))

        self._section_label(card, "FORMATO DE SAÍDA")
        fmt_row = ttk.Frame(card, style="Card.TFrame")
        fmt_row.pack(anchor="w", pady=(0, 4))
        ttk.Radiobutton(fmt_row, text="Excel .xlsx", variable=self.report_format, value="xlsx").pack(side="left")
        ttk.Radiobutton(fmt_row, text="CSV .csv",    variable=self.report_format, value="csv").pack(side="left", padx=(16, 0))

        dash_row = ttk.Frame(card, style="Card.TFrame")
        dash_row.pack(anchor="w", pady=(4, 0))
        ttk.Checkbutton(dash_row, text="📊 Incluir aba Dashboard (por máquina) no Excel",
                        variable=self.report_dashboard).pack(side="left")

        action = ttk.Frame(parent)
        action.pack(fill="x", pady=(12, 0))
        ttk.Button(action, text="▶  Gerar Relatório", style="Accent.TButton", command=self.start_report).pack(side="left")
        ttk.Button(action, text="Abrir pasta", style="Ghost.TButton", command=self.open_reports_folder).pack(side="left", padx=(8, 0))

    def _build_reports_especificos(self, parent):
        c = self.COLORS

        tk.Label(parent, text="Relatório Específicos — fora do horário + finais de semana",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 12))

        outer, card = self._card(parent)
        outer.pack(fill="x", pady=(0, 12))

        # ── Usuários ──────────────────────────────────────────────────────
        self._section_label(card, "USUÁRIOS (selecione um ou mais)")

        users_frame = tk.Frame(card, bg=c["card"])
        users_frame.pack(fill="x", pady=(0, 4))

        self._esp_user_listbox = tk.Listbox(
            users_frame, selectmode="multiple", height=8,
            bg=c["card2"], fg=c["text"], selectbackground=c["teal"],
            selectforeground="#FFFFFF", font=("Segoe UI", 9),
            highlightthickness=1, highlightbackground=c["border2"],
            activestyle="none",
        )
        sb = ttk.Scrollbar(users_frame, orient="vertical", command=self._esp_user_listbox.yview)
        self._esp_user_listbox.configure(yscrollcommand=sb.set)
        self._esp_user_listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        btn_load = ttk.Button(card, text="↻  Carregar usuários da base",
                              command=self._esp_load_users)
        btn_load.pack(anchor="w", pady=(6, 0))

        tk.Frame(card, bg=c["border"], height=1).pack(fill="x", pady=(12, 12))

        # ── Filtros de período ────────────────────────────────────────────
        self._section_label(card, "FILTROS DE PERÍODO")

        filters_frame = tk.Frame(card, bg=c["card"])
        filters_frame.pack(anchor="w", pady=(0, 12))

        self._esp_seg_sex_var  = tk.BooleanVar(value=True)
        self._esp_sab_dom_var  = tk.BooleanVar(value=True)
        self._esp_seg_sex_start = tk.StringVar(value="18:00")
        self._esp_seg_sex_end   = tk.StringVar(value="00:00")

        # Seg-Sex
        row1 = tk.Frame(filters_frame, bg=c["card"])
        row1.pack(anchor="w", pady=(0, 6))
        tk.Checkbutton(row1, text="Seg a Sex  —  das", variable=self._esp_seg_sex_var,
                       bg=c["card"], fg=c["text"], selectcolor=c["card2"],
                       activebackground=c["card"], font=("Segoe UI", 9),
                       cursor="hand2").pack(side="left")
        ttk.Entry(row1, textvariable=self._esp_seg_sex_start, width=7).pack(side="left", padx=(6, 2))
        tk.Label(row1, text="até", bg=c["card"], fg=c["muted"], font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        ttk.Entry(row1, textvariable=self._esp_seg_sex_end, width=7).pack(side="left")

        # Sab-Dom
        row2 = tk.Frame(filters_frame, bg=c["card"])
        row2.pack(anchor="w")
        tk.Checkbutton(row2, text="Sáb e Dom  —  Horário completo (dia todo)",
                       variable=self._esp_sab_dom_var,
                       bg=c["card"], fg=c["text"], selectcolor=c["card2"],
                       activebackground=c["card"], font=("Segoe UI", 9),
                       cursor="hand2").pack(side="left")

        tk.Frame(card, bg=c["border"], height=1).pack(fill="x", pady=(0, 12))

        # ── Formato ───────────────────────────────────────────────────────
        self._section_label(card, "FORMATO DE SAÍDA")
        self._esp_fmt_var = tk.StringVar(value="xlsx")
        fmt_row = tk.Frame(card, bg=c["card"])
        fmt_row.pack(anchor="w", pady=(0, 4))
        ttk.Radiobutton(fmt_row, text="Excel .xlsx", variable=self._esp_fmt_var, value="xlsx").pack(side="left")
        ttk.Radiobutton(fmt_row, text="CSV .csv",    variable=self._esp_fmt_var, value="csv").pack(side="left", padx=(16, 0))

        # ── Botão gerar ───────────────────────────────────────────────────
        action = tk.Frame(parent, bg=c["card"])
        action.pack(fill="x", pady=(12, 0))
        ttk.Button(action, text="▶  Gerar Relatório Específicos",
                   style="Accent.TButton",
                   command=self._esp_start_report).pack(side="left")
        ttk.Button(action, text="Abrir pasta", style="Ghost.TButton",
                   command=self.open_reports_folder).pack(side="left", padx=(8, 0))

    def _esp_load_users(self):
        """Carrega lista de usuários únicos do banco tratado."""
        try:
            base = Path(self.output_path.get())
            db_candidates = sorted(base.rglob("base_macro_lock.duckdb"),
                                   key=lambda p: p.stat().st_mtime, reverse=True)
            if not db_candidates:
                messagebox.showwarning("Sem base", "Processe os arquivos primeiro para carregar os usuários.")
                return
            con = self._connect_db(db_candidates[0].parent)
            users = [r[0] for r in con.execute(
                "SELECT DISTINCT Usuario FROM tratado WHERE Usuario IS NOT NULL ORDER BY Usuario"
            ).fetchall()]
            con.close()
            self._esp_user_listbox.delete(0, "end")
            for u in users:
                self._esp_user_listbox.insert("end", u)
            messagebox.showinfo("Usuários", f"{len(users)} usuário(s) carregado(s).")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def _esp_start_report(self):
        threading.Thread(target=self._esp_generate_report, daemon=True).start()

    def _esp_generate_report(self):
        try:
            # Usuários selecionados
            sel_indices = self._esp_user_listbox.curselection()
            if not sel_indices:
                messagebox.showwarning("Específicos", "Selecione ao menos um usuário.")
                return
            users = [self._esp_user_listbox.get(i) for i in sel_indices]

            # Monta WHERE
            user_placeholders = ", ".join(["?" for _ in users])
            conditions = []
            params = []

            use_seg_sex = self._esp_seg_sex_var.get()
            use_sab_dom = self._esp_sab_dom_var.get()
            start_t = self._validate_time(self._esp_seg_sex_start.get(), "Horário inicial Seg-Sex")
            end_t   = self._validate_time(self._esp_seg_sex_end.get(),   "Horário final Seg-Sex")

            if use_seg_sex:
                # Seg-Sex: 18:00 → 00:00 (ou final da madrugada)
                # Lógica: hora >= 18:00 OU hora < 00:00 (madrugada)
                if start_t <= end_t:
                    # ex: 08:00–12:00 (intervalo normal)
                    cond_seg_sex = (
                        "( CAST(strftime(DataHoraTs, '%w') AS INTEGER) BETWEEN 1 AND 5 "
                        f"AND CAST(DataHoraTs AS TIME) >= TIME '{start_t}' "
                        f"AND CAST(DataHoraTs AS TIME) < TIME '{end_t}' )"
                    )
                else:
                    # ex: 18:00–00:00 (passa da meia-noite)
                    cond_seg_sex = (
                        "( CAST(strftime(DataHoraTs, '%w') AS INTEGER) BETWEEN 1 AND 5 "
                        f"AND ( CAST(DataHoraTs AS TIME) >= TIME '{start_t}' "
                        f"   OR CAST(DataHoraTs AS TIME) < TIME '{end_t}' ) )"
                    )
                conditions.append(cond_seg_sex)

            if use_sab_dom:
                conditions.append(
                    "( CAST(strftime(DataHoraTs, '%w') AS INTEGER) IN (0, 6) )"
                )

            if not conditions:
                messagebox.showwarning("Específicos", "Selecione ao menos um filtro de período.")
                return

            period_where = "(" + " OR ".join(conditions) + ")"
            where = f"Usuario IN ({user_placeholders}) AND {period_where}"
            params = list(users)

            # Banco
            base = Path(self.output_path.get())
            db_candidates = sorted(base.rglob("base_macro_lock.duckdb"),
                                   key=lambda p: p.stat().st_mtime, reverse=True)
            if not db_candidates:
                messagebox.showerror("Erro", "Nenhuma base encontrada. Processe os arquivos primeiro.")
                return
            output_dir  = db_candidates[0].parent
            reports_dir = output_dir / "relatorios"
            reports_dir.mkdir(parents=True, exist_ok=True)

            con = self._connect_db(output_dir)
            query = f"SELECT {', '.join(EXPORT_COLUMNS)} FROM tratado WHERE {where} ORDER BY Usuario, DataHoraTs"
            total = con.execute(f"SELECT COUNT(*) FROM ({query}) q", params).fetchone()[0]
            self.set_status(f"[Específicos] {total:,} registros encontrados".replace(",", "."), 50)

            stamp     = datetime.now().strftime("%Y%m%d_%H%M%S")
            fmt       = self._esp_fmt_var.get()
            out_path  = reports_dir / f"especificos_{stamp}.{fmt}"

            if fmt == "csv":
                # Para CSV com params, copia para tabela temp
                con.execute("DROP TABLE IF EXISTS _esp_tmp")
                con.execute(f"CREATE TEMP TABLE _esp_tmp AS SELECT {', '.join(EXPORT_COLUMNS)} FROM tratado WHERE {where}", params)
                con.execute(f"COPY _esp_tmp TO '{sql_path(out_path)}' (HEADER, DELIMITER ';')")
            else:
                self._export_query_to_xlsx(con, query, params, EXPORT_COLUMNS, out_path,
                                           "Específicos", total)

            con.close()
            self.set_status(f"[Específicos] Relatório gerado: {out_path}", 100)
            messagebox.showinfo("Relatório gerado",
                                f"Relatório gerado com sucesso!\n\n{out_path}\n\n{total:,} registros.".replace(",", "."))
        except Exception as e:
            self.set_status("Erro ao gerar relatório Específicos.", 0)
            messagebox.showerror("Erro", str(e))

    def select_file(self):
        """Mantido para compatibilidade; redireciona para seleção múltipla."""
        self._files_add()

    # ── Gerenciamento da lista de múltiplos arquivos ──────────────────
    def _files_refresh(self):
        self._files_lb.delete(0, "end")
        for p in self._multi_files:
            self._files_lb.insert("end", Path(p).name + "  —  " + p)
        n = len(self._multi_files)
        self._files_count_var.set(
            f"{n} arquivo(s) na fila" if n else "Nenhum arquivo selecionado"
        )

    def _files_add(self):
        paths = filedialog.askopenfilenames(
            title="Selecione os arquivos do Macro Lock",
            filetypes=[
                ("CSV e Excel", "*.csv *.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx *.xls"),
                ("Todos", "*.*"),
            ],
        )
        for p in paths:
            if p not in self._multi_files:
                self._multi_files.append(p)
        self._files_refresh()

    def _files_remove(self):
        sel = list(self._files_lb.curselection())
        for i in reversed(sel):
            del self._multi_files[i]
        self._files_refresh()

    def _files_clear(self):
        self._multi_files.clear()
        self._files_refresh()

    def select_output(self):
        path = filedialog.askdirectory(title="Escolha a pasta de saída")
        if path:
            self.output_path.set(path)

    def open_output_folder(self):
        self._open_folder(Path(self.output_path.get()))

    def open_reports_folder(self):
        self._open_folder(self._dated_output_dir() / "relatorios")

    def _dated_output_dir(self) -> Path:
        """Retorna output_path / DD-MM-AA (data de hoje), criando a pasta se necessário."""
        date_folder = datetime.now().strftime("%d-%m-%y")
        path = Path(self.output_path.get()) / date_folder
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _open_folder(self, path: Path):
        path.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)])
        else:
            subprocess.run(["xdg-open", str(path)])

    def write_log(self, msg, level="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}]  {msg}\n"
        try:
            self.log.configure(state="normal")
        except Exception:
            pass
        tag = level if level in ("ok", "warn", "error") else "info"
        self.log.insert("end", line, tag)
        self.log.see("end")
        try:
            self.root.update_idletasks()
        except Exception:
            pass

    def set_status(self, msg, progress=None):
        self.status_text.set(msg)
        if progress is not None:
            self.progress_value.set(progress)
            try:
                self._proc_pct_var.set(f"{int(progress)}%")
                self._proc_step_var.set(msg)
            except Exception:
                pass
        self.write_log(msg)

    def _set_proc_step(self, active_index: int, done: bool = False):
        """Acende os indicadores de etapa (0=Lendo, 1=Tratando, 2=Filtros, 3=Exportando, 4=Concluído)."""
        c = self.COLORS
        try:
            for i, (dot, lbl) in enumerate(self._step_labels):
                if done and i <= active_index:
                    color = c["ok"]
                    fg    = c["ok"]
                elif i < active_index:
                    color = c["accent2"]
                    fg    = c["accent"]
                elif i == active_index:
                    color = c["accent"]
                    fg    = c["accent"]
                else:
                    color = c["border"]
                    fg    = c["muted"]
                dot.itemconfig("dot", fill=color, outline=color)
                lbl.configure(fg=fg)
            self.root.update_idletasks()
        except Exception:
            pass

    def start_processing(self):
        threading.Thread(target=self.process, daemon=True).start()

    def start_report(self):
        # Lê vars Tkinter na thread principal antes de iniciar a thread de background
        _add_dash = bool(getattr(self, "report_dashboard", None) and self.report_dashboard.get())
        _fmt = self.report_format.get()
        _fname = self.output_filename.get().strip()
        threading.Thread(target=self.generate_report, args=(_add_dash, _fmt, _fname), daemon=True).start()

    def _select_import_file(self):
        path = filedialog.askopenfilename(
            title="Selecione a base tratada (CSV ou XLSX)",
            filetypes=[("CSV e Excel", "*.csv *.xlsx *.xls"), ("CSV", "*.csv"), ("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if path:
            self._import_path.set(path)

    def start_import_treated(self):
        threading.Thread(target=self.import_treated_base, daemon=True).start()

    def import_treated_base(self):
        """Importa um arquivo de base já tratada (CSV ou XLSX) para o DuckDB,
        dispensando o reprocessamento do arquivo bruto do Macro Lock."""
        import_file = Path(self._import_path.get().strip())

        if not import_file.is_file():
            messagebox.showerror("Erro", "Selecione um arquivo de base tratada antes de importar.")
            return

        output_dir = self._dated_output_dir()
        db_path = output_dir / "base_macro_lock.duckdb"

        try:
            import duckdb
        except Exception:
            messagebox.showerror("Dependência ausente", "Instale o duckdb:\npython -m pip install duckdb")
            return

        # Remove banco anterior se existir
        if db_path.exists():
            try:
                db_path.unlink()
            except Exception:
                messagebox.showerror("Erro", f"Feche qualquer relatório/base antes de importar novamente:\n{db_path}")
                return

        con = duckdb.connect(str(db_path))
        try:
            is_xlsx = import_file.suffix.lower() in (".xlsx", ".xls")
            if is_xlsx:
                self.set_status("Convertendo Excel para CSV temporário...", 5)
                read_path = self._xlsx_to_csv(import_file, output_dir)
                delim = ","
            else:
                self.set_status("Normalizando encoding...", 5)
                enc = self._detect_source_encoding(import_file)
                delim = self._detect_delimiter(import_file, enc)
                self.write_log(f"Delimitador detectado: '{delim}'")
                read_path = self._normalize_csv_to_utf8(import_file, output_dir)

            self._set_proc_step(0)
            self.set_status("Lendo base tratada...", 15)

            # Lê com header=true — arquivo tratado já tem cabeçalho
            con.execute("DROP TABLE IF EXISTS tratado_import")
            con.execute(f"""
                CREATE TABLE tratado_import AS
                SELECT * FROM read_csv(
                    '{sql_path(read_path)}',
                    header = true,
                    delim = '{delim}',
                    quote = '"',
                    escape = '"',
                    ignore_errors = true,
                    null_padding = true
                )
            """)

            raw_cols = [row[0].strip() for row in con.execute("DESCRIBE tratado_import").fetchall()]
            self.write_log(f"Colunas encontradas: {raw_cols}")

            # Verifica se parece com a base tratada pelo Umbrela Analytics
            expected = {"Data", "Hora", "Equipamento", "Usuario", "Status"}
            found = {c.lower() for c in raw_cols}
            missing = {e.lower() for e in expected} - found
            if missing:
                raise RuntimeError(
                    f"O arquivo não parece ser uma base tratada do Umbrela Analytics.\n"
                    f"Colunas esperadas não encontradas: {', '.join(missing)}\n"
                    f"Colunas presentes: {', '.join(raw_cols)}"
                )

            self._set_proc_step(1)
            self.set_status("Montando tabela tratada...", 40)

            # Mapeia colunas de forma case-insensitive
            col_map = {c.lower(): c for c in raw_cols}

            def get_col(name, default="NULL"):
                return f'"{col_map[name.lower()]}"' if name.lower() in col_map else default

            # Reconstrói DataHoraTs a partir de Data + Hora
            data_col = get_col("Data")
            hora_col = get_col("Hora")

            # Colunas opcionais
            diasemana_col   = get_col("DiaSemana", "NULL")
            periodo_col     = get_col("Periodo", "NULL")
            equip_col       = get_col("Equipamento")
            usuario_col     = get_col("Usuario")
            status_col      = get_col("Status", "NULL")
            aplicativo_col  = get_col("Aplicativo", "NULL")
            campo_col       = get_col("CampoExtra", "NULL")
            janela_col      = get_col("JanelaEvento", "NULL")
            duracao_col     = get_col("DuracaoSegundos", "NULL")
            tempo_col       = get_col("TempoFormatado", "NULL")

            con.execute("DROP TABLE IF EXISTS tratado")
            con.execute(f"""
                CREATE TABLE tratado AS
                SELECT
                    TRY_CAST(({data_col}::VARCHAR || ' ' || {hora_col}::VARCHAR) AS TIMESTAMP) AS DataHoraTs,
                    {data_col}::VARCHAR   AS Data,
                    {hora_col}::VARCHAR   AS Hora,
                    COALESCE(
                        CASE WHEN {diasemana_col} IS NULL OR {diasemana_col}::VARCHAR = '' THEN NULL ELSE {diasemana_col}::VARCHAR END,
                        CASE CAST(strftime(TRY_CAST(({data_col}::VARCHAR || ' ' || {hora_col}::VARCHAR) AS TIMESTAMP), '%w') AS INTEGER)
                            WHEN 0 THEN 'Domingo'      WHEN 1 THEN 'Segunda-feira'
                            WHEN 2 THEN 'Terça-feira'  WHEN 3 THEN 'Quarta-feira'
                            WHEN 4 THEN 'Quinta-feira' WHEN 5 THEN 'Sexta-feira'
                            WHEN 6 THEN 'Sábado'
                        END
                    ) AS DiaSemana,
                    COALESCE(
                        CASE WHEN {periodo_col} IS NULL OR {periodo_col}::VARCHAR = '' THEN NULL ELSE {periodo_col}::VARCHAR END,
                        CASE WHEN CAST(TRY_CAST(({data_col}::VARCHAR || ' ' || {hora_col}::VARCHAR) AS TIMESTAMP) AS TIME) >= TIME '18:00:00'
                              OR  CAST(TRY_CAST(({data_col}::VARCHAR || ' ' || {hora_col}::VARCHAR) AS TIMESTAMP) AS TIME) < TIME '07:00:00'
                             THEN 'Fora do Horário' ELSE 'Horário Comercial' END
                    ) AS Periodo,
                    {equip_col}::VARCHAR  AS Equipamento,
                    {usuario_col}::VARCHAR AS Usuario,
                    {status_col}::VARCHAR AS Status,
                    {aplicativo_col}::VARCHAR AS Aplicativo,
                    {campo_col}::VARCHAR  AS CampoExtra,
                    {janela_col}::VARCHAR AS JanelaEvento,
                    TRY_CAST({duracao_col} AS BIGINT) AS DuracaoSegundos,
                    {tempo_col}::VARCHAR  AS TempoFormatado
                FROM tratado_import
                WHERE TRY_CAST(({data_col}::VARCHAR || ' ' || {hora_col}::VARCHAR) AS TIMESTAMP) IS NOT NULL
            """)

            con.execute("DROP TABLE IF EXISTS tratado_import")

            total = con.execute("SELECT COUNT(*) FROM tratado").fetchone()[0]
            self._set_proc_step(4, done=True)
            self.set_status(f"Base importada com sucesso! {total:,} registros carregados.".replace(",", "."), 100)
            self.write_log(f"Importação concluída: {total:,} linhas".replace(",", "."))

            try:
                self.refresh_dashboard()
            except Exception:
                pass

            messagebox.showinfo(
                "Importação concluída",
                f"Base tratada importada com sucesso!\n\n"
                f"Registros: {total:,}\n"
                f"Agora vá para a aba Relatórios.\n\n"
                f"Base salva em:\n{db_path}".replace(",", ".")
            )
        except Exception as e:
            self.set_status("Erro na importação.", 0)
            self.write_log(str(e))
            messagebox.showerror("Erro na importação", str(e))
        finally:
            try:
                con.close()
            except Exception:
                pass

    def _validate_time(self, value: str, field_name: str) -> str:
        value = value.strip()
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(value, fmt).strftime("%H:%M:%S")
            except ValueError:
                pass
        raise ValueError(f"{field_name} inválido. Use HH:MM ou HH:MM:SS. Exemplo: 18:00")

    def _validate_date_optional(self, value: str, field_name: str):
        value = value.strip()
        if not value:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        raise ValueError(f"{field_name} inválida. Use DD/MM/AAAA ou AAAA-MM-DD.")

    def _connect_db(self, output_dir: Path):
        import duckdb
        db_path = output_dir / "base_macro_lock.duckdb"
        if not db_path.exists():
            raise RuntimeError(f"Base tratada não encontrada. Primeiro processe o CSV.\n\nEsperado em:\n{db_path}")
        con = duckdb.connect(str(db_path))
        try:
            con.execute("SELECT 1 FROM tratado LIMIT 1")
        except Exception:
            con.close()
            raise RuntimeError("A base existe, mas a tabela tratada não foi encontrada. Reprocesse o CSV.")
        return con

    def _detect_delimiter(self, csv_path: Path, encoding: str) -> str:
        """Detecta o delimitador real do CSV (vírgula, ponto-e-vírgula ou tab)."""
        import csv as csv_mod
        with open(csv_path, "r", encoding=encoding, errors="replace", newline="") as f:
            sample = f.read(8192)
        try:
            dialect = csv_mod.Sniffer().sniff(sample, delimiters=",;\t|")
            return dialect.delimiter
        except Exception:
            # fallback: conta ocorrências na primeira linha
            first_line = sample.split("\n")[0]
            counts = {d: first_line.count(d) for d in (",", ";", "\t", "|")}
            return max(counts, key=counts.get)

    def _detect_source_encoding(self, csv_path: Path) -> str:
        """Detecta uma codificação provável para converter o CSV para UTF-8.
        Macro Lock pode exportar arquivos em ANSI/Windows-1252 ou UTF-8.
        """
        sample = csv_path.read_bytes()[:1024 * 1024]
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                sample.decode(enc)
                return enc
            except UnicodeDecodeError:
                continue
        return "cp1252"

    def _xlsx_to_csv(self, xlsx_path: Path, output_dir: Path) -> Path:
        """Converte xlsx/xls para CSV UTF-8 temporário para o DuckDB ler."""
        try:
            import openpyxl
        except ImportError:
            raise RuntimeError("Instale o openpyxl:\npython -m pip install openpyxl")
        self.write_log("Convertendo Excel para CSV temporário...")
        wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
        ws = wb.active
        out = output_dir / "_entrada_macro_lock_utf8.csv"
        import csv as csv_mod
        with open(out, "w", encoding="utf-8", newline="") as f:
            writer = csv_mod.writer(f, delimiter=",")
            for row in ws.iter_rows(values_only=True):
                writer.writerow(["" if v is None else str(v) for v in row])
        wb.close()
        return out

    def _normalize_csv_to_utf8(self, csv_path: Path, output_dir: Path) -> Path:
        """Cria uma cópia UTF-8 segura para o DuckDB ler sem erro de encoding.
        Faz leitura em streaming para suportar arquivos grandes.
        """
        encoding = self._detect_source_encoding(csv_path)
        normalized = output_dir / "_entrada_macro_lock_utf8.csv"
        self.write_log(f"Codificação detectada/conversão: {encoding} -> utf-8")
        with open(csv_path, "r", encoding=encoding, errors="replace", newline="") as src, \
             open(normalized, "w", encoding="utf-8", newline="") as dst:
            while True:
                chunk = src.read(1024 * 1024 * 8)
                if not chunk:
                    break
                dst.write(chunk)
        return normalized

    def _read_one_file_to_bruto(self, con, csv_path: Path, output_dir: Path, idx: int):
        """Lê um arquivo (CSV ou XLSX) e faz INSERT INTO bruto com as colunas padronizadas.
        Na primeira chamada (idx==0) cria a tabela; nas seguintes, insere."""
        is_xlsx = csv_path.suffix.lower() in (".xlsx", ".xls")
        if is_xlsx:
            self.set_status(f"[{idx+1}] Convertendo Excel para CSV...", None)
            read_path = self._xlsx_to_csv(csv_path, output_dir)
            delim = ","
        else:
            self.set_status(f"[{idx+1}] Normalizando encoding: {csv_path.name}", None)
            enc = self._detect_source_encoding(csv_path)
            delim = self._detect_delimiter(csv_path, enc)
            self.write_log(f"[{idx+1}] Delimitador: '{delim}'")
            read_path = self._normalize_csv_to_utf8(csv_path, output_dir)

        con.execute("DROP TABLE IF EXISTS bruto_raw_tmp")
        con.execute(f"""
            CREATE TEMP TABLE bruto_raw_tmp AS
            SELECT * FROM read_csv(
                '{sql_path(read_path)}',
                header = true,
                delim = '{delim}',
                quote = '"',
                escape = '"',
                ignore_errors = true,
                null_padding = true
            )
        """)
        raw_cols = [row[0] for row in con.execute("DESCRIBE bruto_raw_tmp").fetchall()]
        self.write_log(f"[{idx+1}] Colunas detectadas: {raw_cols}")

        # Mapeamento automático por nome de coluna (todos os tipos de relatório do Macro Lock)
        col_lower = {c.strip().lower(): f'"{c}"' for c in raw_cols}

        def find_col(*candidates):
            """Retorna a expressão SQL da primeira coluna encontrada (case-insensitive)."""
            for name in candidates:
                if name.lower() in col_lower:
                    return col_lower[name.lower()]
            return "''"

        # DataHora — sempre col 0
        c_datahora  = find_col("data/hora", "data hora", "datahora", "horário", "horario", raw_cols[0] if raw_cols else "")
        # Equipamento
        c_equip     = find_col("computador", "equipamento", "máquina", "maquina", "pc", "workstation")
        # Usuário
        c_usuario   = find_col("usuário", "usuario", "user", "utilizador")
        # Status / Ação / Tipo — col(3) para qualquer formato
        c_status    = find_col("status", "ação", "acao", "tipo", "evento", "type", "action")
        # Aplicativo / Categoria / Host / Programa / Software / Dispositivo / Arquivo
        c_aplicativo = find_col("aplicativo", "categoria", "programa", "host", "software",
                                "dispositivo", "arquivo", "produtividade", "website", "app")
        # JanelaEvento / Atividade / URL / Título / Janela / Descrição
        c_janela    = find_col("atividade", "url", "título", "titulo", "janela", "janelae vento",
                               "janelae", "janelaevento", "descrição", "descricao", "detalhes", "detail")
        # Duração
        c_duracao   = find_col("duração", "duracao", "duration", "tempo", "segundos", "time")
        # Tempo formatado
        c_tempo_fmt = find_col("tempo formatado", "tempoformatado", "formatted", "hh:mm:ss")
        # CampoExtra — colunas não mapeadas em nenhuma das anteriores
        mapped = {c_datahora, c_equip, c_usuario, c_status, c_aplicativo, c_janela, c_duracao, c_tempo_fmt}
        extras = [f'"{c}"' for c in raw_cols if f'"{c}"' not in mapped and f'"{c}"' != "''"]
        c_extra = f"CONCAT_WS(' | ', {', '.join(extras)})" if extras else "''"

        # Detecta tipo para log
        tipo_csv = "Internet" if "host" in col_lower or "url" in col_lower else \
                   "Aplicativos" if "aplicativo" in col_lower else \
                   "Dispositivos" if "dispositivo" in col_lower else \
                   "Arquivos" if "arquivo" in col_lower else \
                   "Software" if "software" in col_lower else \
                   "Produtividade" if "produtividade" in col_lower else \
                   "Inventário" if "inventário" in col_lower or "inventario" in col_lower else \
                   "Atividades"
        self.write_log(f"[{idx+1}] Tipo detectado: {tipo_csv}")

        # DataHora col de referência para filtro de linhas vazias
        c_dh_ref = c_datahora if c_datahora != "''" else f'"{raw_cols[0]}"' if raw_cols else "''"

        select_sql = f"""
            SELECT
                {c_datahora}::VARCHAR AS DataHora,
                {c_equip}::VARCHAR AS Equipamento,
                {c_usuario}::VARCHAR AS Usuario,
                {c_status}::VARCHAR AS Status,
                {c_aplicativo}::VARCHAR AS Aplicativo,
                ({c_extra})::VARCHAR AS CampoExtra,
                {c_janela}::VARCHAR AS JanelaEvento,
                TRY_CAST({c_duracao} AS BIGINT) AS DuracaoSegundos,
                {c_tempo_fmt}::VARCHAR AS TempoFormatado
            FROM bruto_raw_tmp
            WHERE {c_dh_ref} IS NOT NULL
              AND TRIM({c_dh_ref}::VARCHAR) != ''
        """
        if idx == 0:
            con.execute("DROP TABLE IF EXISTS bruto")
            con.execute(f"CREATE TABLE bruto AS {select_sql}")
        else:
            con.execute(f"INSERT INTO bruto {select_sql}")
        con.execute("DROP TABLE IF EXISTS bruto_raw_tmp")
        n = con.execute("SELECT COUNT(*) FROM bruto").fetchone()[0]
        self.write_log(f"[{idx+1}] {csv_path.name}: acumulado {n:,} linhas".replace(",", "."))

    def process(self):
        # Usa lista múltipla; fallback para file_path se lista vazia
        files = list(self._multi_files)
        if not files:
            single = self.file_path.get().strip()
            if single and Path(single).is_file():
                files = [single]
        if not files:
            messagebox.showerror("Erro", "Adicione pelo menos um arquivo CSV na lista da aba 'Processar'.")
            return

        output_dir = self._dated_output_dir()
        try:
            import duckdb
        except Exception:
            messagebox.showerror("Dependência ausente", "Instale o duckdb:\npython -m pip install duckdb")
            return

        db_path = output_dir / "base_macro_lock.duckdb"
        if db_path.exists():
            try:
                db_path.unlink()
            except Exception:
                messagebox.showerror("Erro", f"Feche qualquer relatório/base antes de processar novamente:\n{db_path}")
                return

        con = duckdb.connect(str(db_path))
        try:
            self._set_proc_step(0)
            self.set_status(f"Lendo {len(files)} arquivo(s)...", 5)

            for i, fpath in enumerate(files):
                self._read_one_file_to_bruto(con, Path(fpath), output_dir, i)
                prog = 5 + int((i + 1) / len(files) * 15)
                self.progress_value.set(prog)
                self.root.update_idletasks()

            total = con.execute("SELECT COUNT(*) FROM bruto").fetchone()[0]
            self.set_status(f"Total lido ({len(files)} arquivo(s)): {total:,} linhas".replace(",", "."), 20)

            terms = [t.strip() for t in self.exclude_terms.get().split(",") if t.strip()]
            where_parts, params = [], []
            for term in terms:
                where_parts.append("""
                    NOT (
                        upper(coalesce(Status, '')) LIKE upper(?)
                        OR upper(coalesce(Aplicativo, '')) LIKE upper(?)
                        OR upper(coalesce(JanelaEvento, '')) LIKE upper(?)
                    )
                """)
                like = f"%{term}%"
                params.extend([like, like, like])
            filter_sql = " AND ".join(where_parts) if where_parts else "1=1"

            self._set_proc_step(1)
            self.set_status("Tratando base: separando Data/Hora, dia da semana e removendo eventos...", 45)
            con.execute("DROP TABLE IF EXISTS tratado")
            con.execute(f"""
                CREATE TABLE tratado AS
                SELECT
                    DataHoraTs,
                    strftime(DataHoraTs, '%d-%m-%Y') AS Data,
                    strftime(DataHoraTs, '%H:%M:%S') AS Hora,
                    CASE CAST(strftime(DataHoraTs, '%w') AS INTEGER)
                        WHEN 0 THEN 'Domingo'
                        WHEN 1 THEN 'Segunda-feira'
                        WHEN 2 THEN 'Terça-feira'
                        WHEN 3 THEN 'Quarta-feira'
                        WHEN 4 THEN 'Quinta-feira'
                        WHEN 5 THEN 'Sexta-feira'
                        WHEN 6 THEN 'Sábado'
                    END AS DiaSemana,
                    CASE
                        WHEN CAST(DataHoraTs AS TIME) >= TIME '18:00:00'
                          OR CAST(DataHoraTs AS TIME) < TIME '07:00:00'
                        THEN 'Fora do Horário'
                        ELSE 'Horário Comercial'
                    END AS Periodo,
                    Equipamento,
                    Usuario,
                    Status,
                    Aplicativo,
                    CampoExtra,
                    JanelaEvento,
                    DuracaoSegundos,
                    CASE
                        WHEN DuracaoSegundos IS NOT NULL AND DuracaoSegundos >= 0
                        THEN printf('%02d:%02d:%02d',
                            CAST(DuracaoSegundos / 3600 AS INTEGER),
                            CAST((DuracaoSegundos % 3600) / 60 AS INTEGER),
                            CAST(DuracaoSegundos % 60 AS INTEGER))
                        ELSE TempoFormatado
                    END AS Duracao
                FROM (
                    SELECT
                        COALESCE(
                            TRY_STRPTIME(DataHora, '%d/%m/%Y %H:%M:%S'),
                            TRY_STRPTIME(DataHora, '%d/%m/%Y %H:%M'),
                            TRY_STRPTIME(DataHora, '%Y-%m-%d %H:%M:%S'),
                            TRY_STRPTIME(DataHora, '%Y-%m-%d %H:%M'),
                            TRY_CAST(DataHora AS TIMESTAMP)
                        ) AS DataHoraTs,
                        Equipamento,
                        Usuario,
                        Status,
                        Aplicativo,
                        CampoExtra,
                        JanelaEvento,
                        DuracaoSegundos,
                        TempoFormatado
                    FROM bruto
                ) base
                WHERE DataHoraTs IS NOT NULL
                  AND {filter_sql}
                ORDER BY Equipamento, Usuario, DataHoraTs
            """, params)

            final_total = con.execute("SELECT COUNT(*) FROM tratado").fetchone()[0]
            removidas = total - final_total
            self.set_status(f"Base tratada criada. Mantidas: {final_total:,} | Removidas: {removidas:,}".replace(",", "."), 70)

            self._set_proc_step(2)
            # Aplicar correções de nomes definidas na aba "Corrigir Nomes"
            for col_field, mapping in self.rename_maps.items():
                for old_name, new_name in mapping.items():
                    if old_name and new_name and old_name != new_name:
                        con.execute(
                            f"UPDATE tratado SET {col_field} = ? WHERE {col_field} = ?",
                            [new_name, old_name],
                        )
                        self.write_log(f"Renomeado [{col_field}]: '{old_name}' → '{new_name}'")

            self._set_proc_step(3)
            query_completa = f"SELECT {', '.join(EXPORT_COLUMNS)} FROM tratado ORDER BY Equipamento, Usuario, DataHoraTs"
            export_fmt = self.base_export_format.get()
            if export_fmt in ("csv", "csv_xlsx"):
                completo_path = output_dir / "macro_lock_base_tratada_completa.csv"
                self.set_status("Exportando CSV completo tratado...", 82)
                con.execute(f"COPY ({query_completa}) TO '{sql_path(completo_path)}' (HEADER, DELIMITER ';')")
            if export_fmt in ("xlsx", "csv_xlsx"):
                completo_xlsx = output_dir / "macro_lock_base_tratada_completa.xlsx"
                self.set_status("Exportando Excel completo tratado com finais de semana em vermelho...", 88)
                self._export_query_to_xlsx(con, query_completa, [], EXPORT_COLUMNS, completo_xlsx, "Base tratada completa", final_total, color_weekend=True)

            self._create_base_summary(con, output_dir)
            self._set_proc_step(4, done=True)
            self.set_status("Base finalizada. Agora use a aba Relatórios para gerar filtros específicos.", 100)
            try:
                self.refresh_dashboard()
            except Exception:
                pass
            nf = len(files)
            messagebox.showinfo(
                "Concluído",
                f"Base tratada criada com sucesso!\n\n"
                f"{nf} arquivo(s) consolidado(s).\n"
                f"Registros finais: {final_total:,}\n\n"
                f"Saída:\n{output_dir}".replace(",", ".")
            )
        except Exception as e:
            self.set_status("Erro durante o processamento.", 0)
            self.write_log(str(e))
            messagebox.showerror("Erro", str(e))
        finally:
            try:
                con.close()
            except Exception:
                pass

    def _create_base_summary(self, con, output_dir: Path):
        resumo_path = output_dir / "resumo_por_equipamento_usuario.csv"
        con.execute(f"""
            COPY (
                SELECT
                    Equipamento,
                    Usuario,
                    COUNT(*) AS Registros,
                    SUM(coalesce(DuracaoSegundos, 0)) AS Duracao_Total_Segundos,
                    CAST(SUM(coalesce(DuracaoSegundos, 0)) / 3600.0 AS DECIMAL(18,2)) AS Duracao_Total_Horas
                FROM tratado
                GROUP BY Equipamento, Usuario
                ORDER BY Equipamento, Usuario
            ) TO '{sql_path(resumo_path)}' (HEADER, DELIMITER ';')
        """)
        self.write_log(f"Resumo base gerado: {resumo_path}")

    def _report_where(self, report_type: str, start: str, end: str):
        params = []
        if report_type == "fora_horario":
            if start <= end:
                # Intervalo no mesmo dia (ex: 18:00 → 23:59)
                where = f"(CAST(DataHoraTs AS TIME) >= TIME '{start}' AND CAST(DataHoraTs AS TIME) <= TIME '{end}')"
            else:
                # Intervalo que cruza meia-noite (ex: 18:00 → 07:00)
                where = f"(CAST(DataHoraTs AS TIME) >= TIME '{start}' OR CAST(DataHoraTs AS TIME) < TIME '{end}')"
        elif report_type == "fim_semana":
            where = "CAST(strftime(DataHoraTs, '%w') AS INTEGER) IN (0, 6)"
        elif report_type == "fora_horario_fim_semana":
            if start <= end:
                hora_cond = f"(CAST(DataHoraTs AS TIME) >= TIME '{start}' AND CAST(DataHoraTs AS TIME) <= TIME '{end}')"
            else:
                hora_cond = f"(CAST(DataHoraTs AS TIME) >= TIME '{start}' OR CAST(DataHoraTs AS TIME) < TIME '{end}')"
            where = f"({hora_cond} OR CAST(strftime(DataHoraTs, '%w') AS INTEGER) IN (0, 6))"
        elif report_type == "por_usuario_equipamento":
            where = "1=1"
        elif report_type == "intervalo_horario":
            if start <= end:
                # intervalo direto, ex: 00:00 até 09:00
                where = f"(CAST(DataHoraTs AS TIME) >= TIME '{start}' AND CAST(DataHoraTs AS TIME) <= TIME '{end}')"
            else:
                # intervalo que cruza meia-noite, ex: 22:00 até 06:00
                where = f"(CAST(DataHoraTs AS TIME) >= TIME '{start}' OR CAST(DataHoraTs AS TIME) <= TIME '{end}')"
        elif report_type == "blacklist":
            if not self.blacklist_terms:
                raise ValueError("A Blacklist está vazia. Cadastre ao menos um termo na aba 'Blacklist'.")
            clauses = []
            for term in self.blacklist_terms:
                like = term if "%" in term else f"%{term.replace('*', '%')}%"
                clauses.append(f"upper(coalesce(Aplicativo,'')) LIKE upper(?)")
                params.append(like)
                clauses.append(f"upper(coalesce(JanelaEvento,'')) LIKE upper(?)")
                params.append(like)
            where = "(" + " OR ".join(clauses) + ")"
        elif report_type in ("ociosidade_diario", "ociosidade_mensal"):
            # Filtra somente dias úteis; as janelas de horário são aplicadas
            # como CASE dentro do SELECT para não excluir dias com pouca atividade
            where = "CAST(strftime(DataHoraTs, '%w') AS INTEGER) BETWEEN 1 AND 5"
        else:
            raise ValueError("Tipo de relatório inválido.")

        user = self.report_user.get().strip()
        equip = self.report_equipment.get().strip()
        date_start = self._validate_date_optional(self.report_date_start.get(), "Data inicial")
        date_end = self._validate_date_optional(self.report_date_end.get(), "Data final")
        if user:
            where += " AND upper(coalesce(Usuario, '')) LIKE upper(?)"
            params.append(f"%{user}%")
        if equip:
            where += " AND upper(coalesce(Equipamento, '')) LIKE upper(?)"
            params.append(f"%{equip}%")
        if date_start:
            where += " AND DataHoraTs >= CAST(? AS TIMESTAMP)"
            params.append(date_start + " 00:00:00")
        if date_end:
            where += " AND DataHoraTs < CAST(? AS TIMESTAMP) + INTERVAL 1 DAY"
            params.append(date_end + " 00:00:00")
        return where, params

    def generate_report(self, _add_dash: bool = False, _fmt: str = "xlsx", _fname: str = ""):
        # Busca banco mais recente
        base = Path(self.output_path.get())
        db_candidates = sorted(base.rglob("base_macro_lock.duckdb"), key=lambda p: p.stat().st_mtime, reverse=True)
        output_dir = db_candidates[0].parent if db_candidates else self._dated_output_dir()
        reports_dir = output_dir / "relatorios"
        reports_dir.mkdir(parents=True, exist_ok=True)
        try:
            start = self._validate_time(self.night_start.get(), "Horário inicial")
            end = self._validate_time(self.night_end.get(), "Horário final")
            self.night_start.set(start)
            self.night_end.set(end)
            report_type = self.report_type.get()
            report_name = REPORTS.get(report_type, report_type)
            con = self._connect_db(output_dir)
            where, params = self._report_where(report_type, start, end)

            self.set_status(f"Gerando relatório: {report_name}...", 20)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = safe_name(_fname) if _fname else safe_name(report_type + "_" + stamp)

            if report_type in ("ociosidade_diario", "ociosidade_mensal"):
                out_path = reports_dir / f"{base_name}.{'csv' if _fmt == 'csv' else 'xlsx'}"
                self._generate_ociosidade_report(con, report_type, where, params, out_path, report_name)
                self.set_status(f"Relatório de ociosidade gerado: {out_path}", 100)
                messagebox.showinfo("Relatório gerado", f"Relatório gerado com sucesso.\n\n{out_path}")
                return

            if report_type == "por_usuario_equipamento":
                query = f"""
                    SELECT
                        Equipamento,
                        Usuario,
                        COUNT(*) AS Registros,
                        MIN(DataHoraTs)::VARCHAR AS Primeiro_Evento,
                        MAX(DataHoraTs)::VARCHAR AS Ultimo_Evento,
                        SUM(coalesce(DuracaoSegundos, 0)) AS Duracao_Total_Segundos,
                        CAST(SUM(coalesce(DuracaoSegundos, 0)) / 3600.0 AS DECIMAL(18,2)) AS Duracao_Total_Horas
                    FROM tratado
                    WHERE {where}
                    GROUP BY Equipamento, Usuario
                    ORDER BY Equipamento, Usuario
                """
                columns = ["Equipamento", "Usuario", "Registros", "Primeiro_Evento", "Ultimo_Evento", "Duracao_Total_Segundos", "Duracao_Total_Horas"]
            else:
                query = f"SELECT {', '.join(EXPORT_COLUMNS)} FROM tratado WHERE {where} ORDER BY Equipamento, Usuario, DataHoraTs"
                columns = EXPORT_COLUMNS

            total = con.execute(f"SELECT COUNT(*) FROM ({query}) q", params).fetchone()[0]
            self.set_status(f"Registros encontrados: {total:,}".replace(",", "."), 45)

            if _fmt == "csv":
                out_path = reports_dir / f"{base_name}.csv"
                con.execute(f"COPY ({query}) TO '{sql_path(out_path)}' (HEADER, DELIMITER ';')", params)
            else:
                out_path = reports_dir / f"{base_name}.xlsx"
                self._export_query_to_xlsx(con, query, params, columns, out_path, report_name, total, add_dashboard=_add_dash)

            self.set_status(f"Relatório gerado: {out_path}", 100)
            messagebox.showinfo("Relatório gerado", f"Relatório gerado com sucesso.\n\n{out_path}")
        except Exception as e:
            self.set_status("Erro ao gerar relatório.", 0)
            self.write_log(str(e))
            messagebox.showerror("Erro", str(e))
        finally:
            try:
                con.close()
            except Exception:
                pass

    # ── helper: monta tupla final combinando main + gap + upload ─────────
    @staticmethod
    def _oci_merge_row(raw, columns, gap_dict, upload_set, is_diario):
        """
        Recebe a tupla da q_main e injeta Interpretacao + Maior_Ocioso_Horas
        nos índices corretos, juntando com gap_dict e upload_set.
        """
        # índices na q_main (antes de adicionar as colunas extras)
        if is_diario:
            # raw: Data, DiaSemana, Usuario, Ocioso_Horas, Perc_Ativo, Perc_Ocioso
            data_v   = str(raw[0])
            usuario  = str(raw[2])
            perc_oc  = float(raw[5]) if raw[5] is not None else 0.0
        else:
            # raw: AnoMes, Usuario, Dias_Uteis, Ocioso_Horas, Perc_Ativo, Perc_Ocioso
            data_v   = str(raw[0])
            usuario  = str(raw[1])
            perc_oc  = float(raw[5]) if raw[5] is not None else 0.0

        is_upload = (data_v, usuario) in upload_set
        maior_gap_h = round(gap_dict.get((data_v, usuario), 0.0) / 3600.0, 2)

        if perc_oc < 25:
            interp = "Baixa ociosidade"
        elif is_upload and perc_oc >= 25:
            interp = "Aguardando upload/download"
        elif perc_oc < 60:
            interp = "Atenção"
        else:
            interp = "Alta ociosidade"

        return tuple(raw) + (interp, maior_gap_h)

    def _generate_ociosidade_report(self, con, report_type: str, where: str, params: list, out_path: Path, report_name: str):
        """Gera relatório de ociosidade simplificado (diário ou mensal) + aba de gráfico."""
        is_diario = report_type == "ociosidade_diario"

        # Expressão CASE para somar apenas duração dentro das janelas comerciais
        case_ativo = (
            "SUM(CASE WHEN "
            "(CAST(DataHoraTs AS TIME) >= TIME '08:00:00' AND CAST(DataHoraTs AS TIME) < TIME '12:00:00') "
            "OR (CAST(DataHoraTs AS TIME) >= TIME '13:00:00' AND CAST(DataHoraTs AS TIME) < TIME '17:00:00') "
            "THEN COALESCE(DuracaoSegundos, 0) ELSE 0 END)"
        )

        # ── Tabela temporária com eventos filtrados pelo WHERE do usuário ──
        # Inclui JanelaEvento e Aplicativo para detectar uploads/downloads.
        con.execute("DROP TABLE IF EXISTS oci_tmp")
        con.execute(f"""
            CREATE TEMP TABLE oci_tmp AS
            SELECT Data, DiaSemana, Usuario, DataHoraTs, DuracaoSegundos,
                   COALESCE(JanelaEvento, '') AS JanelaEvento,
                   COALESCE(Aplicativo,   '') AS Aplicativo
            FROM tratado
            WHERE {where}
        """, params)

        # ── Query 1: agregação principal (simples, sem CTEs) ─────────────
        if is_diario:
            q_main = f"""
                SELECT Data, DiaSemana, Usuario,
                    CAST(GREATEST(0, {BUSINESS_SECONDS_PER_DAY} - CAST({case_ativo} AS BIGINT)) / 3600.0
                         AS DECIMAL(10,2)) AS Ocioso_Horas,
                    CAST(LEAST(100.0, CAST({case_ativo} AS BIGINT) * 100.0 / {BUSINESS_SECONDS_PER_DAY})
                         AS DECIMAL(10,1)) AS Perc_Ativo,
                    CAST(GREATEST(0.0, ({BUSINESS_SECONDS_PER_DAY} - CAST({case_ativo} AS BIGINT))
                         * 100.0 / {BUSINESS_SECONDS_PER_DAY}) AS DECIMAL(10,1)) AS Perc_Ocioso
                FROM oci_tmp
                GROUP BY Data, DiaSemana, Usuario
                ORDER BY Data, Usuario
            """
            columns = ["Data", "DiaSemana", "Usuario",
                       "Ocioso_Horas", "Perc_Ativo", "Perc_Ocioso",
                       "Interpretacao", "Maior_Ocioso_Horas"]
            grp_key = lambda r: (r[0], r[2])   # (Data, Usuario)
        else:
            q_main = f"""
                SELECT strftime(DataHoraTs, '%Y-%m') AS AnoMes, Usuario,
                    COUNT(DISTINCT Data) AS Dias_Uteis,
                    CAST(GREATEST(0, COUNT(DISTINCT Data) * {BUSINESS_SECONDS_PER_DAY}
                         - CAST({case_ativo} AS BIGINT)) / 3600.0 AS DECIMAL(10,2)) AS Ocioso_Horas,
                    CAST(LEAST(100.0, CAST({case_ativo} AS BIGINT) * 100.0
                         / NULLIF(COUNT(DISTINCT Data) * {BUSINESS_SECONDS_PER_DAY}, 0))
                         AS DECIMAL(10,1)) AS Perc_Ativo,
                    CAST(GREATEST(0.0, (COUNT(DISTINCT Data) * {BUSINESS_SECONDS_PER_DAY}
                         - CAST({case_ativo} AS BIGINT)) * 100.0
                         / NULLIF(COUNT(DISTINCT Data) * {BUSINESS_SECONDS_PER_DAY}, 0))
                         AS DECIMAL(10,1)) AS Perc_Ocioso
                FROM oci_tmp
                GROUP BY strftime(DataHoraTs, '%Y-%m'), Usuario
                ORDER BY strftime(DataHoraTs, '%Y-%m'), Usuario
            """
            columns = ["AnoMes", "Usuario", "Dias_Uteis",
                       "Ocioso_Horas", "Perc_Ativo", "Perc_Ocioso",
                       "Interpretacao", "Maior_Ocioso_Horas"]
            grp_key = lambda r: (r[0], r[1])   # (AnoMes, Usuario)

        # ── Query 2: maior gap entre eventos consecutivos em horas comerciais
        q_gap = """
            SELECT Data, Usuario,
                MAX(GREATEST(0,
                    EPOCH(DataHoraTs) - EPOCH(prev_ts)
                    - CASE
                        WHEN CAST(prev_ts AS TIME) < TIME '12:00:00'
                         AND CAST(DataHoraTs AS TIME) >= TIME '13:00:00'
                        THEN 3600 ELSE 0
                      END
                )) AS maior_gap_s
            FROM (
                SELECT Data, Usuario, DataHoraTs,
                    LAG(DataHoraTs) OVER (PARTITION BY Data, Usuario ORDER BY DataHoraTs) AS prev_ts
                FROM oci_tmp
                WHERE (CAST(DataHoraTs AS TIME) >= TIME '08:00:00'
                       AND CAST(DataHoraTs AS TIME) < TIME '12:00:00')
                   OR (CAST(DataHoraTs AS TIME) >= TIME '13:00:00'
                       AND CAST(DataHoraTs AS TIME) < TIME '17:00:00')
            ) sub
            WHERE prev_ts IS NOT NULL
            GROUP BY Data, Usuario
        """

        # ── Query 3: flag de upload/download ─────────────────────────────
        _TRANSFER_KEYS = [
            "upload", "download", "carregando", "baixando", "enviando",
            "transferindo", "aguardando arquivo", "sending", "receiving",
            "uploading", "downloading",
        ]
        transfer_cond = " OR ".join(
            f"LOWER(JanelaEvento) LIKE '%{k}%' OR LOWER(Aplicativo) LIKE '%{k}%'"
            for k in _TRANSFER_KEYS
        )
        q_upload = f"""
            SELECT Data, Usuario
            FROM oci_tmp
            WHERE {transfer_cond}
            GROUP BY Data, Usuario
        """

        # ── Executa as 3 queries separadas e junta em Python ─────────────
        main_rows = con.execute(q_main).fetchall()
        total = len(main_rows)
        self.write_log(f"Ociosidade: {total} linha(s) encontradas para exportar.")
        self.set_status(f"Registros de ociosidade: {total:,}".replace(",", "."), 40)

        if total == 0:
            try:
                diag = con.execute(
                    "SELECT COUNT(*) FROM oci_tmp "
                    "WHERE CAST(strftime(DataHoraTs, '%w') AS INTEGER) BETWEEN 1 AND 5"
                ).fetchone()[0]
                self.write_log(f"Diagnóstico: {diag} evento(s) em dias úteis na base.")
            except Exception:
                pass
            messagebox.showwarning(
                "Sem dados",
                "Nenhum registro encontrado para o relatório de ociosidade.\n\n"
                "Verifique:\n"
                "• A base contém eventos de segunda a sexta-feira?\n"
                "• Os filtros de Data ou Usuário estão muito restritos?\n"
                "• A coluna DuracaoSegundos está preenchida na base?\n\n"
                "Confira o LOG para mais detalhes."
            )
            return

        # ── Executa gap e upload e monta dicts ───────────────────────────
        gap_dict = {}       # chave: (Data_ou_AnoMes, Usuario) -> segundos (float)
        upload_set = set()  # chave: (Data_ou_AnoMes, Usuario)
        try:
            raw_gap = con.execute(q_gap).fetchall()
            if is_diario:
                for gdata, gusu, mgap in raw_gap:
                    gap_dict[(str(gdata), str(gusu))] = float(mgap) if mgap else 0.0
            else:
                # agrega o maior gap por (AnoMes, Usuario)
                for gdata, gusu, mgap in raw_gap:
                    anomes = str(gdata)[:7]  # "YYYY-MM"
                    key = (anomes, str(gusu))
                    v = float(mgap) if mgap else 0.0
                    if v > gap_dict.get(key, 0.0):
                        gap_dict[key] = v
        except Exception as e:
            self.write_log(f"Aviso: query de gap falhou ({e}). Maior_Ocioso_Horas será 0.")

        try:
            for udata, uusu in con.execute(q_upload).fetchall():
                if is_diario:
                    upload_set.add((str(udata), str(uusu)))
                else:
                    upload_set.add((str(udata)[:7], str(uusu)))
        except Exception as e:
            self.write_log(f"Aviso: query de upload falhou ({e}). Detecção de transferência desativada.")

        # Ranking por usuário (query simples, sem CTE)
        q_ranking = f"""
            SELECT Usuario,
                CAST(SUM(GREATEST(0.0, {BUSINESS_SECONDS_PER_DAY} - ativo_s)) / 3600.0 AS DECIMAL(10,2))
                    AS Total_Ocioso_Horas
            FROM (
                SELECT Data, Usuario,
                    CAST(SUM(CASE
                        WHEN (CAST(DataHoraTs AS TIME) >= TIME '08:00:00'
                              AND CAST(DataHoraTs AS TIME) < TIME '12:00:00')
                          OR (CAST(DataHoraTs AS TIME) >= TIME '13:00:00'
                              AND CAST(DataHoraTs AS TIME) < TIME '17:00:00')
                        THEN COALESCE(DuracaoSegundos, 0) ELSE 0
                    END) AS DOUBLE) AS ativo_s
                FROM oci_tmp
                GROUP BY Data, Usuario
            ) sub
            GROUP BY Usuario
            ORDER BY Total_Ocioso_Horas DESC
            LIMIT 20
        """

        if self.report_format.get() == "csv":
            # Monta linhas para CSV manualmente
            import csv, io
            buf = io.StringIO()
            writer = csv.writer(buf, delimiter=";")
            writer.writerow(columns)
            for r in main_rows:
                writer.writerow(self._oci_merge_row(r, columns, gap_dict, upload_set, is_diario))
            with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                f.write(buf.getvalue())
            self.write_log(f"CSV de ociosidade exportado: {out_path}")
            return

        # ── Excel com formatação + gráfico ──────────────────────────────
        try:
            import xlsxwriter
        except Exception:
            raise RuntimeError("Instale o xlsxwriter:\npython -m pip install xlsxwriter")

        if out_path.exists():
            out_path.unlink()
        workbook = xlsxwriter.Workbook(str(out_path))

        # Formatos
        title_fmt  = workbook.add_format({"bold": True, "font_size": 14, "font_color": "#1F4E78"})
        hdr_fmt    = workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                                          "border": 1, "align": "center", "valign": "vcenter",
                                          "text_wrap": True})
        normal_fmt = workbook.add_format({"border": 1, "valign": "vcenter"})
        num2_fmt   = workbook.add_format({"border": 1, "num_format": "0.00", "valign": "vcenter"})
        pct_fmt    = workbook.add_format({"border": 1, "num_format": "0.0", "valign": "vcenter"})
        ok_fmt     = workbook.add_format({"border": 1, "bg_color": "#C6EFCE", "font_color": "#276221",
                                          "num_format": "0.0", "valign": "vcenter"})
        warn_fmt   = workbook.add_format({"border": 1, "bg_color": "#FFEB9C", "font_color": "#9C5700",
                                          "num_format": "0.0", "valign": "vcenter"})
        bad_fmt    = workbook.add_format({"border": 1, "bg_color": "#FFC7CE", "font_color": "#9C0006",
                                          "num_format": "0.0", "valign": "vcenter"})
        # Azul para upload/download
        upload_fmt = workbook.add_format({"border": 1, "bg_color": "#DDEEFF", "font_color": "#1A5276",
                                          "num_format": "0.0", "valign": "vcenter"})
        meta_fmt   = workbook.add_format({"bold": True, "bg_color": "#D6E4F7", "font_color": "#1F4E78",
                                          "border": 1})
        val_fmt    = workbook.add_format({"border": 1})

        # ── Aba 1: Dados ─────────────────────────────────────────────────
        ws = workbook.add_worksheet("Ociosidade")
        ws.freeze_panes(2, 0)
        ws.autofilter(1, 0, 1, len(columns) - 1)
        ws.set_row(0, 22)
        ws.set_row(1, 30)

        ws.merge_range(0, 0, 0, len(columns) - 1,
                       f"Relatório de Ociosidade  |  {report_name}  |  Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                       title_fmt)

        col_widths = {"Data": 14, "DiaSemana": 16, "AnoMes": 12, "Usuario": 28,
                      "Dias_Uteis": 12, "Ocioso_Horas": 15, "Perc_Ativo": 13,
                      "Perc_Ocioso": 15, "Interpretacao": 22, "Maior_Ocioso_Horas": 20}
        col_labels = {"Data": "Data", "DiaSemana": "Dia da Semana", "AnoMes": "Mês",
                      "Usuario": "Usuário", "Dias_Uteis": "Dias Úteis",
                      "Ocioso_Horas": "Ocioso (h)", "Perc_Ativo": "% Ativo",
                      "Perc_Ocioso": "% Ocioso",
                      "Interpretacao": "Interpretação",
                      "Maior_Ocioso_Horas": "Maior Tempo Ocioso (h)"}

        for ci, col_name in enumerate(columns):
            ws.set_column(ci, ci, col_widths.get(col_name, 14))
            ws.write(1, ci, col_labels.get(col_name, col_name), hdr_fmt)

        perc_ocioso_idx  = columns.index("Perc_Ocioso")
        ocioso_h_idx     = columns.index("Ocioso_Horas")
        perc_ativo_idx   = columns.index("Perc_Ativo")
        interp_idx       = columns.index("Interpretacao")
        maior_oc_idx     = columns.index("Maior_Ocioso_Horas")

        # Formatos para a coluna Interpretação
        ok_interp     = workbook.add_format({"border": 1, "bg_color": "#C6EFCE",
                                             "font_color": "#276221", "valign": "vcenter"})
        warn_interp   = workbook.add_format({"border": 1, "bg_color": "#FFEB9C",
                                             "font_color": "#9C5700", "valign": "vcenter"})
        bad_interp    = workbook.add_format({"border": 1, "bg_color": "#FFC7CE",
                                             "font_color": "#9C0006", "valign": "vcenter"})
        upload_interp = workbook.add_format({"border": 1, "bg_color": "#DDEEFF",
                                             "font_color": "#1A5276", "valign": "vcenter",
                                             "bold": True})

        all_rows = [self._oci_merge_row(r, columns, gap_dict, upload_set, is_diario) for r in main_rows]
        for ri, row in enumerate(all_rows, start=2):
            perc_oc   = float(row[perc_ocioso_idx]) if row[perc_ocioso_idx] is not None else 0.0
            interp_v  = str(row[interp_idx]) if row[interp_idx] is not None else ""
            is_upload = "upload" in interp_v.lower() or "download" in interp_v.lower()

            for ci, value in enumerate(row):
                v = value if value is not None else ""
                if ci == perc_ocioso_idx:
                    if is_upload:
                        fmt = upload_fmt
                    elif perc_oc < 25:
                        fmt = ok_fmt
                    elif perc_oc < 60:
                        fmt = warn_fmt
                    else:
                        fmt = bad_fmt
                    ws.write(ri, ci, v, fmt)
                elif ci == interp_idx:
                    if is_upload:
                        fmt = upload_interp
                    elif perc_oc < 25:
                        fmt = ok_interp
                    elif perc_oc < 60:
                        fmt = warn_interp
                    else:
                        fmt = bad_interp
                    ws.write(ri, ci, v, fmt)
                elif ci in (ocioso_h_idx, maior_oc_idx):
                    ws.write(ri, ci, v, num2_fmt)
                elif ci == perc_ativo_idx:
                    ws.write(ri, ci, v, pct_fmt)
                else:
                    ws.write(ri, ci, v, normal_fmt)
            self.progress_value.set(40 + (ri / max(total, 1)) * 30)

        # ── Aba 2: Ranking + Gráfico ──────────────────────────────────────
        ranking_rows = con.execute(q_ranking).fetchall()

        wg = workbook.add_worksheet("Ranking_Gráfico")
        wg.set_column(0, 0, 32)
        wg.set_column(1, 1, 18)
        wg.set_row(0, 22)

        wg.merge_range(0, 0, 0, 1, "Ranking de Ociosidade por Funcionário (Total de Horas)", title_fmt)
        wg.write(1, 0, "Usuário", hdr_fmt)
        wg.write(1, 1, "Total Ocioso (h)", hdr_fmt)

        rank_hdr_fmt  = workbook.add_format({"bold": True, "bg_color": "#E8EEFC", "border": 1,
                                              "num_format": "0.00"})
        for ri, (usuario, horas) in enumerate(ranking_rows, start=2):
            wg.write(ri, 0, usuario or "", normal_fmt)
            wg.write(ri, 1, float(horas) if horas is not None else 0.0, num2_fmt)

        n = len(ranking_rows)

        # Gráfico de barras horizontais
        chart = workbook.add_chart({"type": "bar"})
        chart.set_title({"name": "Funcionários com maior tempo de ociosidade"})
        chart.set_x_axis({"name": "Horas ociosas"})
        chart.set_y_axis({"name": "Usuário"})
        chart.set_legend({"none": True})
        chart.set_size({"width": 720, "height": max(300, n * 28 + 80)})
        chart.set_style(10)

        if n > 0:
            chart.add_series({
                "name":       "Ocioso (h)",
                "categories": ["Ranking_Gráfico", 2, 0, 1 + n, 0],
                "values":     ["Ranking_Gráfico", 2, 1, 1 + n, 1],
                "fill":       {"color": "#C0392B"},
                "data_labels": {"value": True, "num_format": "0.00"},
            })

        wg.insert_chart(2 + n + 2, 0, chart)

        # ── Aba 3: Informações ────────────────────────────────────────────
        resumo = workbook.add_worksheet("Informações")
        resumo.set_column(0, 0, 34)
        resumo.set_column(1, 1, 45)
        resumo.write(0, 0, "Ociosidade – Macro Lock Analytics", title_fmt)
        infos = [
            ("Tipo",                    report_name),
            ("Gerado em",               datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
            ("Janelas comerciais",      "08:00–12:00 e 13:00–17:00 (Seg–Sex)"),
            ("Esperado por dia útil",   "8 horas"),
            ("Total de linhas",         total),
            ("Filtro usuário",          self.report_user.get().strip() or "Todos"),
            ("Data inicial",            self.report_date_start.get().strip() or "Toda a base"),
            ("Data final",              self.report_date_end.get().strip() or "Toda a base"),
        ]
        for i, (k, v) in enumerate(infos, start=2):
            resumo.write(i, 0, k, meta_fmt)
            resumo.write(i, 1, v, val_fmt)

        r0 = len(infos) + 4
        resumo.write(r0, 0, "Legenda % Ocioso:", title_fmt)
        resumo.write(r0+1, 0, "< 25%           →  verde  (baixa ociosidade)", ok_fmt)
        resumo.write(r0+2, 0, "25–59%          →  amarelo (atenção)", warn_fmt)
        resumo.write(r0+3, 0, "≥ 60%           →  vermelho (alta ociosidade)", bad_fmt)
        upload_leg = workbook.add_format({"border": 1, "bg_color": "#DDEEFF",
                                          "font_color": "#1A5276", "bold": True})
        resumo.write(r0+4, 0, "≥ 25% + upload  →  azul (aguardando upload/download de arquivo)", upload_leg)

        self.write_log(f"Ociosidade: {total} linhas  |  Ranking: {n} usuários  |  Arquivo: {out_path.name}")
        workbook.close()

    def _export_query_to_xlsx(self, con, query: str, params: list, columns: list, out_path: Path, report_name: str, total_rows: int, color_weekend: bool = False, add_dashboard: bool = False):
        try:
            import xlsxwriter
        except Exception:
            raise RuntimeError("Instale o xlsxwriter:\npython -m pip install xlsxwriter")
        if out_path.exists():
            out_path.unlink()
        workbook = xlsxwriter.Workbook(str(out_path), {"constant_memory": True})
        title_fmt = workbook.add_format({"bold": True, "font_size": 15, "font_color": "#1F4E78"})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1})
        normal_fmt = workbook.add_format({"border": 1})
        weekend_fmt = workbook.add_format({"border": 1, "bg_color": "#FFC7CE", "font_color": "#9C0006"})

        resumo = workbook.add_worksheet("Resumo")
        resumo.set_column(0, 0, 34)
        resumo.set_column(1, 1, 45)
        resumo.write(0, 0, "Relatório Macro Lock", title_fmt)
        info = [
            ("Tipo", report_name),
            ("Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
            ("Total de registros", total_rows),
            ("Horário inicial", self.night_start.get()),
            ("Horário final", self.night_end.get()),
            ("Filtro usuário", self.report_user.get().strip() or "Todos"),
            ("Filtro equipamento", self.report_equipment.get().strip() or "Todos"),
            ("Data inicial", self.report_date_start.get().strip() or "Toda a base"),
            ("Data final", self.report_date_end.get().strip() or "Toda a base"),
        ]
        for idx, (k, v) in enumerate(info, start=2):
            resumo.write(idx, 0, k, header_fmt)
            resumo.write(idx, 1, v, normal_fmt)

        if total_rows == 0:
            ws = workbook.add_worksheet("Relatorio")
            for col, name in enumerate(columns):
                ws.write(0, col, name, header_fmt)
            ws.write(1, 0, "Sem registros para este filtro.")
            workbook.close()
            return

        parts = (total_rows + EXCEL_DATA_ROWS - 1) // EXCEL_DATA_ROWS
        for part in range(parts):
            offset = part * EXCEL_DATA_ROWS
            suffix = f"_{part + 1}" if parts > 1 else ""
            sheet_name = f"Relatorio{suffix}"[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, 0, len(columns) - 1)
            ws.set_column(0, len(columns) - 1, 20)
            for col, name in enumerate(columns):
                ws.write(0, col, name, header_fmt)
            cur = con.execute(f"{query} LIMIT {EXCEL_DATA_ROWS} OFFSET {offset}", params)
            row_idx = 1
            while True:
                rows = cur.fetchmany(20000)
                if not rows:
                    break
                dia_idx = columns.index("DiaSemana") if "DiaSemana" in columns else -1
                for row in rows:
                    row_fmt = normal_fmt
                    if color_weekend and dia_idx >= 0 and str(row[dia_idx]).strip().lower() in ("sábado", "sabado", "domingo"):
                        row_fmt = weekend_fmt
                    for col_idx, value in enumerate(row):
                        ws.write(row_idx, col_idx, value if value is not None else "", row_fmt)
                    row_idx += 1
            self.write_log(f"Aba {sheet_name} gerada com {row_idx - 1:,} linhas".replace(",", "."))
            pct = 55 + ((part + 1) / parts) * 40
            self.progress_value.set(pct)
        workbook.close()
        if add_dashboard:
            try:
                self.write_log("📊 Gerando Dashboard...")
                self._add_dashboard_to_xlsx(out_path)
                self.write_log("📊 Dashboard adicionado com sucesso.")
            except Exception as _de:
                self.write_log(f"⚠ Dashboard não gerado: {_de}", "warn")


    def _add_dashboard_to_xlsx(self, xlsx_path):
        """Adiciona abas Dados_Dash + Dashboard ao arquivo Excel gerado."""
        import pandas as pd
        import openpyxl
        from openpyxl.chart import BarChart, Reference
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.datavalidation import DataValidation

        def _fill(h): return PatternFill("solid", fgColor=h)
        def _font(h, sz=10, bold=False, italic=False):
            return Font(name="Arial", color=h, size=sz, bold=bold, italic=italic)
        def _side(): return Side(style="thin", color="CBD5E0")
        def _ball(): return Border(left=_side(), right=_side(), top=_side(), bottom=_side())
        def _ctr():  return Alignment(horizontal="center", vertical="center", wrap_text=True)
        def _lft():  return Alignment(horizontal="left", vertical="center")

        NAVY="1E3A5F"; TEAL="0D9488"; AMBER="D97706"
        LTBG="F0F4F8"; WHITE="FFFFFF"; TEXT="1A202C"; MUTED="718096"
        TOP_N=10
        DAY_ORDER=["Segunda-feira","Terça-feira","Quarta-feira",
                   "Quinta-feira","Sexta-feira","Sábado","Domingo"]
        BLOCK=max(TOP_N, len(DAY_ORDER))

        # Lê dados
        df = pd.read_excel(str(xlsx_path), sheet_name=None)
        # Pega a primeira aba que começa com "Relatorio"
        rel_key = next((k for k in df if k.startswith("Relatorio")), None)
        if not rel_key:
            raise ValueError("Aba Relatorio não encontrada.")
        rdf = df[rel_key]
        rdf["DuracaoSegundos"] = pd.to_numeric(rdf.get("DuracaoSegundos", pd.Series(dtype=float)), errors="coerce").fillna(0)

        machines = sorted(rdf["Equipamento"].dropna().unique().tolist())
        if not machines:
            raise ValueError("Nenhuma máquina encontrada no relatório.")

        # Mapeia máquina → label "usuario (MAQUINA)"
        def _machine_label(m):
            u = rdf[rdf["Equipamento"] == m]["Usuario"].dropna()
            user = u.mode()[0] if not u.empty else ""
            return f"{user} ({m})" if user else m
        machine_labels = {m: _machine_label(m) for m in machines}
        machines_display = [machine_labels[m] for m in machines]

        def _top_apps(mdf):
            t = (mdf[mdf["Aplicativo"].notna() & (mdf["Aplicativo"] != "")]
                 .groupby("Aplicativo")
                 .agg(Contagem=("Aplicativo","count"), Horas=("DuracaoSegundos","sum"))
                 .reset_index().sort_values("Contagem", ascending=False).head(TOP_N))
            t["Horas"] = (t["Horas"] / 3600).round(1)
            return t.reset_index(drop=True)

        def _top_sites(mdf):
            browsers = {"chrome.exe","msedge.exe","firefox.exe","iexplore.exe","opera.exe"}
            bdf = mdf[mdf["Aplicativo"].str.lower().isin(browsers) & mdf["JanelaEvento"].notna()] \
                if "JanelaEvento" in mdf.columns else pd.DataFrame()
            if bdf.empty and "CampoExtra" in mdf.columns:
                bdf = mdf[mdf["CampoExtra"].str.contains(r"\.", na=False)]
            col = "JanelaEvento" if ("JanelaEvento" in bdf.columns and not bdf.empty) else "CampoExtra"
            if bdf.empty or col not in bdf.columns:
                return pd.DataFrame(columns=["Janela","Contagem"])
            t = (bdf[bdf[col].notna()].groupby(col).size()
                 .reset_index(name="Contagem").sort_values("Contagem",ascending=False).head(TOP_N))
            t.columns = ["Janela","Contagem"]
            return t.reset_index(drop=True)

        def _by_day(mdf):
            t = (mdf.groupby("DiaSemana").size().reindex(DAY_ORDER)
                 .fillna(0).astype(int).reset_index())
            t.columns = ["Dia","Contagem"]
            return t

        mdata = {}
        for m in machines:
            mdf = rdf[rdf["Equipamento"] == m]
            mdata[m] = {"apps": _top_apps(mdf), "sites": _top_sites(mdf),
                        "days": _by_day(mdf)}

        wb = openpyxl.load_workbook(str(xlsx_path))
        for sn in ["Dashboard","Dados_Dash"]:
            if sn in wb.sheetnames: del wb[sn]

        # ── Dados_Dash (oculta) ──────────────────────────────────────────
        wsd = wb.create_sheet("Dados_Dash")
        wsd.sheet_state = "hidden"
        wsd.append(["Maquina","Rank","Aplicativo","Contagem_App","Horas_App",
                    "Site_Janela","Contagem_Site","Dia","Contagem_Dia"])
        for m in machines:
            da = mdata[m]; lbl = machine_labels[m]
            apps=da["apps"]; sites=da["sites"]; days=da["days"]
            for i in range(BLOCK):
                an = apps.iloc[i]["Aplicativo"] if i<len(apps) else ""
                ac = int(apps.iloc[i]["Contagem"]) if i<len(apps) else ""
                ah = apps.iloc[i]["Horas"]     if i<len(apps) else ""
                sn_ = sites.iloc[i]["Janela"]  if i<len(sites) else ""
                sc = int(sites.iloc[i]["Contagem"]) if i<len(sites) else ""
                dn = days.iloc[i]["Dia"]       if i<len(days) else ""
                dc = int(days.iloc[i]["Contagem"]) if i<len(days) else ""
                wsd.append([lbl, i+1, an, ac, ah, sn_, sc, dn, dc])

        # ── Dashboard ────────────────────────────────────────────────────
        ws = wb.create_sheet("Dashboard", 0)
        ws.sheet_view.showGridLines = False
        for col,w in {"A":2,"B":4,"C":32,"D":14,"E":12,"F":2,"G":32,"H":14,"I":2,"J":28,"K":16,"L":2}.items():
            ws.column_dimensions[col].width = w

        # Título
        ws.row_dimensions[1].height = 36
        ws.merge_cells("A1:L1")
        c=ws["A1"]; c.value="⏰  DASHBOARD — CONTROLE FORA DO HORÁRIO"
        c.fill=_fill(NAVY); c.font=Font(name="Arial",color=WHITE,size=16,bold=True); c.alignment=_ctr()
        ws.row_dimensions[2].height=18
        ws.merge_cells("A2:L2")
        c=ws["A2"]; c.value="Selecione a máquina no campo abaixo — os gráficos atualizam automaticamente"
        c.fill=_fill(NAVY); c.font=Font(name="Arial",color="A0AEC0",size=9,italic=True); c.alignment=_ctr()

        # Seletor
        ws.row_dimensions[3].height=8
        ws.row_dimensions[4].height=28
        ws.merge_cells("A4:B4"); ws["A4"].value="🖥  MÁQUINA:"
        ws["A4"].fill=_fill(LTBG); ws["A4"].font=Font(name="Arial",color=NAVY,size=10,bold=True)
        ws["A4"].alignment=Alignment(horizontal="right",vertical="center")
        ws.merge_cells("C4:E4")
        sel=ws["C4"]; sel.value=machines_display[0]
        sel.fill=_fill(WHITE); sel.font=Font(name="Arial",color=NAVY,size=11,bold=True); sel.alignment=_lft()
        sel.border=Border(left=Side(style="medium",color=TEAL), right=_side(), top=_side(), bottom=_side())
        dv=DataValidation(type="list",formula1='"'+",".join(machines_display)+'"',allow_blank=False,showDropDown=False)
        dv.sqref="C4"; ws.add_data_validation(dv)

        # M4 (oculto) extrai o nome real da máquina de dentro dos parênteses
        ws["M4"].value='=IFERROR(MID(C4,FIND("(",C4)+1,FIND(")",C4)-FIND("(",C4)-1),C4)'
        ws.column_dimensions["M"].hidden=True

        SEL="Dashboard!$C$4"
        MKEY="Dashboard!$M$4"   # nome real da máquina para COUNTIF/SUMIF
        MF=f"MATCH({SEL},Dados_Dash!$A:$A,0)"

        # Fórmulas helper (colunas J/K) — apps 12..21, sites 23..32, days 34..40
        APPS_S=12; SITES_S=23; DAYS_S=34
        for i in range(TOP_N):
            r=APPS_S+i
            ws.cell(r,10).value=f"=IFERROR(INDEX(Dados_Dash!$C:$C,{MF}+{i}),\"\")"
            ws.cell(r,11).value=f"=IFERROR(INDEX(Dados_Dash!$D:$D,{MF}+{i}),\"\")"
        for i in range(TOP_N):
            r=SITES_S+i
            ws.cell(r,10).value=f"=IFERROR(INDEX(Dados_Dash!$F:$F,{MF}+{i}),\"\")"
            ws.cell(r,11).value=f"=IFERROR(INDEX(Dados_Dash!$G:$G,{MF}+{i}),\"\")"
        for i in range(len(DAY_ORDER)):
            r=DAYS_S+i
            ws.cell(r,10).value=f"=IFERROR(INDEX(Dados_Dash!$H:$H,{MF}+{i}),\"\")"
            ws.cell(r,11).value=f"=IFERROR(INDEX(Dados_Dash!$I:$I,{MF}+{i}),\"\")"
        ws.column_dimensions["J"].hidden=True
        ws.column_dimensions["K"].hidden=True

        # Cards de resumo (linha 6-8)
        ws.row_dimensions[5].height=10
        ws.row_dimensions[6].height=20; ws.row_dimensions[7].height=36; ws.row_dimensions[8].height=4
        cards=[
            (3, "TOTAL DE REGISTROS",        f"=COUNTIF(Relatorio!$E:$E,{MKEY})", "#,##0", NAVY),
            (5, "HORAS FORA DO HORÁRIO",      f"=ROUND(SUMIF(Relatorio!$E:$E,{MKEY},Relatorio!$K:$K)/3600,1)", '#,##0.0 "h"', TEAL),
            (7, "APLICATIVOS DISTINTOS",
             f'=IFERROR(SUMPRODUCT(({MKEY}=Relatorio!$E$2:$E$20000)*(Relatorio!$H$2:$H$20000<>"")*IFERROR(1/COUNTIFS(Relatorio!$E$2:$E$20000,{MKEY},Relatorio!$H$2:$H$20000,Relatorio!$H$2:$H$20000),0)),0)',
             "#,##0", AMBER),
            (9, "MÁQUINAS NO RELATÓRIO",      f"={len(machines)}", "#,##0", "5B21B6"),
        ]
        for cs,lbl,formula,fmt,acc in cards:
            ws.merge_cells(f"{get_column_letter(cs)}6:{get_column_letter(cs+1)}6")
            ws.merge_cells(f"{get_column_letter(cs)}7:{get_column_letter(cs+1)}7")
            ws.merge_cells(f"{get_column_letter(cs)}8:{get_column_letter(cs+1)}8")
            tc=ws[f"{get_column_letter(cs)}6"]; tc.value=lbl
            tc.fill=_fill(acc); tc.font=Font(name="Arial",color=WHITE,size=8,bold=True); tc.alignment=_ctr()
            vc=ws[f"{get_column_letter(cs)}7"]; vc.value=formula; vc.number_format=fmt
            vc.fill=_fill(WHITE); vc.font=Font(name="Arial",color=NAVY,size=16,bold=True); vc.alignment=_ctr()
            for r in [6,8]:
                for col in range(cs,cs+2):
                    try: ws.cell(row=r,column=col).fill=_fill(acc)
                    except AttributeError: pass

        # ── Tabela Apps ──────────────────────────────────────────────────
        ws.row_dimensions[9].height=10; ws.row_dimensions[10].height=24; ws.row_dimensions[11].height=18
        ws.merge_cells("C10:E10"); t=ws["C10"]; t.value="📱  TOP 10 APLICATIVOS"
        t.fill=_fill(TEAL); t.font=Font(name="Arial",color=WHITE,size=11,bold=True); t.alignment=_lft()
        t.border=Border(left=Side(style="medium",color=TEAL))
        for col,lbl in zip(["C","D","E"],["Aplicativo","Registros","Horas"]):
            c=ws[f"{col}11"]; c.value=lbl; c.fill=_fill("E8F4F8")
            c.font=Font(name="Arial",color=NAVY,size=9,bold=True); c.alignment=_ctr(); c.border=_ball()
        for i in range(TOP_N):
            row=APPS_S+i; bg=WHITE if i%2==0 else LTBG
            ws.row_dimensions[row].height=16
            ca=ws[f"C{row}"]; ca.value=f"=J{row}"; ca.fill=_fill(bg)
            ca.font=Font(name="Arial",color=TEXT,size=9); ca.alignment=_lft(); ca.border=_ball()
            cb=ws[f"D{row}"]; cb.value=f"=IF(K{row}=\"\",\"\",K{row})"; cb.number_format="#,##0"
            cb.fill=_fill(bg); cb.font=Font(name="Arial",color=NAVY,size=9,bold=True); cb.alignment=_ctr(); cb.border=_ball()
            cc=ws[f"E{row}"]; cc.value=f"=IFERROR(INDEX(Dados_Dash!$E:$E,{MF}+{i}),\"\")"
            cc.number_format='#,##0.0 "h"'; cc.fill=_fill(bg)
            cc.font=Font(name="Arial",color=MUTED,size=9); cc.alignment=_ctr(); cc.border=_ball()

        # ── Tabela Sites ─────────────────────────────────────────────────
        ws.row_dimensions[22].height=10; ws.row_dimensions[22].height=24; ws.row_dimensions[22].height=10
        S2T=22
        ws.row_dimensions[S2T].height=24; ws.row_dimensions[S2T+1].height=18
        ws.merge_cells(f"C{S2T}:E{S2T}"); t=ws[f"C{S2T}"]; t.value="🌐  TOP 10 SITES / JANELAS"
        t.fill=_fill(AMBER); t.font=Font(name="Arial",color=WHITE,size=11,bold=True); t.alignment=_lft()
        t.border=Border(left=Side(style="medium",color=AMBER))
        for col,lbl in zip(["C","D"],["Site / Janela","Registros"]):
            c=ws[f"{col}{S2T+1}"]; c.value=lbl; c.fill=_fill("FEF9EC")
            c.font=Font(name="Arial",color=NAVY,size=9,bold=True); c.alignment=_ctr(); c.border=_ball()
        for i in range(TOP_N):
            row=SITES_S+i; jrow=SITES_S+i; bg=WHITE if i%2==0 else "FFFBEB"
            ws.row_dimensions[row].height=16
            ca=ws[f"C{row}"]; ca.value=f"=J{jrow}"; ca.fill=_fill(bg)
            ca.font=Font(name="Arial",color=TEXT,size=9); ca.alignment=_lft(); ca.border=_ball()
            cb=ws[f"D{row}"]; cb.value=f"=IF(K{jrow}=\"\",\"\",K{jrow})"; cb.number_format="#,##0"
            cb.fill=_fill(bg); cb.font=Font(name="Arial",color=NAVY,size=9,bold=True); cb.alignment=_ctr(); cb.border=_ball()

        # ── Tabela Dias ───────────────────────────────────────────────────
        S3T=DAYS_S-1
        ws.row_dimensions[S3T].height=24; ws.row_dimensions[S3T+1].height=18
        ws.merge_cells(f"C{S3T}:E{S3T}"); t=ws[f"C{S3T}"]; t.value="📅  REGISTROS POR DIA DA SEMANA"
        t.fill=_fill(NAVY); t.font=Font(name="Arial",color=WHITE,size=11,bold=True); t.alignment=_lft()
        for col,lbl in zip(["C","D"],["Dia","Registros"]):
            c=ws[f"{col}{S3T+1}"]; c.value=lbl; c.fill=_fill("E8EAF6")
            c.font=Font(name="Arial",color=NAVY,size=9,bold=True); c.alignment=_ctr(); c.border=_ball()
        for i in range(len(DAY_ORDER)):
            row=DAYS_S+i; bg=WHITE if i%2==0 else LTBG
            ws.row_dimensions[row].height=16
            ca=ws[f"C{row}"]; ca.value=f"=J{row}"; ca.fill=_fill(bg)
            ca.font=Font(name="Arial",color=TEXT,size=9); ca.alignment=_lft(); ca.border=_ball()
            cb=ws[f"D{row}"]; cb.value=f"=IF(K{row}=\"\",\"\",K{row})"; cb.number_format="#,##0"
            cb.fill=_fill(bg); cb.font=Font(name="Arial",color=NAVY,size=9,bold=True); cb.alignment=_ctr(); cb.border=_ball()

        # ── Gráficos ──────────────────────────────────────────────────────
        def _make_chart(chart_type, bar_dir, title, col_color, cats_rows, vals_rows, anchor):
            ch = BarChart(); ch.type=bar_dir; ch.grouping="clustered"
            ch.title=title; ch.style=10; ch.legend=None
            ch.width=14; ch.height=10; ch.roundedCorners=False
            ch.x_axis.delete=False; ch.y_axis.delete=False
            cats=Reference(ws, min_col=10, min_row=cats_rows[0], max_row=cats_rows[1])
            vals=Reference(ws, min_col=11, min_row=vals_rows[0], max_row=vals_rows[1])
            ch.add_data(vals); ch.set_categories(cats)
            ch.series[0].graphicalProperties.solidFill=col_color
            ch.series[0].graphicalProperties.line.solidFill=col_color
            ws.add_chart(ch, anchor)

        _make_chart("bar","bar","Top 10 Aplicativos",TEAL,(APPS_S,APPS_S+TOP_N-1),(APPS_S,APPS_S+TOP_N-1),"G10")
        _make_chart("bar","bar","Top 10 Sites / Janelas",AMBER,(SITES_S,SITES_S+TOP_N-1),(SITES_S,SITES_S+TOP_N-1),"G22")
        _make_chart("col","col","Registros por Dia da Semana",NAVY,(DAYS_S,DAYS_S+6),(DAYS_S,DAYS_S+6),"G34")

        ws.freeze_panes="A9"

        # Reordenar: Dashboard primeiro
        order=["Dashboard"]+[s for s in wb.sheetnames if s not in ("Dashboard","Dados_Dash")]+["Dados_Dash"]
        wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)

        wb.save(str(xlsx_path))


def main():
    root = tk.Tk()
    MacroLockAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
