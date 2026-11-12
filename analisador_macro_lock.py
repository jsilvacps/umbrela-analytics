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
APP_VERSION = "7.0"
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
    "Status", "Aplicativo", "CampoExtra", "JanelaEvento", "DuracaoSegundos", "TempoFormatado",
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
        saved_email = self._mk_load_email()
        self._mk_email   = tk.StringVar(value=saved_email)
        self._mk_pass    = tk.StringVar()
        self._mk_status  = tk.StringVar(value="Não conectado")
        self._mk_logged  = False
        self._mk_pass_md5: str = ""   # hash em memória para re-login automático
        self._mk_base_url = "https://dmsempresarial.makrolock.com.br"
        self._mk_group_computers: dict[str, list[str]] = {}  # group_id → [computer_ids]
        self._mk_date_start = tk.StringVar(value="")
        self._mk_date_end   = tk.StringVar(value="")
        self._mk_report_type = tk.StringVar(value="sessions")
        self._mk_busy    = False
        self.output_path = tk.StringVar(value=r"C:\Users\TI\Desktop\indicadores")
        self.exclude_terms = tk.StringVar(value=DEFAULT_EXCLUDE)
        self.status_text = tk.StringVar(value="Selecione o CSV exportado do Macro Lock.")
        self.progress_value = tk.DoubleVar(value=0)

        self.night_start = tk.StringVar(value="18:00:00")
        self.night_end = tk.StringVar(value="07:00:00")
        self.report_type = tk.StringVar(value="fora_horario")
        self.report_format = tk.StringVar(value="xlsx")
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

    def _load_logo(self):
        """Carrega a logo PNG quando ela estiver na mesma pasta do .py/.exe."""
        try:
            logo_path = app_dir() / LOGO_FILE
            if logo_path.exists():
                img = tk.PhotoImage(file=str(logo_path))
                # Reduz automaticamente se a imagem estiver muito grande.
                max_w = 330
                if img.width() > max_w:
                    factor = max(1, int(img.width() / max_w))
                    img = img.subsample(factor, factor)
                self._logo_img = img
                return img
        except Exception:
            pass
        self._logo_img = None
        return None

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

        # Linha futurista topo: gradiente teal→accent
        top_bar = tk.Frame(header, height=3)
        top_bar.pack(fill="x", side="top")
        # Simulação de gradiente: 3 segmentos de cor
        tk.Frame(top_bar, bg=c["teal"],   height=3).pack(side="left", fill="y", expand=True)
        tk.Frame(top_bar, bg=c["accent"], height=3).pack(side="left", fill="y", expand=True)
        tk.Frame(top_bar, bg=c["purple"], height=3).pack(side="left", fill="y", expand=True)

        # Logo
        logo = self._load_logo()
        if logo is not None:
            tk.Label(header, image=logo, bg=c["sidebar"], borderwidth=0).pack(
                side="left", padx=(18, 8), pady=6)

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
        tk.Label(sub_frame, text="Painel de Análise", bg=c["sidebar"],
                 fg="#B8D8F8", font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(sub_frame, text="Makrolock · Analytics · RH", bg=c["sidebar"],
                 fg="#6B8EB0", font=("Segoe UI", 8)).pack(anchor="w")

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
        tk.Label(u_info, text="TI · DMS Empresarial",
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

        # Log area
        log_wrap = tk.Frame(bottom, bg=c["sidebar"], pady=0)
        log_wrap.pack(fill="x")
        log_hdr_row = tk.Frame(log_wrap, bg=c["sidebar"])
        log_hdr_row.pack(fill="x", padx=14, pady=(5, 2))
        tk.Label(log_hdr_row, text="LOG", bg=c["sidebar"],
                 fg=c["muted"], font=("Segoe UI", 7, "bold")).pack(side="left")
        self.log = tk.Text(
            log_wrap, height=3,
            bg="#0A1020", fg="#94A3B8",
            insertbackground=c["accent"],
            selectbackground="#1E3A6E",
            selectforeground="#FFFFFF",
            relief="flat", borderwidth=0,
            font=("Consolas", 9),
            wrap="word", padx=14, pady=6,
        )
        self.log.pack(fill="x")
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
            ("dashboard",  "▣", "Dashboard"),
            ("makrolock",  "⬡", "Makrolock"),
            ("processar",  "⚙", "Processar"),
            ("relatorios", "≡", "Relatórios"),
        ]
        for key, icon, label in nav_list:
            page = tk.Frame(content_area, bg=c["bg"], padx=24, pady=20)
            self._pages[key] = page

        def _show(key):
            if self._cur_page:
                self._pages[self._cur_page].pack_forget()
            self._pages[key].pack(fill="both", expand=True)
            self._cur_page = key
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
        tk.Label(sidebar, text=f"v{APP_VERSION} · © 2025 DMS",
                 bg=c["sidebar"], fg="#2D5A8A",
                 font=("Segoe UI", 7)).pack(side="bottom", pady=(4, 8))

        # Efeito futurista: scan lines no rodapé da sidebar
        def _draw_scanlines(cv, w, h):
            cv.delete("all")
            # Fundo gradiente simulado — faixas de azul decrescente de baixo
            steps = 18
            for i in range(steps):
                y0 = h - (i + 1) * (h // steps)
                y1 = h - i * (h // steps)
                alpha_hex = format(int(20 + i * 6), "02x")
                cv.create_rectangle(0, y0, w, y1, fill=f"#06{alpha_hex}FF", outline="")
            # Linhas horizontais luminosas (cyan/teal)
            line_colors = ["#063050", "#0A4A70", "#0E6090", "#0A4A70",
                           "#063050", "#042030", "#031020"]
            for idx, lc in enumerate(line_colors):
                y = h - 6 - idx * 8
                cv.create_line(0, y, w, y, fill=lc, width=1)
            # Linha brilhante no topo do bloco
            cv.create_line(0, 0, w, 0, fill="#1A8AAA", width=1)

        scan_cv = tk.Canvas(sidebar, bg=c["sidebar"], height=70,
                            highlightthickness=0, bd=0)
        scan_cv.pack(side="bottom", fill="x")
        scan_cv.bind("<Configure>",
                     lambda e: _draw_scanlines(scan_cv, e.width, e.height))

        # Constrói páginas
        tab_dashboard = self._pages["dashboard"]
        tab_makrolock = self._pages["makrolock"]
        tab_process   = self._pages["processar"]
        tab_reports   = self._pages["relatorios"]

        self._build_dashboard_tab(tab_dashboard)
        self._build_makrolock_tab(tab_makrolock)
        self._build_process_tab(tab_process)
        self._build_reports_tab(tab_reports)

        _show("dashboard")

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

    def _build_dashboard_tab(self, parent):
        c = self.COLORS

        # Scroll para o dashboard
        canvas = tk.Canvas(parent, bg=c["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
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

        # ── Cabeçalho: logo centralizada + botões ─────────────────────
        hdr_wrap = tk.Frame(inner, bg=c["bg"])
        hdr_wrap.pack(fill="x", pady=(0, 24))

        # Logo centralizada
        logo_dash = self._load_logo()
        if logo_dash is not None:
            lbl_logo = tk.Label(hdr_wrap, image=logo_dash, bg=c["bg"], borderwidth=0)
            lbl_logo.image = logo_dash
            lbl_logo.pack(anchor="center", pady=(12, 6))
        else:
            tk.Label(hdr_wrap, text="Umbrela Analytics", bg=c["bg"], fg=c["text"],
                     font=("Segoe UI", 20, "bold")).pack(anchor="center", pady=(12, 2))

        tk.Label(hdr_wrap, text="Visão geral da base de dados processada",
                 bg=c["bg"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="center")

        # Separador fino
        tk.Frame(hdr_wrap, bg=c["border2"], height=1).pack(fill="x", pady=(14, 0))

        # Botões centralizados abaixo
        btn_right = tk.Frame(hdr_wrap, bg=c["bg"])
        btn_right.pack(anchor="center", pady=(10, 0))
        ttk.Button(btn_right, text="↻  Atualizar", style="Ghost.TButton",
                   command=self.refresh_dashboard).pack(side="left", padx=(0, 8))
        ttk.Button(btn_right, text="▶  Processar base", style="Ghost.TButton",
                   command=lambda: self._show_page("processar")).pack(side="left", padx=(0, 8))
        ttk.Button(btn_right, text="⬆  Carregar CSV", style="Accent.TButton",
                   command=lambda: self._show_page("processar")).pack(side="left")

        # ── Fila de 6 KPI tiles ───────────────────────────────────────
        tiles_row = tk.Frame(inner, bg=c["bg"])
        tiles_row.pack(fill="x", pady=(0, 16))
        for i in range(6):
            tiles_row.columnconfigure(i, weight=1)

        tiles = [
            ("Total de Eventos", self.metric_total,   "registros na base", "▣", c["accent"],  c["accentbg"]),
            ("Fora do Horário",  self.metric_night,   "18:00 – 06:59",     "◑", c["orange"],  c["orangebg"]),
            ("Fim de Semana",    self.metric_weekend, "sábado e domingo",   "◷", c["purple"],  c["purplebg"]),
            ("Usuários Únicos",  self.metric_users,   "identificados",      "◈", c["teal"],    c["tealbg"]),
            ("Equipamentos",     self.metric_equips,  "na base",            "▦", c["ok"],      c["okbg"]),
            ("Período Coberto",  self.metric_period,  None,                 "◫", c["pink"],    c["pinkbg"]),
        ]
        for i, (title, var, sub, icon, col, bgcol) in enumerate(tiles):
            t = self._metric_card(tiles_row, title, var, sub, icon, col, bgcol)
            t.grid(row=0, column=i, sticky="nsew",
                   padx=(0 if i == 0 else 6, 0))

        # ── Seção do meio: empty state + atividade ────────────────────
        mid = tk.Frame(inner, bg=c["bg"])
        mid.pack(fill="x", pady=(0, 16))
        mid.columnconfigure(0, weight=3)
        mid.columnconfigure(1, weight=2)

        # Painel esquerdo: "Nenhuma base carregada"
        left_panel = tk.Frame(mid, bg=c["card"],
                              highlightthickness=1, highlightbackground=c["border2"])
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        # Conteúdo centralizado
        empty_inner = tk.Frame(left_panel, bg=c["card"], pady=40)
        empty_inner.pack(expand=True, fill="both")

        # Ícone de banco de dados (texto artístico)
        db_icon = tk.Label(empty_inner, text="⬡", bg=c["card"], fg=c["teal"],
                           font=("Segoe UI", 48))
        db_icon.pack()

        tk.Label(empty_inner, text="Nenhuma base carregada", bg=c["card"],
                 fg=c["text"], font=("Segoe UI", 13, "bold")).pack(pady=(12, 4))
        tk.Label(empty_inner,
                 text="Ainda não há uma base de dados disponível para análise.\nCarregue um arquivo CSV para processar e visualizar os resultados.",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 9),
                 justify="center").pack()

        tk.Frame(empty_inner, bg=c["card"], height=20).pack()
        ttk.Button(empty_inner, text="⬆  Carregar CSV", style="Accent.TButton",
                   command=lambda: self._show_page("processar")).pack()
        tk.Label(empty_inner,
                 text="Após o carregamento, clique em Processar base para gerar os indicadores.",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 8)).pack(pady=(10, 0))

        # Painel direito: visão de atividade
        right_panel = tk.Frame(mid, bg=c["card"],
                               highlightthickness=1, highlightbackground=c["border2"])
        right_panel.grid(row=0, column=1, sticky="nsew")

        rp_top = tk.Frame(right_panel, bg=c["card"], padx=20, pady=16)
        rp_top.pack(fill="x")
        tk.Label(rp_top, text="▐ Visão de Atividade", bg=c["card"], fg=c["text"],
                 font=("Segoe UI", 11, "bold")).pack(side="left")

        rp_body = tk.Frame(right_panel, bg=c["card"], pady=50)
        rp_body.pack(expand=True, fill="both")
        tk.Label(rp_body, text="▐▐▐", bg=c["card"], fg=c["border2"],
                 font=("Segoe UI", 28)).pack()
        tk.Label(rp_body, text="Sem dados para exibir", bg=c["card"],
                 fg=c["text"], font=("Segoe UI", 11, "bold")).pack(pady=(12, 4))
        tk.Label(rp_body,
                 text="Carregue e processe uma base para visualizar\na distribuição de eventos ao longo do tempo.",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 9), justify="center").pack()

        # ── Início rápido ─────────────────────────────────────────────
        qstart = tk.Frame(inner, bg=c["card"],
                          highlightthickness=1, highlightbackground=c["border2"])
        qstart.pack(fill="x")
        qs_hdr = tk.Frame(qstart, bg=c["card"], padx=20, pady=16)
        qs_hdr.pack(fill="x")
        tk.Label(qs_hdr, text="⚡ Início rápido", bg=c["card"], fg=c["text"],
                 font=("Segoe UI", 11, "bold")).pack(side="left")
        tk.Label(qs_hdr, text="Siga os passos abaixo para processar sua base e gerar relatórios.",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 9)).pack(side="left", padx=(12, 0))

        steps_row = tk.Frame(qstart, bg=c["card"])
        steps_row.pack(fill="x", padx=20, pady=(0, 20))

        step_data = [
            ("1", "▣", "Carregue o CSV",   "Selecione o arquivo da base de dados."),
            ("2", "▶", "Processe a base",  "Execute o processamento dos registros."),
            ("3", "▐", "Analise os resultados", "Acompanhe os indicadores e gere relatórios."),
        ]
        for idx, (num, icon, title, desc) in enumerate(step_data):
            if idx > 0:
                # Seta separadora
                tk.Label(steps_row, text=" → ", bg=c["card"], fg=c["muted"],
                         font=("Segoe UI", 16)).pack(side="left")

            step_frm = tk.Frame(steps_row, bg=c["card2"],
                                highlightthickness=1, highlightbackground=c["border"])
            step_frm.pack(side="left", fill="x", expand=True, padx=(0 if idx == 0 else 0, 0))
            step_inner = tk.Frame(step_frm, bg=c["card2"], padx=16, pady=14)
            step_inner.pack(fill="both")

            top_step = tk.Frame(step_inner, bg=c["card2"])
            top_step.pack(fill="x")

            # Número em círculo
            num_frm = tk.Frame(top_step, bg=c["teal"], width=28, height=28)
            num_frm.pack_propagate(False)
            num_frm.pack(side="left")
            tk.Label(num_frm, text=num, bg=c["teal"], fg="#FFFFFF",
                     font=("Segoe UI", 10, "bold")).place(relx=0.5, rely=0.5, anchor="center")

            # Ícone
            tk.Label(top_step, text=icon, bg=c["card2"], fg=c["muted2"],
                     font=("Segoe UI", 16)).pack(side="left", padx=(10, 0))

            tk.Label(step_inner, text=title, bg=c["card2"], fg=c["text"],
                     font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
            tk.Label(step_inner, text=desc, bg=c["card2"], fg=c["muted"],
                     font=("Segoe UI", 8), wraplength=160, justify="left").pack(anchor="w")

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

        # Cria os quatro frames
        self._mk_frame_login    = self._build_mk_login_frame(self._mk_container)
        self._mk_frame_choice   = self._build_mk_choice_frame(self._mk_container)
        self._mk_frame_reports  = self._build_mk_reports_frame(self._mk_container)
        self._mk_frame_classify = self._build_mk_classify_frame(self._mk_container)

        # Começa na tela de login
        self._mk_show_login_frame()

    # ── Navegação entre frames ────────────────────────────────────────────
    def _mk_hide_all(self):
        for f in (self._mk_frame_login, self._mk_frame_choice,
                  self._mk_frame_reports, self._mk_frame_classify):
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
        pass_entry.grid(row=4, column=0, sticky="ew", pady=(0, 24))
        pass_entry.bind("<Return>", lambda e: self._mk_login())

        # Botão entrar — full width
        ttk.Button(form, text="  Entrar no Makrolock  ", style="Accent.TButton",
                   command=self._mk_login).grid(row=5, column=0, sticky="ew")

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
        ttk.Entry(date_row, textvariable=self._mk_date_start, width=14).pack(side="left", padx=(4, 0))
        tk.Label(date_row, text="ATÉ", bg=c["card"], fg=c["muted"],
                 font=("Segoe UI", 8)).pack(side="left", padx=(12, 0))
        ttk.Entry(date_row, textvariable=self._mk_date_end, width=14).pack(side="left", padx=(4, 0))
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
            ("sessions",     "📋 Atividades"),
            ("applications", "💻 Aplicativos"),
            ("internet",     "🌐 Internet"),
            ("devices",      "🖨 Dispositivos"),
            ("files",        "📁 Arquivos"),
            ("productivity", "📈 Produtividade"),
            ("inventory",    "🗂 Inventário"),
            ("software",     "⚙ Software"),
        ]
        self._mk_type_vars = {}
        for key, label in _MK_REPORT_TYPES:
            var = tk.BooleanVar(value=(key == "sessions"))
            self._mk_type_vars[key] = var
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

    # ── Métodos de integração Makrolock ───────────────────────────────────
    def _mk_load_email(self) -> str:
        import json
        try:
            if self._mk_config_file.exists():
                return json.loads(self._mk_config_file.read_text(encoding="utf-8")).get("email", "")
        except Exception:
            pass
        return ""

    def _mk_save_email(self, email: str):
        import json
        try:
            self._mk_config_file.write_text(
                json.dumps({"email": email}, ensure_ascii=False), encoding="utf-8"
            )
        except Exception:
            pass

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
                self._mk_save_email(email)
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
            # Formato esperado: lista de {id, name} ou {id, alias/name}
            groups = data if isinstance(data, list) else data.get("groups", data.get("data", []))
            self._mk_log(f"Grupos recebidos: {len(groups)}  raw: {str(groups)[:300]}")

            # Atualiza UI na thread principal
            def _update():
                c = self.COLORS
                # Remove widgets antigos (exceto o label de instrução e o botão recarregar)
                for w in list(self._mk_groups_frame.winfo_children()):
                    w.destroy()
                self._mk_group_vars.clear()

                tk.Label(self._mk_groups_frame, text="GRUPOS / COMPUTADORES", bg=c["card"], fg=c["muted"],
                         font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

                if not groups:
                    tk.Label(self._mk_groups_frame, text="Nenhum grupo encontrado.",
                             bg=c["card"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w")
                else:
                    for g in groups:
                        gid  = str(g.get("id", g.get("_id", "")))
                        name = g.get("name") or g.get("alias") or g.get("label") or gid
                        var = tk.BooleanVar(value=True)
                        self._mk_group_vars[gid] = (var, name)
                        self._mk_group_computers[gid] = [str(c) for c in (g.get("computerIds") or [])]
                        ttk.Checkbutton(self._mk_groups_frame,
                                        text=name, variable=var).pack(anchor="w", pady=1)

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
        threading.Thread(target=self._mk_do_fetch, daemon=True).start()

    def _mk_do_fetch(self):
        import tempfile, traceback as _tb
        self._mk_busy = True
        self._mk_fetch_btn.configure(state="disabled")
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
            selected_types = [(k, v) for k, v in self._mk_type_vars.items() if v.get()]
            if not selected_types:
                raise ValueError("Selecione ao menos um tipo de relatório.")

            # Grupos selecionados (None = sem filtro de grupo = todos)
            selected_groups = []  # lista de (group_id, group_name)
            if self._mk_group_vars:
                for gid, (var, name) in self._mk_group_vars.items():
                    if var.get():
                        selected_groups.append((gid, name))
                if not selected_groups:
                    raise ValueError("Selecione ao menos um grupo.")
            else:
                selected_groups = [(None, "Todos")]

            base = self._mk_base_url
            from_ms = int(dt_start.timestamp() * 1000)
            to_ms   = int(dt_end.timestamp() * 1000)

            self._mk_log(f"Período: {dt_start.strftime('%d/%m/%Y')} → {dt_end.strftime('%d/%m/%Y')}")
            self._mk_log(f"Tipos: {[t for t,_ in selected_types]}  |  Grupos: {[n for _,n in selected_groups]}")

            total_combos = len(selected_types) * len(selected_groups)
            done = 0
            added_files = []

            for type_key, type_lbl in selected_types:
                for group_id, group_name in selected_groups:
                    done += 1
                    prog = int(done / total_combos * 90)
                    self._mk_prog_var.set(prog)
                    label = f"[{done}/{total_combos}] {type_lbl} — {group_name}"
                    self._mk_prog_lbl.set(label)
                    self._mk_log(f"Baixando: {label}")
                    self.root.update_idletasks()

                    params = {
                        "type":   type_key,
                        "format": "csv",
                        "from":   from_ms,
                        "to":     to_ms,
                        "lang":   "pt",
                        "theme":  "light",
                    }
                    if group_id:
                        params["group"] = group_id

                    self._mk_log(f"  GET /api/front/reports/export  params={params}")
                    resp = self._mk_session.get(
                        f"{base}/api/front/reports/export",
                        params=params, timeout=120, stream=True,
                    )
                    self._mk_log(f"  HTTP {resp.status_code}  Content-Type: {resp.headers.get('Content-Type','?')}")

                    if resp.status_code != 200:
                        self._mk_log(f"  ⚠ Pulado: {resp.text[:200]}", "warn")
                        continue

                    # Salva CSV temporário
                    safe_group = group_name.replace(" ", "_")[:20]
                    tmp = tempfile.NamedTemporaryFile(
                        delete=False, suffix=".csv",
                        prefix=f"mk_{type_key}_{safe_group}_",
                    )
                    total_bytes = 0
                    for chunk in resp.iter_content(chunk_size=1024 * 64):
                        if chunk:
                            tmp.write(chunk)
                            total_bytes += len(chunk)
                    tmp.close()

                    # Prévia das primeiras linhas
                    try:
                        with open(tmp.name, "r", encoding="utf-8", errors="replace") as _f:
                            preview = "".join(_f.readline() for _ in range(3))
                        self._mk_log(f"  Prévia: {preview.strip()[:300]}")
                    except Exception:
                        pass

                    if total_bytes < 50:
                        self._mk_log(f"  ⚠ Vazio ({total_bytes}B) — sem dados para esse período/grupo.", "warn")
                    else:
                        self._mk_log(f"  ✅ {total_bytes/1024:.1f} KB → {tmp.name}", "ok")
                        added_files.append(tmp.name)

            # Adiciona todos à fila sem processar
            for f in added_files:
                if f not in self._multi_files:
                    self._multi_files.append(f)
            try:
                self._files_refresh()
            except Exception:
                pass

            self._mk_prog_var.set(100)
            msg = f"✅  {len(added_files)} arquivo(s) adicionado(s) à fila."
            self._mk_prog_lbl.set(msg)
            self._mk_log(msg, "ok")
            self._mk_log("Vá para ⚙ Processar e clique em '▶ Processar' quando quiser.", "ok")

            # Navega para Processar mas NÃO processa
            self._show_page("processar")
            messagebox.showinfo(
                "Relatórios baixados",
                f"{len(added_files)} arquivo(s) adicionado(s) à fila.\n\n"
                "Clique em '▶ Processar' para consolidar os dados."
            )

        except Exception as e:
            self._mk_prog_lbl.set(f"❌  {e}")
            self._mk_log(f"ERRO: {e}", "error")
            self._mk_log(_tb.format_exc(), "error")
            messagebox.showerror("Erro Makrolock", str(e))
        finally:
            self._mk_busy = False
            try:
                self._mk_fetch_btn.configure(state="normal")
            except Exception:
                pass

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
        ttk.Entry(imp_row, textvariable=self._import_path).grid(row=0, column=0, sticky="ew")
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
        ttk.Entry(grid, textvariable=self.night_start, width=14).grid(row=1, column=0, sticky="w", pady=(0, 10))
        ttk.Entry(grid, textvariable=self.night_end,   width=14).grid(row=1, column=1, sticky="w", padx=(16, 0), pady=(0, 10))

        lbl(grid, "USUÁRIO (opcional)",      2, 0)
        lbl(grid, "EQUIPAMENTO (opcional)",  2, 1, (16, 0))
        ttk.Entry(grid, textvariable=self.report_user,      width=28).grid(row=3, column=0, sticky="ew", pady=(0, 10))
        ttk.Entry(grid, textvariable=self.report_equipment, width=28).grid(row=3, column=1, sticky="ew", padx=(16, 0), pady=(0, 10))

        lbl(grid, "DATA INICIAL (opcional)", 4, 0)
        lbl(grid, "DATA FINAL (opcional)",   4, 1, (16, 0))
        ttk.Entry(grid, textvariable=self.report_date_start, width=18).grid(row=5, column=0, sticky="w", pady=(0, 4))
        ttk.Entry(grid, textvariable=self.report_date_end,   width=18).grid(row=5, column=1, sticky="w", padx=(16, 0), pady=(0, 4))
        tk.Label(grid, text="DD/MM/AAAA ou AAAA-MM-DD  —  vazio = toda a base",
                 bg=c["card"], fg=c["muted"], font=("Segoe UI", 8)).grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 12))

        tk.Frame(card, bg=c["border"], height=1).pack(fill="x", pady=(0, 10))

        self._section_label(card, "FORMATO DE SAÍDA")
        fmt_row = ttk.Frame(card, style="Card.TFrame")
        fmt_row.pack(anchor="w", pady=(0, 4))
        ttk.Radiobutton(fmt_row, text="Excel .xlsx", variable=self.report_format, value="xlsx").pack(side="left")
        ttk.Radiobutton(fmt_row, text="CSV .csv",    variable=self.report_format, value="csv").pack(side="left", padx=(16, 0))

        action = ttk.Frame(parent)
        action.pack(fill="x", pady=(12, 0))
        ttk.Button(action, text="▶  Gerar Relatório", style="Accent.TButton", command=self.start_report).pack(side="left")
        ttk.Button(action, text="Abrir pasta", style="Ghost.TButton", command=self.open_reports_folder).pack(side="left", padx=(8, 0))

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
        threading.Thread(target=self.generate_report, daemon=True).start()

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
                header = false,
                delim = '{delim}',
                quote = '"',
                escape = '"',
                ignore_errors = true,
                null_padding = true
            )
        """)
        raw_cols = [row[0] for row in con.execute("DESCRIBE bruto_raw_tmp").fetchall()]
        self.write_log(f"[{idx+1}] Colunas: {raw_cols[:9]}")

        def col(i):
            if i < len(raw_cols):
                return raw_cols[i]
            raise RuntimeError(
                f"Arquivo '{csv_path.name}' tem apenas {len(raw_cols)} colunas — esperado ao menos {i+1}."
            )

        select_sql = f"""
            SELECT
                {col(0)}::VARCHAR AS DataHora,
                {col(1)}::VARCHAR AS Equipamento,
                {col(2)}::VARCHAR AS Usuario,
                {col(3)}::VARCHAR AS Status,
                {col(4)}::VARCHAR AS Aplicativo,
                {col(5) if len(raw_cols) > 5 else "''"}::VARCHAR AS CampoExtra,
                {col(6) if len(raw_cols) > 6 else "''"}::VARCHAR AS JanelaEvento,
                TRY_CAST({col(7) if len(raw_cols) > 7 else "NULL"} AS BIGINT) AS DuracaoSegundos,
                {col(8) if len(raw_cols) > 8 else "''"}::VARCHAR AS TempoFormatado
            FROM bruto_raw_tmp
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
                    CAST(DataHoraTs AS DATE)::VARCHAR AS Data,
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
                    TempoFormatado
                FROM (
                    SELECT
                        TRY_CAST(DataHora AS TIMESTAMP) AS DataHoraTs,
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
            where = f"(CAST(DataHoraTs AS TIME) >= TIME '{start}' OR CAST(DataHoraTs AS TIME) < TIME '{end}')"
        elif report_type == "fim_semana":
            where = "CAST(strftime(DataHoraTs, '%w') AS INTEGER) IN (0, 6)"
        elif report_type == "fora_horario_fim_semana":
            where = f"((CAST(DataHoraTs AS TIME) >= TIME '{start}' OR CAST(DataHoraTs AS TIME) < TIME '{end}') OR CAST(strftime(DataHoraTs, '%w') AS INTEGER) IN (0, 6))"
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

    def generate_report(self):
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
            base_name = safe_name(report_type + "_" + stamp)

            if report_type in ("ociosidade_diario", "ociosidade_mensal"):
                out_path = reports_dir / f"{base_name}.{'csv' if self.report_format.get() == 'csv' else 'xlsx'}"
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

            if self.report_format.get() == "csv":
                out_path = reports_dir / f"{base_name}.csv"
                con.execute(f"COPY ({query}) TO '{sql_path(out_path)}' (HEADER, DELIMITER ';')", params)
            else:
                out_path = reports_dir / f"{base_name}.xlsx"
                self._export_query_to_xlsx(con, query, params, columns, out_path, report_name, total)

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

    def _export_query_to_xlsx(self, con, query: str, params: list, columns: list, out_path: Path, report_name: str, total_rows: int, color_weekend: bool = False):
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


def main():
    root = tk.Tk()
    MacroLockAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
