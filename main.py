import calendar
import argparse
import copy
import csv
import hashlib
import json
import os
import platform
import re
import subprocess
import sys
import tempfile
import threading
import time
import unicodedata
import uuid
import webbrowser
import base64
import hmac
from datetime import datetime, timedelta
from tkinter import Tk, Toplevel, StringVar, DoubleVar, BooleanVar, IntVar, END, PhotoImage, Canvas, Text, Listbox, Button, Frame, Label, Entry
from tkinter import ttk, messagebox, filedialog, colorchooser, font as tkfont

from lugest_desktop.legacy import app_misc_actions
from lugest_desktop.legacy import clientes_actions
from lugest_desktop.legacy import clientes_rooting
from lugest_desktop.legacy import encomendas_actions
from lugest_desktop.legacy import materia_actions
from lugest_desktop.legacy import menu_rooting
from lugest_desktop.legacy import ne_expedicao_actions
from lugest_desktop.legacy import operador_ordens_actions
from lugest_desktop.legacy import orc_actions
from lugest_desktop.legacy import plan_actions
from lugest_desktop.legacy import produtos_actions
from lugest_desktop.legacy import qualidade_actions
from lugest_desktop.legacy import ui_build_blocks
from lugest_infra.storage import files as lugest_storage

# UI mode: auto | ttk | custom
# Default em "custom" (visual original).
UI_MODE = str(os.environ.get("LUGEST_UI_MODE", "custom") or "custom").strip().lower()
if UI_MODE not in ("auto", "ttk", "custom"):
    UI_MODE = "auto"
_UI_CUSTOM_FLAGS = (
    "USE_CUSTOM_LOGIN",
    "USE_CUSTOM_MENU",
    "USE_CUSTOM_STOCK",
    "USE_CUSTOM_ENC",
    "USE_CUSTOM_PLANO",
    "USE_CUSTOM_QUALIDADE",
    "USE_CUSTOM_ORC",
    "USE_CUSTOM_OP",
    "USE_CUSTOM_OF",
    "USE_CUSTOM_EXPORT",
    "USE_CUSTOM_PROD",
    "USE_CUSTOM_NE",
)


def _apply_ui_mode_env():
    if UI_MODE == "auto":
        return
    forced = "1" if UI_MODE == "custom" else "0"
    for key in _UI_CUSTOM_FLAGS:
        os.environ[key] = forced


_apply_ui_mode_env()

# customtkinter (opcional para o ecra de login)
try:
    import customtkinter as ctk  # type: ignore
    CUSTOM_TK_AVAILABLE = UI_MODE != "ttk"
except Exception:
    CUSTOM_TK_AVAILABLE = False

try:
    import pymysql  # type: ignore
    from pymysql.cursors import DictCursor  # type: ignore
    MYSQL_AVAILABLE = True
except Exception:
    pymysql = None  # type: ignore
    DictCursor = None  # type: ignore
    MYSQL_AVAILABLE = False

STOCK_AMARELO = 10.0
STOCK_VERMELHO = 5.0
COND_PAGAMENTO_OPCOES = ["Pronto Pagamento", "30 dias", "60 dias"]
MATERIAIS_PRESET = [
    "S235JR",
    "S275JR",
    "S355JR",
    "S355J2+N",
    "S420MC",
    "DX51D+Z",
    "DX52D+Z",
    "Z275",
    "AISI 304",
    "AISI 304L",
    "AISI 304L 2B",
    "AISI 304L BA",
    "AISI 304L Escovado",
    "AISI 316",
    "AISI 316L",
    "AISI 316L 2B",
    "AISI 316L BA",
    "AISI 316Ti",
    "AISI 430",
    "AISI 441",
    "AISI 310S",
    "AISI 321",
    "EN AW-1050A",
    "EN AW-5754 H111",
    "EN AW-5083 H111",
    "EN AW-6060 T66",
    "EN AW-6061 T6",
    "EN AW-6082 T6",
    "CuZn37 (Latao)",
    "Cobre ETP",
]
INOX_ACABAMENTOS = ["Escovado", "2B", "LQ"]
ESPESSURAS_PRESET = [0.5] + [float(i) for i in range(1, 26)]
LOCALIZACOES_PRESET = [f"RACK01-GV{str(i).zfill(2)}" for i in range(1, 11)] + [f"RACK02-GV{str(i).zfill(2)}" for i in range(1, 11)]
PLANO_CORES = ["#fbecee", "#fde2e4", "#fff0f2", "#fce8ea", "#ffe6e9", "#f7dfe3"]
LOGO_PATH = "logo.png"
OFF_OPERACOES_DISPONIVEIS = [
    "Corte Laser",
    "Quinagem",
    "Serralharia",
    "Roscagem",
    "Soldadura",
    "Lacagem",
    "Pintura",
    "Maquinacao",
    "Montagem",
    "Embalamento",
]
OFF_OPERACAO_OBRIGATORIA = "Embalamento"
PLANEAMENTO_OPERACOES_DISPONIVEIS = [
    "Corte Laser",
    "Quinagem",
    "Serralharia",
    "Maquinacao",
    "Roscagem",
    "Lacagem",
    "Montagem",
]
ORC_LINE_TYPE_PIECE = "peca_fabricada"
ORC_LINE_TYPE_PRODUCT = "produto_stock"
ORC_LINE_TYPE_SERVICE = "servico_montagem"
BASE_DIR = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
_ACTIVE_LUGEST_ENV_PATH = ""


def _load_env_file(path, override=False):
    try:
        if not path or not os.path.exists(path):
            return False
        loaded_any = False
        for raw in open(path, "r", encoding="utf-8").read().splitlines():
            line = str(raw or "").strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip().lstrip("\ufeff")
            value = value.strip().strip('"').strip("'")
            if key and (override or key not in os.environ):
                os.environ[key] = value
                loaded_any = True
        return loaded_any
    except Exception:
        return False


def _resolve_primary_env_file():
    candidates = [os.path.join(BASE_DIR, "lugest.env")]
    base_parent = os.path.dirname(os.path.abspath(BASE_DIR))
    if base_parent and os.path.normcase(os.path.abspath(base_parent)) != os.path.normcase(os.path.abspath(BASE_DIR)):
        candidates.append(os.path.join(base_parent, "lugest.env"))
    cwd_candidate = os.path.join(os.getcwd(), "lugest.env")
    normalized = {os.path.normcase(os.path.abspath(path)) for path in candidates if path}
    if os.path.normcase(os.path.abspath(cwd_candidate)) not in normalized:
        candidates.append(cwd_candidate)
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return ""


_ACTIVE_LUGEST_ENV_PATH = _resolve_primary_env_file()
if _ACTIVE_LUGEST_ENV_PATH:
    _load_env_file(_ACTIVE_LUGEST_ENV_PATH, override=True)


def active_lugest_env_path():
    return str(_ACTIVE_LUGEST_ENV_PATH or "").strip()


def mysql_runtime_source_label():
    env_path = active_lugest_env_path()
    if env_path:
        return env_path
    return "Variáveis de ambiente do Windows/processo"


def shared_storage_root_label():
    return str(lugest_storage.shared_storage_label(BASE_DIR))


def shared_storage_root_path():
    return str(lugest_storage.shared_storage_root(BASE_DIR))

DEFAULT_PRIMARY_RED = "#ba2d3d"
DEFAULT_PRIMARY_RED_HOVER = "#a32035"
DEFAULT_THEME_HEADER_BG = "#9b2233"
DEFAULT_THEME_HEADER_ACTIVE = "#992233"
DEFAULT_THEME_SELECT_BG = "#fde2e4"
DEFAULT_THEME_SELECT_FG = "#7a0f1a"

CTK_PRIMARY_RED = DEFAULT_PRIMARY_RED
CTK_PRIMARY_RED_HOVER = DEFAULT_PRIMARY_RED_HOVER
CTK_PRIMARY_TEXT = "#ffffff"
THEME_HEADER_BG = DEFAULT_THEME_HEADER_BG
THEME_HEADER_ACTIVE = DEFAULT_THEME_HEADER_ACTIVE
THEME_SELECT_BG = DEFAULT_THEME_SELECT_BG
THEME_SELECT_FG = DEFAULT_THEME_SELECT_FG

UI_FONT_FAMILY = "Segoe UI"
UI_FONT_MONO_FAMILY = "Consolas"
_AVAILABLE_FONT_FAMILIES = set()
_FONT_COMPAT_READY = False


def _ui_font(size=11, weight="normal", family=None):
    fam = str(family or UI_FONT_FAMILY or "Segoe UI").strip()
    try:
        size_i = int(size)
    except Exception:
        size_i = 11
    w = str(weight or "normal").strip().lower()
    if w and w != "normal":
        return (fam, size_i, w)
    return (fam, size_i)


def _ui_font_available(name):
    return str(name or "").strip().lower() in _AVAILABLE_FONT_FAMILIES


def _normalize_font_value(value):
    if isinstance(value, tuple) and value:
        fam = value[0]
        if isinstance(fam, str) and fam.strip() and not _ui_font_available(fam):
            return (UI_FONT_FAMILY, *value[1:])
        return value
    return value


def _patch_ctk_font_fallback():
    if not CUSTOM_TK_AVAILABLE or ctk is None:
        return
    widget_names = (
        "CTkLabel",
        "CTkButton",
        "CTkEntry",
        "CTkTextbox",
        "CTkComboBox",
        "CTkCheckBox",
        "CTkSwitch",
        "CTkSegmentedButton",
    )
    for name in widget_names:
        cls = getattr(ctk, name, None)
        if cls is None:
            continue
        if getattr(cls, "__lugest_font_fallback_patched__", False):
            continue
        orig_init = cls.__init__

        def _patched_init(self, *args, __orig=orig_init, **kwargs):
            if "font" in kwargs:
                kwargs["font"] = _normalize_font_value(kwargs.get("font"))
            return __orig(self, *args, **kwargs)

        cls.__init__ = _patched_init
        cls.__lugest_font_fallback_patched__ = True


def _configure_font_compatibility(root):
    global UI_FONT_FAMILY, UI_FONT_MONO_FAMILY, _AVAILABLE_FONT_FAMILIES, _FONT_COMPAT_READY
    if _FONT_COMPAT_READY:
        return
    families = set()
    try:
        families = {str(f).strip().lower() for f in tkfont.families(root)}
    except Exception:
        families = set()
    _AVAILABLE_FONT_FAMILIES = families

    default_family = UI_FONT_FAMILY
    try:
        default_family = str(tkfont.nametofont("TkDefaultFont").cget("family") or UI_FONT_FAMILY).strip()
    except Exception:
        pass

    env_pref = str(os.environ.get("LUGEST_FONT_FAMILY", "auto") or "auto").strip()
    pref_candidates = []
    if env_pref and env_pref.lower() not in ("auto", "default"):
        pref_candidates.append(env_pref)
    pref_candidates.extend(["Segoe UI", "Calibri", "Arial", "Tahoma", "Verdana", default_family])

    UI_FONT_FAMILY = default_family
    for cand in pref_candidates:
        key = str(cand or "").strip().lower()
        if key and key in families:
            UI_FONT_FAMILY = str(cand).strip()
            break

    mono_pref = str(os.environ.get("LUGEST_FONT_MONO", "") or "").strip()
    mono_candidates = []
    if mono_pref:
        mono_candidates.append(mono_pref)
    mono_candidates.extend(["Consolas", "Cascadia Mono", "Courier New", "DejaVu Sans Mono", UI_FONT_FAMILY])
    UI_FONT_MONO_FAMILY = UI_FONT_FAMILY
    for cand in mono_candidates:
        key = str(cand or "").strip().lower()
        if key and key in families:
            UI_FONT_MONO_FAMILY = str(cand).strip()
            break

    try:
        for name in (
            "TkDefaultFont",
            "TkTextFont",
            "TkHeadingFont",
            "TkMenuFont",
            "TkIconFont",
            "TkCaptionFont",
            "TkSmallCaptionFont",
            "TkTooltipFont",
            "TkFixedFont",
        ):
            f = tkfont.nametofont(name)
            f.configure(family=UI_FONT_MONO_FAMILY if name == "TkFixedFont" else UI_FONT_FAMILY)
    except Exception:
        pass

    try:
        root.option_add("*Font", _ui_font(10))
        root.option_add("*TCombobox*Listbox*Font", _ui_font(10))
        root.option_add("*Menu*Font", _ui_font(10))
    except Exception:
        pass

    _patch_ctk_font_fallback()
    _FONT_COMPAT_READY = True


def _set_ctk_button_defaults_red():
    if not CUSTOM_TK_AVAILABLE or ctk is None:
        return
    if getattr(ctk.CTkButton, "__lugest_red_defaults_patched__", False):
        return
    _orig_init = ctk.CTkButton.__init__

    def _patched_init(self, *args, **kwargs):
        fg = str(kwargs.get("fg_color", "") or "").strip().lower()
        hv = str(kwargs.get("hover_color", "") or "").strip().lower()
        if not fg or fg in {DEFAULT_PRIMARY_RED, "#ba2d3d"}:
            kwargs["fg_color"] = CTK_PRIMARY_RED
        if not hv or hv in {DEFAULT_PRIMARY_RED_HOVER, "#a32035"}:
            kwargs["hover_color"] = CTK_PRIMARY_RED_HOVER
        kwargs.setdefault("text_color", CTK_PRIMARY_TEXT)
        return _orig_init(self, *args, **kwargs)

    ctk.CTkButton.__init__ = _patched_init
    ctk.CTkButton.__lugest_red_defaults_patched__ = True


def _normalize_hex_color(value, fallback=""):
    txt = str(value or "").strip()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", txt):
        return txt.lower()
    return str(fallback or "").strip().lower()


def _shade_hex_color(value, factor):
    base = _normalize_hex_color(value, DEFAULT_PRIMARY_RED)
    if not base:
        base = DEFAULT_PRIMARY_RED
    try:
        r = int(base[1:3], 16)
        g = int(base[3:5], 16)
        b = int(base[5:7], 16)
        f = max(0.0, min(1.0, float(factor)))
        rr = max(0, min(255, int(r * f)))
        gg = max(0, min(255, int(g * f)))
        bb = max(0, min(255, int(b * f)))
        return f"#{rr:02x}{gg:02x}{bb:02x}"
    except Exception:
        return _normalize_hex_color(value, DEFAULT_PRIMARY_RED)


def _mix_hex_color(base_hex, mix_hex="#ffffff", mix_ratio=0.82):
    base = _normalize_hex_color(base_hex, DEFAULT_PRIMARY_RED) or DEFAULT_PRIMARY_RED
    mix = _normalize_hex_color(mix_hex, "#ffffff") or "#ffffff"
    r = max(0.0, min(1.0, float(mix_ratio)))
    try:
        br, bg, bb = int(base[1:3], 16), int(base[3:5], 16), int(base[5:7], 16)
        mr, mg, mb = int(mix[1:3], 16), int(mix[3:5], 16), int(mix[5:7], 16)
        rr = int((br * (1.0 - r)) + (mr * r))
        gg = int((bg * (1.0 - r)) + (mg * r))
        bb2 = int((bb * (1.0 - r)) + (mb * r))
        return f"#{rr:02x}{gg:02x}{bb2:02x}"
    except Exception:
        return DEFAULT_THEME_SELECT_BG


def apply_primary_theme_color(color_hex):
    global CTK_PRIMARY_RED, CTK_PRIMARY_RED_HOVER, THEME_HEADER_BG, THEME_HEADER_ACTIVE, THEME_SELECT_BG, THEME_SELECT_FG
    base = _normalize_hex_color(color_hex, DEFAULT_PRIMARY_RED) or DEFAULT_PRIMARY_RED
    CTK_PRIMARY_RED = base
    CTK_PRIMARY_RED_HOVER = _shade_hex_color(base, 0.84)
    THEME_HEADER_BG = _shade_hex_color(base, 0.82)
    THEME_HEADER_ACTIVE = _shade_hex_color(base, 0.76)
    THEME_SELECT_BG = _mix_hex_color(base, "#ffffff", 0.82)
    THEME_SELECT_FG = _shade_hex_color(base, 0.58)
ORC_CONDICOES_DEFAULT = [
    "Prazo de validade do orcamento: 30 dias",
    "Prazo de entrega: a confirmar com o cliente",
    "Condicoes de pagamento: conforme acordado",
    "Valores sem IVA",
]
ORC_EMPRESA_INFO_RODAPE = [
    "Barcelbal - Balancas e Basculas, S.A.",
    "Rua dos Canteiros, n. 53 - Adaufe",
    "Tel: 253 606 590  |  Email: geral@barcelbal.pt",
    "NIF: 502403843  |  Capital Social: 100.000 EUR",
]
ORC_NOTAS_DEFAULT = [
    "PROPOSTA RETIFICADA PARA ESPESSURAS DEFINIDAS PELO CLIENTE.",
    "- Foi considerado servico de corte laser.",
    "- A materia-prima e transporte sao do encargo do cliente.",
]
ORC_CONDICOES_GERAIS = [
    "- Prazo de validade do orcamento: 30 dias, salvo rutura de stock.",
    "- Em caso de adjudicacao parcial da proposta os precos serao revistos.",
    "- Prazo de entrega: a combinar com o departamento de planeamento.",
    "- Os precos nao incluem IVA.",
    "- Valor minimo por encomenda: 25 EUR.",
]
ORC_LEGENDA_OPERACOES = [
    "CL - CORTE LASER",
    "Q - QUINAGEM",
    "R - ROSCAGEM",
    "F - FURO MANUAL",
    "S - SOLDADURA",
    "P - PINTURA",
]
ORC_RECLAMACOES = (
    "A verificacao de eventuais anomalias (quantidade, qualidade, dimensoes, etc.) tera que ser comunicada, por escrito, "
    "no periodo de 20 dias apos a rececao dos nossos produtos. Apos este prazo, nao poderemos aceitar reclamacoes, a menos "
    "que se refiram a problemas nao detetaveis no acto da rececao. Todas as reclamacoes serao objeto de um relatorio de nao conformidade."
)
ORC_DEVOLUCOES = (
    "Todas as pecas rejeitadas devem ser devolvidas para a devida analise/tratamento da reclamacao. "
    "Devem ser acompanhadas da guia de transporte de devolucao indicando o n. da nossa guia de transporte."
)

ORC_LOGO_CANDIDATES = [
    LOGO_PATH,
    os.path.join(BASE_DIR, "logo.png"),
    os.path.join(BASE_DIR, "logo.jpg"),
    os.path.join(BASE_DIR, "Logos", "logo.png"),
    os.path.join(BASE_DIR, "Logos", "image (1).jpg"),
    os.path.join(BASE_DIR, "Logos", "image.jpg"),
    os.path.join(BASE_DIR, "Logos", "logo(1).jpg"),
]
BRANDING_FILE = "lugest_branding.json"
_BRANDING_CACHE = None
_LAST_SAVE_FINGERPRINT = None
_LAST_SAVE_TS = 0.0
_PENDING_SAVE_DATA = None
_SAVE_CHANGE_TOKEN = 0
_LAST_SAVED_TOKEN = 0
_SAVE_MIN_INTERVAL_SEC = float(os.environ.get("LUGEST_SAVE_MIN_INTERVAL_SEC", "0.0"))
try:
    _SAVE_HEARTBEAT_MS = int(float(os.environ.get("LUGEST_SAVE_HEARTBEAT_MS", "5000")))
except Exception:
    _SAVE_HEARTBEAT_MS = 5000
if _SAVE_HEARTBEAT_MS < 900:
    _SAVE_HEARTBEAT_MS = 900
_ASYNC_SAVE_ENABLED = str(os.environ.get("LUGEST_ASYNC_SAVE", "0") or "0").strip().lower() not in ("0", "false", "no", "off")
_ASYNC_SAVE_LOCK = threading.Lock()
_ASYNC_SAVE_EVENT = threading.Event()
_ASYNC_SAVE_STOP = threading.Event()
_ASYNC_SAVE_THREAD = None
_ASYNC_SAVE_PENDING_DATA = None
_ASYNC_SAVE_PENDING_FP = ""
_ASYNC_SAVE_PENDING_TOKEN = 0
_ASYNC_SAVE_IN_PROGRESS = False
_ASYNC_SAVE_LAST_ERROR = ""
_MYSQL_RUNTIME_SCHEMA_FLAGS = set()
_MYSQL_SCHEMA_CACHE_LOCK = threading.Lock()
_MYSQL_TABLES_CACHE = None
_MYSQL_COLUMNS_CACHE = {}
_MYSQL_INDEXES_CACHE = {}
_MYSQL_SAVE_LOCK_TIMEOUT_SEC = max(5, int(float(os.environ.get("LUGEST_SAVE_LOCK_TIMEOUT_SEC", "45"))))
_MYSQL_SAVE_RETRY_COUNT = max(1, int(float(os.environ.get("LUGEST_SAVE_RETRY_COUNT", "3"))))

PROD_CATEGORIAS = [
    "Inox",
    "Aluminio",
    "Ferro",
    "EPIs",
    "Eletronica",
    "Pneumatica",
    "Hidraulica",
    "Fixacao",
    "Consumiveis",
    "Corte Laser",
    "Quinagem",
    "Soldadura",
    "Maquinacao",
    "Ferramentas",
    "Rolamentos & Transmissao",
    "Motores & Redutores",
    "Plasticos Tecnicos",
    "Vedacao & Borracha",
    "MRO & Manutencao",
    "Outros",
]
MATERIA_FORMATOS = ["Chapa", "Tubo", "Perfil", "Cantoneira", "Barra", "Varão nervurado"]
PROD_SUBCATS = [
    "Tubo",
    "Chapa",
    "Perfil",
    "Varao",
    "Luvas",
    "Capacetes",
    "Mascaras",
    "Botas",
    "Oculos",
    "Sensores",
    "Automacao",
    "Cablagem",
    "Valvulas",
    "Cilindros",
    "Ligacoes",
    "Mangueiras",
    "Parafusos",
    "Porcas",
    "Anilhas",
    "Abrasivos",
    "Quimicos",
    "Embalagem",
    "Consumiveis",
    "Gases",
    "Ferramentas",
    "Rolamentos",
    "Correntes",
    "Pinhoes e polias",
    "Motores",
    "Redutores",
    "Variadores",
    "Plasticos",
    "Vedacao",
    "Borracha tecnica",
    "MRO",
    "Outros",
]
PROD_TIPOS = [
    "Redondo", "Quadrado", "Retangular", "Sanitario",
    "Escovada", "Polida", "Decapada", "Perfurada",
    "UPN", "IPE", "HEA", "HEB", "Cantoneira",
    "Corte", "Soldadura", "Nitrilo", "Termicas",
    "Industrial", "Eletrico", "Com viseira",
    "FFP2", "FFP3", "Respiratoria",
    "S1P", "S3", "Panoramicos",
    "PLC", "HMI", "Reles", "Fontes",
    "2 vias", "3 vias", "5 vias", "Proporcional",
    "Compacto", "ISO", "Guiado",
    "Allen", "Sextavado", "Escareado", "Auto perfurante",
    "Disco corte", "Disco flap", "Lixa", "Escova",
    "Oxigenio", "Azoto", "Argon", "CO2",
    "Puncao", "Matriz V", "Arame MIG", "Eletrodo", "Vareta TIG",
    "Topo", "Esferica", "Chanfrar", "Broca HSS", "Carbureto",
    "Paquimetro", "Micrometro", "Torquimetro",
    "Esferas", "Rolos", "Pinhao", "Polia", "Correia",
    "Monofasico", "Trifasico", "Servo", "Planetario",
    "PEAD", "PVC", "PTFE", "POM", "PEEK",
    "O-ring", "Retentor", "EPDM", "NBR", "Silicone",
    "Massa", "Oleo", "Disjuntor", "Contator",
    "Outros",
]
PROD_UNIDS = ["UN", "M", "KG", "CX"]
PDF_TEXT_REPLACEMENTS = {
    "\u00e1": "a", "\u00e0": "a", "\u00e3": "a", "\u00e2": "a", "\u00e4": "a",
    "\u00c1": "A", "\u00c0": "A", "\u00c3": "A", "\u00c2": "A", "\u00c4": "A",
    "\u00e9": "e", "\u00ea": "e", "\u00e8": "e", "\u00eb": "e",
    "\u00c9": "E", "\u00ca": "E", "\u00c8": "E", "\u00cb": "E",
    "\u00ed": "i", "\u00ec": "i", "\u00ee": "i", "\u00ef": "i",
    "\u00cd": "I", "\u00cc": "I", "\u00ce": "I", "\u00cf": "I",
    "\u00f3": "o", "\u00f4": "o", "\u00f5": "o", "\u00f2": "o", "\u00f6": "o",
    "\u00d3": "O", "\u00d4": "O", "\u00d5": "O", "\u00d2": "O", "\u00d6": "O",
    "\u00fa": "u", "\u00f9": "u", "\u00fb": "u", "\u00fc": "u",
    "\u00da": "U", "\u00d9": "U", "\u00db": "U", "\u00dc": "U",
    "\u00e7": "c", "\u00c7": "C",
}

DEFAULT_DATA = {
    "users": [],
    "seq": {"encomenda": 1, "cliente": 1, "ref_interna": {}, "produto": 1, "ne": 1},
    "clientes": [],
    "materiais": [],
    "encomendas": [],
    "transportes": [],
    "transportes_tarifarios": [],
    "qualidade": [],
    "quality_nonconformities": [],
    "quality_documents": [],
    "audit_log": [],
    "refs": [],
    "materiais_hist": [],
    "espessuras_hist": [],
    "stock_log": [],
    "plano": [],
    "peca_hist": {},
    "rejeitadas_hist": [],
    "operadores": ["Operador 1"],
    "orcamentos": [],
    "orc_seq": 1,
    "orc_refs": {},
    "conjuntos_modelo": [],
    "of_seq": 1,
    "opp_seq": 1,
    "orcamentistas": ["Orçamentista 1"],
    "produtos": [],
    "notas_encomenda": [],
    "expedicoes": [],
    "faturacao": [],
    "at_series": [],
    "exp_seq": 1,
    "fornecedores": [],
    "produtos_mov": [],
    "plano_hist": [],
    "op_eventos": [],
    "op_paragens": [],
    "postos_trabalho": [],
    "operador_posto_map": {},
    "tempos_operacao_planeada_min": {},
    "workcenter_catalog": [],
    "plano_bloqueios": [],
}

USE_MYSQL_STORAGE = True
MYSQL_HOST = os.environ.get("LUGEST_DB_HOST", "127.0.0.1")
def _env_int(name, default):
    raw = str(os.environ.get(name, default) or "").strip()
    try:
        return int(raw)
    except Exception:
        return int(default)


MYSQL_PORT = _env_int("LUGEST_DB_PORT", 3306)
MYSQL_USER = os.environ.get("LUGEST_DB_USER", "")
MYSQL_PASSWORD = os.environ.get("LUGEST_DB_PASS", "")
MYSQL_DB_NAME = os.environ.get("LUGEST_DB_NAME", "lugest")
_MYSQL_SCHEMA_SYNCED = False
_RUNTIME_DATA_REF = None

PASSWORD_HASH_ALGO = "pbkdf2_sha256"
try:
    PASSWORD_HASH_ITERATIONS = max(240000, int(os.environ.get("LUGEST_PASSWORD_HASH_ITERATIONS", "260000") or "260000"))
except Exception:
    PASSWORD_HASH_ITERATIONS = 260000
WEAK_PASSWORD_VALUES = {
    "",
    "admin",
    "1234",
    "12345",
    "123456",
    "12345678",
    "123123",
    "password",
    "producao",
    "qualidade",
    "planeamento",
    "orcamentista",
    "operador",
    "test",
}


def _copy_default_data():
    return json.loads(json.dumps(DEFAULT_DATA))


MOJIBAKE_KEY_MAP = {
    "Operações": "Operacoes",
    "Observações": "Observacoes",
    "Observacões": "Observacoes",
}


def _clip(value, max_len=None):
    if value is None:
        return None
    txt = str(value)
    if max_len and len(txt) > max_len:
        return txt[:max_len]
    return txt


def _repair_mojibake_text(value):
    txt = "" if value is None else str(value)
    if any(tok in txt for tok in ("Ã", "Â", "â", "€", "™", "œ", "ž", "Ÿ")):
        for enc in ("cp1252", "latin1"):
            try:
                fixed = txt.encode(enc, errors="strict").decode("utf-8", errors="strict")
                if fixed:
                    txt = fixed
                    break
            except Exception:
                continue
    return txt.replace("\u00a0", " ")


def _repair_mojibake_structure(value):
    if isinstance(value, dict):
        fixed = {}
        for key, val in value.items():
            new_key = _repair_mojibake_text(key) if isinstance(key, str) else key
            if isinstance(new_key, str):
                new_key = MOJIBAKE_KEY_MAP.get(new_key, new_key)
            new_val = _repair_mojibake_structure(val)
            if new_key in fixed and isinstance(fixed[new_key], dict) and isinstance(new_val, dict):
                merged = dict(fixed[new_key])
                merged.update(new_val)
                fixed[new_key] = merged
            elif new_key in fixed and isinstance(fixed[new_key], list) and isinstance(new_val, list):
                fixed[new_key] = list(fixed[new_key]) + list(new_val)
            else:
                fixed[new_key] = new_val
        return fixed
    if isinstance(value, list):
        return [_repair_mojibake_structure(v) for v in value]
    if isinstance(value, str):
        return _repair_mojibake_text(value)
    return value


def _normalize_role_name(value):
    role = _repair_mojibake_text(value).strip()
    if role in ("Orcamentista", "Orçamentista"):
        return "Orcamentista"
    return role


def is_password_hash(value):
    txt = str(value or "").strip()
    parts = txt.split("$")
    return len(parts) == 4 and parts[0] == PASSWORD_HASH_ALGO and str(parts[1]).isdigit()


def hash_password(raw_password, iterations=None):
    password_txt = str(raw_password or "")
    try:
        iter_num = max(120000, int(iterations or PASSWORD_HASH_ITERATIONS))
    except Exception:
        iter_num = PASSWORD_HASH_ITERATIONS
    salt = os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password_txt.encode("utf-8"), salt, iter_num)
    salt_txt = base64.urlsafe_b64encode(salt).decode("ascii").rstrip("=")
    digest_txt = base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")
    return f"{PASSWORD_HASH_ALGO}${iter_num}${salt_txt}${digest_txt}"


def verify_password(candidate_password, stored_password):
    candidate_txt = str(candidate_password or "")
    stored_txt = str(stored_password or "").strip()
    if not stored_txt:
        return False
    if not is_password_hash(stored_txt):
        return hmac.compare_digest(candidate_txt, stored_txt)
    try:
        _, iter_txt, salt_txt, digest_txt = stored_txt.split("$", 3)
        salt_padding = "=" * (-len(salt_txt) % 4)
        digest_padding = "=" * (-len(digest_txt) % 4)
        salt = base64.urlsafe_b64decode(f"{salt_txt}{salt_padding}".encode("ascii"))
        expected = base64.urlsafe_b64decode(f"{digest_txt}{digest_padding}".encode("ascii"))
        calc = hashlib.pbkdf2_hmac("sha256", candidate_txt.encode("utf-8"), salt, max(1, int(iter_txt)))
        return hmac.compare_digest(calc, expected)
    except Exception:
        return False


def password_strength_issues(username, password):
    raw_password = str(password or "")
    if not raw_password:
        return ["nao pode ficar vazia"]
    if is_password_hash(raw_password):
        return []
    pwd = raw_password.strip()
    pwd_lower = pwd.lower()
    username_txt = str(username or "").strip().lower()
    issues = []
    if len(pwd) < 10:
        issues.append("ter pelo menos 10 caracteres")
    classes = sum(
        [
            any(ch.islower() for ch in pwd),
            any(ch.isupper() for ch in pwd),
            any(ch.isdigit() for ch in pwd),
            any(not ch.isalnum() for ch in pwd),
        ]
    )
    if classes < 3:
        issues.append("misturar letras, numeros e simbolos")
    if pwd_lower in WEAK_PASSWORD_VALUES or (username_txt and pwd_lower == username_txt):
        issues.append("nao pode usar um valor previsivel")
    return issues


def is_weak_password_value(username, password):
    return bool(password_strength_issues(username, password))


def validate_local_password(username, password):
    issues = password_strength_issues(username, password)
    if issues:
        raise ValueError("A password deve " + "; ".join(issues) + ".")
    return str(password or "")


def normalize_password_for_storage(username, password, require_strong=False):
    stored_txt = str(password or "")
    if not stored_txt:
        return ""
    if is_password_hash(stored_txt):
        return stored_txt
    if require_strong:
        validate_local_password(username, stored_txt)
    return hash_password(stored_txt)


def bootstrap_user_payloads():
    username = _clip(os.environ.get("LUGEST_BOOTSTRAP_ADMIN_USERNAME", ""), 50) or ""
    password = str(os.environ.get("LUGEST_BOOTSTRAP_ADMIN_PASSWORD", "") or "")
    role = _normalize_role_name(_clip(os.environ.get("LUGEST_BOOTSTRAP_ADMIN_ROLE", "Admin"), 50)) or "Admin"
    username = str(username).strip()
    if not username or not password:
        return []
    if not is_password_hash(password) and is_weak_password_value(username, password):
        return []
    return [
        {
            "username": username,
            "password": _clip(normalize_password_for_storage(username, password, require_strong=False), 255),
            "role": role,
        }
    ]


def find_local_user(data, username):
    target = str(username or "").strip().lower()
    if not target:
        return None
    for row in list((data or {}).get("users", []) or []):
        if str(row.get("username", "") or "").strip().lower() == target:
            return row
    return None


def authenticate_local_user(data, username, password, persist_upgrade=True):
    user = find_local_user(data, username)
    if not isinstance(user, dict):
        return None
    raw_password = str(password or "")
    stored_password = str(user.get("password", "") or "")
    if not verify_password(raw_password, stored_password):
        return None
    if persist_upgrade and stored_password and not is_password_hash(stored_password):
        try:
            user["password"] = normalize_password_for_storage(str(user.get("username", "") or "").strip(), raw_password)
            save_data(data, force=True)
        except Exception:
            pass
    return user


def build_authenticated_user_session(user, password=""):
    session = dict(user or {})
    session.pop("_session_password", None)
    raw_password = str(password or "")
    if raw_password:
        session["_session_password"] = raw_password
    return session


def _to_bool(value):
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    txt = _repair_mojibake_text(value).strip().lower()
    if txt in ("1", "true", "sim", "yes", "y", "on"):
        return True
    if txt in ("0", "false", "nao", "não", "no", "n", "off", ""):
        return False
    return True


def _to_num(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    txt = str(value).strip().replace(",", ".")
    if not txt:
        return None
    try:
        return float(txt)
    except Exception:
        return None


def _to_mysql_datetime(value):
    txt = str(value or "").strip()
    if not txt:
        return None
    try:
        dt = datetime.fromisoformat(txt.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    fmts = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d")
    for fmt in fmts:
        try:
            dt = datetime.strptime(txt, fmt)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            continue
    return None


def _to_mysql_date(value):
    txt = str(value or "").strip()
    if not txt:
        return None
    try:
        dt = datetime.fromisoformat(txt.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            dt = datetime.strptime(txt, fmt)
            return dt.strftime("%Y-%m-%d")
        except Exception:
            continue
    return None


def _derive_year_from_values(*values, default=None):
    for raw in values:
        txt = str(raw or "").strip()
        if len(txt) >= 4 and txt[:4].isdigit():
            try:
                y = int(txt[:4])
                if 1900 <= y <= 2999:
                    return y
            except Exception:
                pass
        m = re.search(r"(19|20)\d{2}", txt)
        if m:
            try:
                y = int(m.group(0))
                if 1900 <= y <= 2999:
                    return y
            except Exception:
                pass
    if default is None:
        return datetime.now().year
    try:
        return int(default)
    except Exception:
        return datetime.now().year


def _extract_cliente_codigo(value, data):
    clientes = data.get("clientes", []) if isinstance(data, dict) else []
    codes = {str(c.get("codigo", "")).strip() for c in clientes if isinstance(c, dict)}

    def _from_text(txt):
        txt = str(txt or "").strip()
        if not txt:
            return None
        m = re.match(r"^(CL\d{4})", txt)
        if m:
            return m.group(1)
        if txt in codes:
            return txt
        for c in clientes:
            if not isinstance(c, dict):
                continue
            if str(c.get("nome", "")).strip() == txt and c.get("codigo"):
                return c.get("codigo")
        return None

    if isinstance(value, dict):
        for key in ("codigo", "cliente", "id"):
            cod = _from_text(value.get(key))
            if cod:
                return cod
        nome = str(value.get("nome", "") or value.get("empresa", "") or "").strip()
        nif = str(value.get("nif", "") or "").strip()
        email = str(value.get("email", "") or "").strip().lower()
        for c in clientes:
            if not isinstance(c, dict):
                continue
            if nif and str(c.get("nif", "")).strip() == nif and c.get("codigo"):
                return c.get("codigo")
            if email and str(c.get("email", "")).strip().lower() == email and c.get("codigo"):
                return c.get("codigo")
            if nome and str(c.get("nome", "")).strip() == nome and c.get("codigo"):
                return c.get("codigo")
        return None

    return _from_text(value)


def _normalize_orc_cliente(value, data):
    cli = {
        "codigo": "",
        "nome": "",
        "empresa": "",
        "nif": "",
        "morada": "",
        "contacto": "",
        "email": "",
    }
    if isinstance(value, dict):
        for k in cli:
            cli[k] = str(value.get(k, "") or "").strip()
    else:
        txt = str(value or "").strip()
        if txt:
            cli["codigo"] = _extract_cliente_codigo(txt, data) or ""
            if not cli["codigo"]:
                cli["nome"] = txt

    cod = _extract_cliente_codigo(cli, data)
    ref = None
    for c in data.get("clientes", []):
        if not isinstance(c, dict):
            continue
        cc = str(c.get("codigo", "") or "").strip()
        if cod and cc == cod:
            ref = c
            break
    if ref is None:
        nome = str(cli.get("nome", "") or cli.get("empresa", "") or "").strip()
        nif = str(cli.get("nif", "") or "").strip()
        email = str(cli.get("email", "") or "").strip().lower()
        for c in data.get("clientes", []):
            if not isinstance(c, dict):
                continue
            if nif and str(c.get("nif", "")).strip() == nif:
                ref = c
                break
            if email and str(c.get("email", "")).strip().lower() == email:
                ref = c
                break
            if nome and str(c.get("nome", "")).strip() == nome:
                ref = c
                break
    if ref is not None:
        cli["codigo"] = str(ref.get("codigo", "") or cli.get("codigo", "") or "").strip()
        cli["nome"] = str(cli.get("nome", "") or ref.get("nome", "") or "").strip()
        cli["empresa"] = str(cli.get("empresa", "") or cli.get("nome", "") or ref.get("nome", "") or "").strip()
        cli["nif"] = str(cli.get("nif", "") or ref.get("nif", "") or "").strip()
        cli["morada"] = str(cli.get("morada", "") or ref.get("morada", "") or "").strip()
        cli["contacto"] = str(cli.get("contacto", "") or ref.get("contacto", "") or "").strip()
        cli["email"] = str(cli.get("email", "") or ref.get("email", "") or "").strip()
    else:
        if cod:
            cli["codigo"] = cod
        if not cli.get("empresa"):
            cli["empresa"] = cli.get("nome", "")
    return cli


def _extract_fornecedor_id(value, data):
    txt = str(value or "").strip()
    if not txt:
        return None
    supplier_rows = list(data.get("fornecedores", []) or [])
    ids = {str(f.get("id", "")).strip(): str(f.get("id", "")).strip() for f in supplier_rows if str(f.get("id", "")).strip()}
    ids_ci = {key.lower(): value for key, value in ids.items()}
    if txt in ids:
        return txt
    if txt.lower() in ids_ci:
        return ids_ci[txt.lower()]
    m = re.match(r"^(FOR-\d{4})\b", txt, re.IGNORECASE)
    if m:
        candidate = m.group(1).upper()
        if candidate in ids:
            return candidate
    if " - " in txt:
        _, supplier_name = [part.strip() for part in txt.split(" - ", 1)]
        txt = supplier_name or txt
    for f in data.get("fornecedores", []):
        if str(f.get("nome", "")).strip().lower() == txt.lower() and f.get("id"):
            return f.get("id")
    return None


def _get_localizacao(mat):
    for k in ("Localização", "Localizacao", "Localiza?o"):
        v = mat.get(k)
        if str(v or "").strip():
            return str(v)
    return ""


def _mysql_connect():
    _assert_mysql_runtime_ready()
    return pymysql.connect(
        host=MYSQL_HOST,
        port=MYSQL_PORT,
        user=MYSQL_USER,
        password=MYSQL_PASSWORD,
        database=MYSQL_DB_NAME,
        charset="utf8mb4",
        autocommit=False,
        cursorclass=DictCursor,
        connect_timeout=5,
        read_timeout=20,
        write_timeout=20,
    )


def _mysql_save_lock_name():
    db_name = str(MYSQL_DB_NAME or "").strip() or "lugest"
    return f"lugest:save:{db_name}"


def _mysql_named_lock_acquire(cur, name, timeout_sec=45):
    cur.execute("SELECT GET_LOCK(%s, %s) AS acquired", (str(name or ""), int(max(1, timeout_sec or 1))))
    row = cur.fetchone() or {}
    try:
        value = row.get("acquired") if isinstance(row, dict) else row[0]
    except Exception:
        value = None
    try:
        return int(value or 0) == 1
    except Exception:
        return False


def _mysql_named_lock_release(cur, name):
    try:
        cur.execute("SELECT RELEASE_LOCK(%s)", (str(name or ""),))
    except Exception:
        pass


def _mysql_retryable_write_error(ex):
    try:
        args = list(getattr(ex, "args", []) or [])
    except Exception:
        args = []
    if args:
        try:
            code = int(args[0])
        except Exception:
            code = None
        if code in (1205, 1213):
            return True
    text = str(ex or "").strip().lower()
    return ("deadlock" in text) or ("lock wait timeout" in text)


def _mysql_runtime_errors():
    errors = []
    if not USE_MYSQL_STORAGE:
        errors.append("O runtime atual exige MySQL como fonte de dados.")
    if not MYSQL_AVAILABLE:
        errors.append("PyMySQL nao esta instalado neste ambiente.")
    if not str(MYSQL_HOST or "").strip():
        errors.append("Falta definir LUGEST_DB_HOST no lugest.env.")
    if not isinstance(MYSQL_PORT, int) or MYSQL_PORT < 1 or MYSQL_PORT > 65535:
        errors.append("LUGEST_DB_PORT tem um valor invalido.")
    if not str(MYSQL_USER or "").strip():
        errors.append("Falta definir LUGEST_DB_USER no lugest.env.")
    if not str(MYSQL_PASSWORD or "").strip():
        errors.append("Falta definir LUGEST_DB_PASS no lugest.env.")
    if not str(MYSQL_DB_NAME or "").strip():
        errors.append("Falta definir LUGEST_DB_NAME no lugest.env.")
    return errors


def _assert_mysql_runtime_ready():
    errors = _mysql_runtime_errors()
    if errors:
        source = mysql_runtime_source_label()
        raise RuntimeError(
            "Configuração MySQL incompleta.\n\n"
            + "\n".join(f"- {item}" for item in errors)
            + f"\n\nOrigem da configuração: {source}"
        )


def format_mysql_runtime_error(exc, action="ligar ao servidor MySQL"):
    host = str(MYSQL_HOST or "").strip() or "-"
    db_name = str(MYSQL_DB_NAME or "").strip() or "-"
    user = str(MYSQL_USER or "").strip() or "-"
    try:
        port = int(MYSQL_PORT)
    except Exception:
        port = MYSQL_PORT
    source = mysql_runtime_source_label()
    details = [
        f"Falha ao {action}.",
        "",
        f"Host: {host}",
        f"Porta: {port}",
        f"Base de dados: {db_name}",
        f"Utilizador: {user}",
        f"Origem da configuração: {source}",
        "",
        "O que verificar:",
    ]
    if host in ("127.0.0.1", "localhost"):
        details.append("- Se esta instalação devia usar outro servidor, corrige LUGEST_DB_HOST no lugest.env desta pasta.")
    else:
        details.append("- Confirma que o servidor MySQL remoto está online e acessível a partir deste posto.")
    details.extend(
        [
            "- Confirma se o ficheiro lugest.env desta instalação é o certo para este ambiente.",
            "- Confirma a porta 3306 e eventuais regras de firewall.",
            "- Confirma utilizador, password e nome da base.",
            "",
            f"Erro técnico: {exc}",
        ]
    )
    return "\n".join(details)


def _mysql_schema_cache_reset():
    global _MYSQL_TABLES_CACHE
    with _MYSQL_SCHEMA_CACHE_LOCK:
        _MYSQL_TABLES_CACHE = None
        _MYSQL_COLUMNS_CACHE.clear()
        _MYSQL_INDEXES_CACHE.clear()


def _mysql_row_lookup(row, *keys):
    if isinstance(row, dict):
        candidates = {}
        for raw_key, value in row.items():
            key_txt = str(raw_key or "").strip()
            if not key_txt:
                continue
            candidates[key_txt] = value
            candidates.setdefault(key_txt.lower(), value)
            candidates.setdefault(key_txt.upper(), value)
        for key in keys:
            key_txt = str(key or "").strip()
            if not key_txt:
                continue
            if key_txt in candidates:
                return candidates[key_txt]
            lower_key = key_txt.lower()
            if lower_key in candidates:
                return candidates[lower_key]
            upper_key = key_txt.upper()
            if upper_key in candidates:
                return candidates[upper_key]
        return None
    if row:
        try:
            return row[0]
        except Exception:
            return None
    return None


def _mysql_existing_tables(cur, force=False):
    global _MYSQL_TABLES_CACHE
    if not force:
        with _MYSQL_SCHEMA_CACHE_LOCK:
            if _MYSQL_TABLES_CACHE is not None:
                return set(_MYSQL_TABLES_CACHE)
    cur.execute(
        "SELECT table_name FROM information_schema.tables WHERE table_schema=%s",
        (MYSQL_DB_NAME,),
    )
    rows = cur.fetchall() or []
    out = set()
    for r in rows:
        name = str(_mysql_row_lookup(r, "table_name", "TABLE_NAME") or "").strip()
        if name:
            out.add(name)
    with _MYSQL_SCHEMA_CACHE_LOCK:
        _MYSQL_TABLES_CACHE = set(out)
    return out


def _mysql_table_columns(cur, table_name, force=False):
    tkey = str(table_name or "").strip().lower()
    if not tkey:
        return set()
    if not force:
        with _MYSQL_SCHEMA_CACHE_LOCK:
            cached = _MYSQL_COLUMNS_CACHE.get(tkey)
        if cached is not None:
            return set(cached)
    cur.execute(
        "SELECT column_name FROM information_schema.columns WHERE table_schema=%s AND table_name=%s",
        (MYSQL_DB_NAME, table_name),
    )
    rows = cur.fetchall() or []
    out = set()
    for r in rows:
        column_name = str(_mysql_row_lookup(r, "column_name", "COLUMN_NAME") or "").strip().lower()
        if column_name:
            out.add(column_name)
    with _MYSQL_SCHEMA_CACHE_LOCK:
        _MYSQL_COLUMNS_CACHE[tkey] = set(out)
    return out


def _mysql_ensure_column(cur, table_name, column_name, column_def):
    cols = _mysql_table_columns(cur, table_name)
    col_key = str(column_name).lower()
    if col_key in cols:
        return
    cur.execute(f"ALTER TABLE `{table_name}` ADD COLUMN `{column_name}` {column_def}")
    with _MYSQL_SCHEMA_CACHE_LOCK:
        tkey = str(table_name or "").strip().lower()
        if tkey:
            _MYSQL_COLUMNS_CACHE.setdefault(tkey, set()).add(col_key)


def _mysql_table_indexes(cur, table_name, force=False):
    tkey = str(table_name or "").strip().lower()
    if not tkey:
        return set()
    if not force:
        with _MYSQL_SCHEMA_CACHE_LOCK:
            cached = _MYSQL_INDEXES_CACHE.get(tkey)
        if cached is not None:
            return set(cached)
    cur.execute(
        "SELECT index_name FROM information_schema.statistics WHERE table_schema=%s AND table_name=%s",
        (MYSQL_DB_NAME, table_name),
    )
    rows = cur.fetchall() or []
    out = set()
    for r in rows:
        index_name = str(_mysql_row_lookup(r, "index_name", "INDEX_NAME") or "").strip().lower()
        if index_name:
            out.add(index_name)
    with _MYSQL_SCHEMA_CACHE_LOCK:
        _MYSQL_INDEXES_CACHE[tkey] = set(out)
    return out


def _mysql_ensure_index(cur, table_name, index_name, cols_sql):
    idx = _mysql_table_indexes(cur, table_name)
    idx_key = str(index_name).lower()
    if idx_key in idx:
        return
    cur.execute(f"ALTER TABLE `{table_name}` ADD INDEX `{index_name}` ({cols_sql})")
    with _MYSQL_SCHEMA_CACHE_LOCK:
        tkey = str(table_name or "").strip().lower()
        if tkey:
            _MYSQL_INDEXES_CACHE.setdefault(tkey, set()).add(idx_key)


def _mysql_runtime_schema_ready(flag):
    return flag in _MYSQL_RUNTIME_SCHEMA_FLAGS


def _mysql_runtime_schema_mark(flag):
    _MYSQL_RUNTIME_SCHEMA_FLAGS.add(flag)


def _mysql_sync_relational_schema(cur, data):
    tables = _mysql_existing_tables(cur, force=True)
    if "notas_encomenda" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS notas_encomenda (
                numero VARCHAR(30) PRIMARY KEY,
                ano INT NULL,
                fornecedor_id VARCHAR(20) NULL,
                contacto VARCHAR(80) NULL,
                data_entrega DATE NULL,
                estado VARCHAR(50) NULL,
                total DECIMAL(12,2) NULL,
                obs TEXT NULL,
                local_descarga VARCHAR(255) NULL,
                meio_transporte VARCHAR(100) NULL,
                oculta BOOLEAN NULL,
                is_draft BOOLEAN NULL,
                data_ultima_entrega DATE NULL,
                guia_ultima VARCHAR(60) NULL,
                fatura_ultima VARCHAR(60) NULL,
                fatura_caminho_ultima VARCHAR(512) NULL,
                data_doc_ultima DATE NULL,
                origem_cotacao VARCHAR(30) NULL,
                ne_geradas TEXT NULL,
                FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id)
            )
            """
        )
    if "notas_encomenda_linhas" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS notas_encomenda_linhas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                ne_numero VARCHAR(30) NOT NULL,
                linha_ordem INT NULL,
                ref_material VARCHAR(20) NULL,
                descricao VARCHAR(255) NULL,
                fornecedor_linha VARCHAR(150) NULL,
                origem VARCHAR(50) NULL,
                qtd DECIMAL(10,2) NULL,
                unid VARCHAR(20) NULL,
                preco DECIMAL(10,2) NULL,
                total DECIMAL(12,2) NULL,
                entregue BOOLEAN NULL,
                qtd_entregue DECIMAL(10,2) NULL,
                lote_fornecedor VARCHAR(100) NULL,
                material VARCHAR(100) NULL,
                espessura VARCHAR(20) NULL,
                comprimento DECIMAL(10,2) NULL,
                largura DECIMAL(10,2) NULL,
                altura DECIMAL(10,2) NULL,
                diametro DECIMAL(10,2) NULL,
                metros DECIMAL(10,2) NULL,
                kg_m DECIMAL(10,4) NULL,
                localizacao VARCHAR(100) NULL,
                peso_unid DECIMAL(10,3) NULL,
                p_compra DECIMAL(10,4) NULL,
                formato VARCHAR(50) NULL,
                material_familia VARCHAR(40) NULL,
                secao_tipo VARCHAR(40) NULL,
                stock_in BOOLEAN NULL,
                guia_entrega VARCHAR(60) NULL,
                fatura_entrega VARCHAR(60) NULL,
                data_doc_entrega DATE NULL,
                data_entrega_real DATE NULL,
                obs_entrega TEXT NULL,
                INDEX idx_ne_linhas_num_ord (ne_numero, linha_ordem),
                FOREIGN KEY (ne_numero) REFERENCES notas_encomenda(numero) ON DELETE CASCADE
            )
            """
        )
    if "notas_encomenda_entregas" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS notas_encomenda_entregas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                ne_numero VARCHAR(30) NOT NULL,
                data_registo DATETIME NULL,
                data_entrega DATE NULL,
                data_documento DATE NULL,
                guia VARCHAR(60) NULL,
                fatura VARCHAR(60) NULL,
                obs TEXT NULL,
                INDEX idx_ne_entregas_num (ne_numero),
                FOREIGN KEY (ne_numero) REFERENCES notas_encomenda(numero) ON DELETE CASCADE
            )
            """
        )
    if "notas_encomenda_linha_entregas" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS notas_encomenda_linha_entregas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                ne_numero VARCHAR(30) NOT NULL,
                linha_ordem INT NOT NULL,
                data_registo DATETIME NULL,
                data_entrega DATE NULL,
                data_documento DATE NULL,
                guia VARCHAR(60) NULL,
                fatura VARCHAR(60) NULL,
                obs TEXT NULL,
                qtd DECIMAL(10,2) NULL,
                lote_fornecedor VARCHAR(100) NULL,
                localizacao VARCHAR(100) NULL,
                entrega_total BOOLEAN NULL,
                stock_ref VARCHAR(30) NULL,
                INDEX idx_ne_linha_entregas_num_ord (ne_numero, linha_ordem),
                FOREIGN KEY (ne_numero) REFERENCES notas_encomenda(numero) ON DELETE CASCADE
            )
            """
        )
    if "notas_encomenda_documentos" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS notas_encomenda_documentos (
                id INT AUTO_INCREMENT PRIMARY KEY,
                ne_numero VARCHAR(30) NOT NULL,
                data_registo DATETIME NULL,
                tipo VARCHAR(40) NULL,
                titulo VARCHAR(150) NULL,
                caminho VARCHAR(512) NULL,
                guia VARCHAR(60) NULL,
                fatura VARCHAR(60) NULL,
                data_entrega DATE NULL,
                data_documento DATE NULL,
                obs TEXT NULL,
                INDEX idx_ne_docs_num (ne_numero),
                FOREIGN KEY (ne_numero) REFERENCES notas_encomenda(numero) ON DELETE CASCADE
            )
            """
        )
    if "expedicoes" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS expedicoes (
                id INT AUTO_INCREMENT PRIMARY KEY,
                numero VARCHAR(30) UNIQUE,
                ano INT NULL,
                tipo VARCHAR(30),
                encomenda_numero VARCHAR(30),
                cliente_codigo VARCHAR(20),
                cliente_nome VARCHAR(150),
                codigo_at VARCHAR(80),
                serie_id VARCHAR(40),
                seq_num INT,
                at_validation_code VARCHAR(40),
                atcud VARCHAR(120),
                emitente_nome VARCHAR(150),
                emitente_nif VARCHAR(20),
                emitente_morada VARCHAR(255),
                destinatario VARCHAR(150),
                dest_nif VARCHAR(20),
                dest_morada VARCHAR(255),
                local_carga VARCHAR(255),
                local_descarga VARCHAR(255),
                data_emissao DATETIME,
                data_transporte DATETIME,
                matricula VARCHAR(30),
                transportador VARCHAR(150),
                estado VARCHAR(50),
                observacoes TEXT,
                created_by VARCHAR(80),
                anulada BOOLEAN,
                anulada_motivo TEXT
            )
            """
        )
        tables.add("expedicoes")
    if "at_series" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS at_series (
                id INT AUTO_INCREMENT PRIMARY KEY,
                doc_type VARCHAR(10) NOT NULL,
                serie_id VARCHAR(40) NOT NULL,
                inicio_sequencia INT NOT NULL DEFAULT 1,
                next_seq INT NOT NULL DEFAULT 1,
                data_inicio_prevista DATE NULL,
                validation_code VARCHAR(40) NULL,
                status VARCHAR(20) NULL,
                last_error TEXT NULL,
                last_sent_payload_hash VARCHAR(64) NULL,
                updated_at DATETIME NULL,
                UNIQUE KEY uq_at_series_doc_serie (doc_type, serie_id)
            )
            """
        )
        tables.add("at_series")
    if "expedicao_linhas" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS expedicao_linhas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                expedicao_numero VARCHAR(30),
                encomenda_numero VARCHAR(30),
                peca_id VARCHAR(30),
                ref_interna VARCHAR(50),
                ref_externa VARCHAR(100),
                descricao VARCHAR(255),
                qtd DECIMAL(10,2),
                unid VARCHAR(20),
                peso DECIMAL(10,3),
                `manual` TINYINT(1) NULL,
                FOREIGN KEY (expedicao_numero) REFERENCES expedicoes(numero) ON DELETE CASCADE
            )
            """
        )
        tables.add("expedicao_linhas")
    if "transportes" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS transportes (
                id INT AUTO_INCREMENT PRIMARY KEY,
                numero VARCHAR(30) UNIQUE,
                ano INT NULL,
                tipo_responsavel VARCHAR(40) NULL,
                estado VARCHAR(40) NULL,
                data_planeada DATE NULL,
                hora_saida VARCHAR(10) NULL,
                viatura VARCHAR(120) NULL,
                matricula VARCHAR(30) NULL,
                motorista VARCHAR(120) NULL,
                telefone_motorista VARCHAR(40) NULL,
                origem VARCHAR(255) NULL,
                transportadora_id VARCHAR(30) NULL,
                transportadora_nome VARCHAR(150) NULL,
                referencia_transporte VARCHAR(80) NULL,
                custo_previsto DECIMAL(12,2) NULL,
                paletes_total_manual DECIMAL(10,2) NULL,
                peso_total_manual_kg DECIMAL(12,2) NULL,
                volume_total_manual_m3 DECIMAL(12,3) NULL,
                pedido_transporte_estado VARCHAR(40) NULL,
                pedido_transporte_ref VARCHAR(80) NULL,
                pedido_transporte_at DATETIME NULL,
                pedido_transporte_by VARCHAR(80) NULL,
                pedido_transporte_obs TEXT NULL,
                pedido_resposta_obs TEXT NULL,
                pedido_confirmado_at DATETIME NULL,
                pedido_confirmado_by VARCHAR(80) NULL,
                pedido_recusado_at DATETIME NULL,
                pedido_recusado_by VARCHAR(80) NULL,
                observacoes TEXT NULL,
                created_by VARCHAR(80) NULL,
                created_at DATETIME NULL,
                updated_at DATETIME NULL
            )
            """
        )
        tables.add("transportes")
    if "transportes_paragens" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS transportes_paragens (
                id INT AUTO_INCREMENT PRIMARY KEY,
                transporte_numero VARCHAR(30) NOT NULL,
                ordem INT NULL,
                encomenda_numero VARCHAR(30) NULL,
                expedicao_numero VARCHAR(30) NULL,
                cliente_codigo VARCHAR(20) NULL,
                cliente_nome VARCHAR(150) NULL,
                local_descarga VARCHAR(255) NULL,
                contacto VARCHAR(120) NULL,
                telefone VARCHAR(40) NULL,
                data_planeada DATETIME NULL,
                paletes DECIMAL(10,2) NULL,
                peso_bruto_kg DECIMAL(12,2) NULL,
                volume_m3 DECIMAL(12,3) NULL,
                preco_transporte DECIMAL(12,2) NULL,
                custo_transporte DECIMAL(12,2) NULL,
                transportadora_id VARCHAR(30) NULL,
                transportadora_nome VARCHAR(150) NULL,
                referencia_transporte VARCHAR(80) NULL,
                check_carga_ok TINYINT(1) NULL,
                check_docs_ok TINYINT(1) NULL,
                check_paletes_ok TINYINT(1) NULL,
                pod_estado VARCHAR(40) NULL,
                pod_recebido_nome VARCHAR(120) NULL,
                pod_recebido_at DATETIME NULL,
                pod_obs TEXT NULL,
                estado VARCHAR(40) NULL,
                observacoes TEXT NULL,
                INDEX idx_transportes_paragens_num_ord (transporte_numero, ordem),
                INDEX idx_transportes_paragens_enc (encomenda_numero),
                FOREIGN KEY (transporte_numero) REFERENCES transportes(numero) ON DELETE CASCADE
            )
            """
        )
        tables.add("transportes_paragens")
    if "transportes_tarifarios" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS transportes_tarifarios (
                id INT AUTO_INCREMENT PRIMARY KEY,
                transportadora_id VARCHAR(30) NULL,
                transportadora_nome VARCHAR(150) NULL,
                zona VARCHAR(120) NOT NULL,
                valor_base DECIMAL(12,2) NULL,
                valor_por_palete DECIMAL(12,2) NULL,
                valor_por_kg DECIMAL(12,4) NULL,
                valor_por_m3 DECIMAL(12,2) NULL,
                custo_minimo DECIMAL(12,2) NULL,
                ativo TINYINT(1) NULL,
                observacoes TEXT NULL,
                INDEX idx_transportes_tarifarios_carrier_zone (transportadora_id, zona)
            )
            """
        )
        tables.add("transportes_tarifarios")
    if "faturacao_registos" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS faturacao_registos (
                numero VARCHAR(30) PRIMARY KEY,
                ano INT NULL,
                origem VARCHAR(30) NULL,
                orcamento_numero VARCHAR(30) NULL,
                encomenda_numero VARCHAR(30) NULL,
                cliente_codigo VARCHAR(20) NULL,
                cliente_nome VARCHAR(150) NULL,
                data_venda DATE NULL,
                data_vencimento DATE NULL,
                valor_venda_manual DECIMAL(12,2) NULL,
                estado_pagamento_manual VARCHAR(30) NULL,
                obs TEXT NULL,
                created_at DATETIME NULL,
                updated_at DATETIME NULL
            )
            """
        )
        tables.add("faturacao_registos")
    if "faturacao_faturas" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS faturacao_faturas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                registo_numero VARCHAR(30) NOT NULL,
                documento_id VARCHAR(60) NULL,
                doc_type VARCHAR(10) NULL,
                numero_fatura VARCHAR(60) NULL,
                serie VARCHAR(40) NULL,
                serie_id VARCHAR(40) NULL,
                seq_num INT NULL,
                at_validation_code VARCHAR(40) NULL,
                atcud VARCHAR(80) NULL,
                guia_numero VARCHAR(30) NULL,
                data_emissao DATE NULL,
                data_vencimento DATE NULL,
                moeda VARCHAR(10) NULL,
                iva_perc DECIMAL(6,2) NULL,
                subtotal DECIMAL(12,2) NULL,
                valor_iva DECIMAL(12,2) NULL,
                valor_total DECIMAL(12,2) NULL,
                caminho VARCHAR(512) NULL,
                obs TEXT NULL,
                estado VARCHAR(30) NULL,
                anulada TINYINT(1) NULL DEFAULT 0,
                anulada_motivo TEXT NULL,
                anulada_at DATETIME NULL,
                legal_invoice_no VARCHAR(80) NULL,
                system_entry_date DATETIME NULL,
                source_id VARCHAR(80) NULL,
                source_billing VARCHAR(1) NULL,
                status_source_id VARCHAR(80) NULL,
                hash VARCHAR(512) NULL,
                hash_control VARCHAR(80) NULL,
                previous_hash VARCHAR(512) NULL,
                document_snapshot_json LONGTEXT NULL,
                communication_status VARCHAR(30) NULL,
                communication_filename VARCHAR(255) NULL,
                communication_error TEXT NULL,
                communicated_at DATETIME NULL,
                communication_batch_id VARCHAR(80) NULL,
                created_at DATETIME NULL,
                INDEX idx_fat_faturas_reg (registo_numero),
                INDEX idx_fat_faturas_doc (documento_id),
                INDEX idx_fat_faturas_emissao (data_emissao),
                INDEX idx_fat_faturas_serie_seq (serie_id, seq_num),
                INDEX idx_fat_faturas_legal (legal_invoice_no),
                INDEX idx_fat_faturas_comm (communication_status),
                FOREIGN KEY (registo_numero) REFERENCES faturacao_registos(numero) ON DELETE CASCADE
            )
            """
        )
        tables.add("faturacao_faturas")
    if "faturacao_pagamentos" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS faturacao_pagamentos (
                id INT AUTO_INCREMENT PRIMARY KEY,
                registo_numero VARCHAR(30) NOT NULL,
                pagamento_id VARCHAR(60) NULL,
                fatura_documento_id VARCHAR(60) NULL,
                data_pagamento DATE NULL,
                valor DECIMAL(12,2) NULL,
                metodo VARCHAR(40) NULL,
                referencia VARCHAR(120) NULL,
                titulo_comprovativo VARCHAR(150) NULL,
                caminho_comprovativo VARCHAR(512) NULL,
                obs TEXT NULL,
                created_at DATETIME NULL,
                INDEX idx_fat_pag_reg (registo_numero),
                INDEX idx_fat_pag_fatura (fatura_documento_id),
                INDEX idx_fat_pag_data (data_pagamento),
                FOREIGN KEY (registo_numero) REFERENCES faturacao_registos(numero) ON DELETE CASCADE
            )
            """
        )
        tables.add("faturacao_pagamentos")
    if "ne_linhas_historico" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS ne_linhas_historico (
                id INT AUTO_INCREMENT PRIMARY KEY,
                created_at DATETIME NOT NULL,
                evento VARCHAR(30) NOT NULL,
                origem_menu VARCHAR(80) NULL,
                utilizador VARCHAR(80) NULL,
                guia_numero VARCHAR(30) NULL,
                produto_codigo VARCHAR(20) NULL,
                descricao VARCHAR(255) NULL,
                qtd DECIMAL(10,2) NULL,
                unid VARCHAR(20) NULL,
                destinatario VARCHAR(150) NULL,
                observacoes TEXT NULL,
                payload_json LONGTEXT NULL,
                INDEX idx_ne_linhas_hist_created_at (created_at),
                INDEX idx_ne_linhas_hist_evento (evento),
                INDEX idx_ne_linhas_hist_guia (guia_numero)
            )
            """
        )
    if "encomenda_espessuras" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS encomenda_espessuras (
                id INT AUTO_INCREMENT PRIMARY KEY,
                encomenda_numero VARCHAR(30) NOT NULL,
                material VARCHAR(100) NOT NULL,
                espessura VARCHAR(20) NOT NULL,
                tempo_min DECIMAL(10,2) NULL,
                tempos_operacao_json LONGTEXT NULL,
                maquinas_operacao_json LONGTEXT NULL,
                estado VARCHAR(50) NULL,
                inicio_producao DATETIME NULL,
                fim_producao DATETIME NULL,
                tempo_producao_min DECIMAL(10,2) NULL,
                lote_baixa VARCHAR(100) NULL,
                UNIQUE KEY uq_enc_esp (encomenda_numero, material, espessura),
                INDEX idx_enc_esp_num (encomenda_numero),
                FOREIGN KEY (encomenda_numero) REFERENCES encomendas(numero) ON DELETE CASCADE
            )
            """
        )
    if "encomenda_reservas" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS encomenda_reservas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                encomenda_numero VARCHAR(30) NOT NULL,
                material_id VARCHAR(30) NULL,
                material VARCHAR(100) NOT NULL,
                espessura VARCHAR(20) NOT NULL,
                quantidade DECIMAL(10,2) NOT NULL,
                created_at DATETIME NULL,
                INDEX idx_enc_res_num (encomenda_numero),
                INDEX idx_enc_res_mat_esp (material, espessura),
                FOREIGN KEY (encomenda_numero) REFERENCES encomendas(numero) ON DELETE CASCADE
            )
            """
        )
    if "plano" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS plano (
                id INT AUTO_INCREMENT PRIMARY KEY,
                bloco_id VARCHAR(60) NOT NULL,
                encomenda_numero VARCHAR(30) NOT NULL,
                ano INT NULL,
                material VARCHAR(100) NULL,
                espessura VARCHAR(20) NULL,
                operacao VARCHAR(80) NULL,
                posto VARCHAR(80) NULL,
                data_planeada DATE NOT NULL,
                inicio VARCHAR(8) NOT NULL,
                duracao_min DECIMAL(10,2) NOT NULL,
                color VARCHAR(20) NULL,
                chapa VARCHAR(120) NULL,
                UNIQUE KEY uq_plano_bloco (bloco_id),
                INDEX idx_plano_data_inicio (data_planeada, inicio),
                INDEX idx_plano_enc (encomenda_numero)
            )
            """
        )
    if "plano_hist" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS plano_hist (
                id INT AUTO_INCREMENT PRIMARY KEY,
                bloco_id VARCHAR(60) NOT NULL,
                encomenda_numero VARCHAR(30) NOT NULL,
                ano INT NULL,
                material VARCHAR(100) NULL,
                espessura VARCHAR(20) NULL,
                operacao VARCHAR(80) NULL,
                posto VARCHAR(80) NULL,
                data_planeada DATE NOT NULL,
                inicio VARCHAR(8) NOT NULL,
                duracao_min DECIMAL(10,2) NOT NULL,
                color VARCHAR(20) NULL,
                chapa VARCHAR(120) NULL,
                movido_em DATETIME NULL,
                estado_final VARCHAR(50) NULL,
                tempo_planeado_min DECIMAL(10,2) NULL,
                tempo_real_min DECIMAL(10,2) NULL,
                INDEX idx_plano_hist_enc (encomenda_numero),
                INDEX idx_plano_hist_data (data_planeada, inicio)
            )
            """
        )
    if "orc_referencias_historico" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS orc_referencias_historico (
                id INT AUTO_INCREMENT PRIMARY KEY,
                ref_externa VARCHAR(100) NOT NULL,
                ref_interna VARCHAR(50) NULL,
                descricao TEXT NULL,
                material VARCHAR(100) NULL,
                espessura VARCHAR(20) NULL,
                preco_unit DECIMAL(10,2) NULL,
                operacao VARCHAR(150) NULL,
                desenho_path VARCHAR(512) NULL,
                updated_at DATETIME NOT NULL,
                UNIQUE KEY uq_orc_ref_externa (ref_externa),
                INDEX idx_orc_ref_interna (ref_interna),
                INDEX idx_orc_ref_updated_at (updated_at)
            )
            """
        )
    if "peca_operacoes_execucao" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS peca_operacoes_execucao (
                id INT AUTO_INCREMENT PRIMARY KEY,
                encomenda_numero VARCHAR(30) NOT NULL,
                peca_id VARCHAR(30) NOT NULL,
                operacao VARCHAR(80) NOT NULL,
                estado VARCHAR(20) NOT NULL DEFAULT 'Livre',
                operador_atual VARCHAR(120) NULL,
                inicio_ts DATETIME NULL,
                fim_ts DATETIME NULL,
                ok_qty DECIMAL(10,2) NULL,
                nok_qty DECIMAL(10,2) NULL,
                qual_qty DECIMAL(10,2) NULL,
                updated_at DATETIME NULL,
                UNIQUE KEY uq_peca_operacao (peca_id, operacao),
                INDEX idx_poe_enc (encomenda_numero),
                INDEX idx_poe_estado (estado),
                INDEX idx_poe_operador (operador_atual)
            )
            """
        )
    if "users" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(50) NOT NULL,
                password VARCHAR(255) NOT NULL,
                role VARCHAR(50) NOT NULL,
                UNIQUE KEY uq_users_username (username)
            )
            """
        )
    if "operadores" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS operadores (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nome VARCHAR(120) NOT NULL,
                UNIQUE KEY uq_operadores_nome (nome)
            )
            """
        )
    if "orcamentistas" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS orcamentistas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nome VARCHAR(120) NOT NULL,
                UNIQUE KEY uq_orcamentistas_nome (nome)
            )
            """
        )
    if "op_eventos" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS op_eventos (
                id INT AUTO_INCREMENT PRIMARY KEY,
                created_at DATETIME NOT NULL,
                evento VARCHAR(30) NOT NULL,
                encomenda_numero VARCHAR(30) NULL,
                peca_id VARCHAR(30) NULL,
                ref_interna VARCHAR(60) NULL,
                material VARCHAR(100) NULL,
                espessura VARCHAR(20) NULL,
                operacao VARCHAR(80) NULL,
                operador VARCHAR(80) NULL,
                qtd_ok DECIMAL(10,2) NULL,
                qtd_nok DECIMAL(10,2) NULL,
                info TEXT NULL,
                INDEX idx_op_eventos_created_at (created_at),
                INDEX idx_op_eventos_evento (evento),
                INDEX idx_op_eventos_enc (encomenda_numero),
                INDEX idx_op_eventos_peca (peca_id),
                INDEX idx_op_eventos_operador (operador)
            )
            """
        )
    if "op_paragens" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS op_paragens (
                id INT AUTO_INCREMENT PRIMARY KEY,
                created_at DATETIME NOT NULL,
                fechada_at DATETIME NULL,
                encomenda_numero VARCHAR(30) NULL,
                peca_id VARCHAR(30) NULL,
                ref_interna VARCHAR(60) NULL,
                material VARCHAR(100) NULL,
                espessura VARCHAR(20) NULL,
                operador VARCHAR(80) NULL,
                origem VARCHAR(20) NULL,
                estado VARCHAR(20) NULL,
                causa VARCHAR(120) NULL,
                detalhe TEXT NULL,
                grupo_id VARCHAR(80) NULL,
                duracao_min DECIMAL(10,2) NULL,
                INDEX idx_op_paragens_created_at (created_at),
                INDEX idx_op_paragens_causa (causa),
                INDEX idx_op_paragens_enc (encomenda_numero),
                INDEX idx_op_paragens_peca (peca_id)
            )
            """
        )
    if "op_paragens" in tables:
        _mysql_ensure_column(cur, "op_paragens", "fechada_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "op_paragens", "origem", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "op_paragens", "estado", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "op_paragens", "grupo_id", "VARCHAR(80) NULL")
    if "produtos_mov" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS produtos_mov (
                id INT AUTO_INCREMENT PRIMARY KEY,
                data DATETIME NULL,
                tipo VARCHAR(40) NULL,
                operador VARCHAR(120) NULL,
                codigo VARCHAR(20) NULL,
                descricao VARCHAR(255) NULL,
                qtd DECIMAL(10,2) NULL,
                antes DECIMAL(10,2) NULL,
                depois DECIMAL(10,2) NULL,
                obs TEXT NULL,
                origem VARCHAR(80) NULL,
                ref_doc VARCHAR(50) NULL,
                INDEX idx_prod_mov_data (data),
                INDEX idx_prod_mov_operador (operador),
                INDEX idx_prod_mov_codigo (codigo)
            )
            """
        )
    if "categories" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS categories (
                id VARCHAR(80) PRIMARY KEY,
                nome VARCHAR(120) NOT NULL,
                icon VARCHAR(30) NULL,
                badge VARCHAR(60) NULL,
                updated_at DATETIME NULL
            )
            """
        )
    if "subcategories" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS subcategories (
                id VARCHAR(80) PRIMARY KEY,
                category_id VARCHAR(80) NOT NULL,
                nome VARCHAR(120) NOT NULL,
                updated_at DATETIME NULL,
                INDEX idx_subcategories_category (category_id)
            )
            """
        )
    if "product_types" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS product_types (
                id VARCHAR(80) PRIMARY KEY,
                subcategory_id VARCHAR(80) NOT NULL,
                nome VARCHAR(120) NOT NULL,
                updated_at DATETIME NULL,
                INDEX idx_product_types_subcategory (subcategory_id)
            )
            """
        )
    if "product_documents" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS product_documents (
                id VARCHAR(80) PRIMARY KEY,
                produto_codigo VARCHAR(20) NOT NULL,
                tipo VARCHAR(40) NULL,
                titulo VARCHAR(180) NULL,
                caminho VARCHAR(512) NULL,
                versao VARCHAR(30) NULL,
                created_at DATETIME NULL,
                updated_at DATETIME NULL,
                INDEX idx_product_documents_codigo (produto_codigo)
            )
            """
        )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS quality_nonconformities (
            id VARCHAR(30) PRIMARY KEY,
            origem VARCHAR(120) NULL,
            referencia VARCHAR(120) NULL,
            entidade_tipo VARCHAR(60) NULL,
            entidade_id VARCHAR(120) NULL,
            entidade_label VARCHAR(255) NULL,
            tipo VARCHAR(60) NULL,
            gravidade VARCHAR(40) NULL,
            estado VARCHAR(40) NULL,
            responsavel VARCHAR(120) NULL,
            prazo DATE NULL,
            descricao TEXT NULL,
            causa TEXT NULL,
            acao TEXT NULL,
            eficacia TEXT NULL,
            fornecedor_id VARCHAR(30) NULL,
            fornecedor_nome VARCHAR(150) NULL,
            material_id VARCHAR(30) NULL,
            lote_fornecedor VARCHAR(100) NULL,
            ne_numero VARCHAR(30) NULL,
            guia VARCHAR(60) NULL,
            fatura VARCHAR(60) NULL,
            decisao VARCHAR(255) NULL,
            movement_id VARCHAR(255) NULL,
            qtd_recebida DECIMAL(10,2) NULL,
            qtd_aprovada DECIMAL(10,2) NULL,
            qtd_rejeitada DECIMAL(10,2) NULL,
            qtd_pendente DECIMAL(10,2) NULL,
            created_at DATETIME NULL,
            updated_at DATETIME NULL,
            created_by VARCHAR(120) NULL,
            updated_by VARCHAR(120) NULL,
            closed_at DATETIME NULL,
            closed_by VARCHAR(120) NULL,
            INDEX idx_quality_nc_estado (estado),
            INDEX idx_quality_nc_entidade (entidade_tipo, entidade_id),
            INDEX idx_quality_nc_referencia (referencia)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS quality_documents (
            id VARCHAR(30) PRIMARY KEY,
            titulo VARCHAR(180) NOT NULL,
            tipo VARCHAR(80) NULL,
            entidade VARCHAR(80) NULL,
            referencia VARCHAR(120) NULL,
            entidade_tipo VARCHAR(60) NULL,
            entidade_id VARCHAR(120) NULL,
            versao VARCHAR(30) NULL,
            estado VARCHAR(40) NULL,
            responsavel VARCHAR(120) NULL,
            caminho VARCHAR(512) NULL,
            obs TEXT NULL,
            created_at DATETIME NULL,
            updated_at DATETIME NULL,
            created_by VARCHAR(120) NULL,
            updated_by VARCHAR(120) NULL,
            INDEX idx_quality_doc_entidade (entidade_tipo, entidade_id),
            INDEX idx_quality_doc_tipo (tipo),
            INDEX idx_quality_doc_estado (estado)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS quality_audit_log (
            id VARCHAR(80) PRIMARY KEY,
            created_at DATETIME NULL,
            user_name VARCHAR(120) NULL,
            action VARCHAR(120) NULL,
            entity_type VARCHAR(80) NULL,
            entity_id VARCHAR(120) NULL,
            summary TEXT NULL,
            before_json LONGTEXT NULL,
            after_json LONGTEXT NULL,
            INDEX idx_quality_audit_created (created_at),
            INDEX idx_quality_audit_entity (entity_type, entity_id)
        )
        """
    )
    if "conjuntos_modelo" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS conjuntos_modelo (
                codigo VARCHAR(40) PRIMARY KEY,
                descricao VARCHAR(150) NOT NULL,
                notas TEXT NULL,
                ativo BOOLEAN NULL,
                template BOOLEAN NULL,
                origem VARCHAR(80) NULL,
                created_at DATETIME NULL,
                updated_at DATETIME NULL
            )
            """
        )
    if "conjuntos" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS conjuntos (
                codigo VARCHAR(40) PRIMARY KEY,
                descricao VARCHAR(150) NOT NULL,
                notas TEXT NULL,
                ativo BOOLEAN NULL,
                template BOOLEAN NULL,
                origem VARCHAR(80) NULL,
                margem_perc DECIMAL(10,2) NULL,
                total_custo DECIMAL(12,2) NULL,
                total_final DECIMAL(12,2) NULL,
                created_at DATETIME NULL,
                updated_at DATETIME NULL
            )
            """
        )
    if "conjuntos_itens" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS conjuntos_itens (
                id INT AUTO_INCREMENT PRIMARY KEY,
                conjunto_codigo VARCHAR(40) NOT NULL,
                linha_ordem INT NULL,
                tipo_item VARCHAR(30) NULL,
                ref_externa VARCHAR(100) NULL,
                descricao TEXT NULL,
                material VARCHAR(100) NULL,
                espessura VARCHAR(20) NULL,
                operacao VARCHAR(150) NULL,
                produto_codigo VARCHAR(20) NULL,
                produto_unid VARCHAR(20) NULL,
                qtd DECIMAL(10,2) NULL,
                tempo_peca_min DECIMAL(10,2) NULL,
                preco_unit DECIMAL(10,4) NULL,
                desenho_path VARCHAR(512) NULL,
                meta_json LONGTEXT NULL,
                INDEX idx_conjuntos_codigo_ord (conjunto_codigo, linha_ordem),
                FOREIGN KEY (conjunto_codigo) REFERENCES conjuntos(codigo) ON DELETE CASCADE
            )
            """
        )
    if "conjuntos_modelo_itens" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS conjuntos_modelo_itens (
                id INT AUTO_INCREMENT PRIMARY KEY,
                conjunto_codigo VARCHAR(40) NOT NULL,
                linha_ordem INT NULL,
                tipo_item VARCHAR(30) NULL,
                ref_externa VARCHAR(100) NULL,
                descricao TEXT NULL,
                material VARCHAR(100) NULL,
                espessura VARCHAR(20) NULL,
                operacao VARCHAR(150) NULL,
                produto_codigo VARCHAR(20) NULL,
                produto_unid VARCHAR(20) NULL,
                qtd DECIMAL(10,2) NULL,
                tempo_peca_min DECIMAL(10,2) NULL,
                preco_unit DECIMAL(10,4) NULL,
                desenho_path VARCHAR(512) NULL,
                meta_json LONGTEXT NULL,
                INDEX idx_conjuntos_itens_codigo_ord (conjunto_codigo, linha_ordem),
                FOREIGN KEY (conjunto_codigo) REFERENCES conjuntos_modelo(codigo) ON DELETE CASCADE
            )
            """
        )
    if "encomenda_montagem_itens" not in tables:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS encomenda_montagem_itens (
                id INT AUTO_INCREMENT PRIMARY KEY,
                encomenda_numero VARCHAR(30) NOT NULL,
                linha_ordem INT NULL,
                tipo_item VARCHAR(30) NULL,
                descricao TEXT NULL,
                produto_codigo VARCHAR(20) NULL,
                produto_unid VARCHAR(20) NULL,
                qtd_planeada DECIMAL(10,2) NULL,
                qtd_consumida DECIMAL(10,2) NULL,
                preco_unit DECIMAL(10,4) NULL,
                conjunto_codigo VARCHAR(40) NULL,
                conjunto_nome VARCHAR(150) NULL,
                grupo_uuid VARCHAR(60) NULL,
                estado VARCHAR(30) NULL,
                obs TEXT NULL,
                created_at DATETIME NULL,
                consumed_at DATETIME NULL,
                consumed_by VARCHAR(120) NULL,
                INDEX idx_enc_montagem_num_ord (encomenda_numero, linha_ordem),
                INDEX idx_enc_montagem_estado (estado),
                FOREIGN KEY (encomenda_numero) REFERENCES encomendas(numero) ON DELETE CASCADE
            )
            """
        )
    tables = _mysql_existing_tables(cur, force=True)
    if "fornecedores" in tables:
        _mysql_ensure_column(cur, "fornecedores", "nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "fornecedores", "nif", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "fornecedores", "morada", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "fornecedores", "contacto", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "fornecedores", "email", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "fornecedores", "codigo_postal", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "fornecedores", "localidade", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "fornecedores", "pais", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "fornecedores", "cond_pagamento", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "fornecedores", "prazo_entrega_dias", "INT NULL")
        _mysql_ensure_column(cur, "fornecedores", "website", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "fornecedores", "obs", "TEXT NULL")
    if "materiais" in tables:
        _mysql_ensure_column(cur, "materiais", "preco_unid", "DECIMAL(12,4) NULL")
        _mysql_ensure_column(cur, "materiais", "origem_lote", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "materiais", "origem_encomenda", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "materiais", "material_familia", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "materiais", "altura", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "materiais", "diametro", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "materiais", "kg_m", "DECIMAL(10,4) NULL")
        _mysql_ensure_column(cur, "materiais", "secao_tipo", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "materiais", "logistic_status", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "materiais", "quality_status", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "materiais", "quality_blocked", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "materiais", "quality_pending_qty", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "materiais", "quality_received_qty", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "materiais", "quality_approved_qty", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "materiais", "quality_return_document_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_status", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_defect", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_decision", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_by", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_note_number", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_supplier_id", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_supplier_name", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_guia", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "materiais", "inspection_fatura", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "materiais", "quality_nc_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "materiais", "supplier_claim_id", "VARCHAR(30) NULL")
    if "produtos" in tables:
        _mysql_ensure_column(cur, "produtos", "category_id", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "produtos", "subcategory_id", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "produtos", "type_id", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "produtos", "logistic_status", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "produtos", "quality_status", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "produtos", "quality_blocked", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "produtos", "quality_pending_qty", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "produtos", "quality_received_qty", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "produtos", "quality_approved_qty", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "produtos", "quality_return_document_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "produtos", "inspection_defect", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "produtos", "inspection_decision", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "produtos", "inspection_note_number", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "produtos", "quality_nc_id", "VARCHAR(30) NULL")
    if "notas_encomenda" in tables:
        _mysql_ensure_column(cur, "notas_encomenda", "ano", "INT NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "fornecedor_id", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "contacto", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "data_entrega", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "estado", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "total", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "obs", "TEXT NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "local_descarga", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "meio_transporte", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "oculta", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "is_draft", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "data_ultima_entrega", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "guia_ultima", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "fatura_ultima", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "fatura_caminho_ultima", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "data_doc_ultima", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "origem_cotacao", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "notas_encomenda", "ne_geradas", "TEXT NULL")
    if "notas_encomenda_linhas" in tables:
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "linha_ordem", "INT NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "ref_material", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "descricao", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "fornecedor_linha", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "origem", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "qtd", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "unid", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "preco", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "desconto", "DECIMAL(6,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "iva", "DECIMAL(6,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "total", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "entregue", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "qtd_entregue", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "lote_fornecedor", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "material", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "espessura", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "comprimento", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "largura", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "altura", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "diametro", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "metros", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "kg_m", "DECIMAL(10,4) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "localizacao", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "peso_unid", "DECIMAL(10,3) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "p_compra", "DECIMAL(10,4) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "formato", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "material_familia", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "secao_tipo", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "stock_in", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "guia_entrega", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "fatura_entrega", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "data_doc_entrega", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "data_entrega_real", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "obs_entrega", "TEXT NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "logistic_status", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "inspection_status", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "inspection_defect", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "inspection_decision", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "quality_status", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linhas", "quality_nc_id", "VARCHAR(30) NULL")
    if "notas_encomenda_entregas" in tables:
        _mysql_ensure_column(cur, "notas_encomenda_entregas", "ne_numero", "VARCHAR(30) NOT NULL")
        _mysql_ensure_column(cur, "notas_encomenda_entregas", "data_registo", "DATETIME NULL")
        _mysql_ensure_column(cur, "notas_encomenda_entregas", "data_entrega", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda_entregas", "data_documento", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda_entregas", "guia", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_entregas", "fatura", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_entregas", "obs", "TEXT NULL")
    if "notas_encomenda_linha_entregas" in tables:
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "ne_numero", "VARCHAR(30) NOT NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "linha_ordem", "INT NOT NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "data_registo", "DATETIME NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "data_entrega", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "data_documento", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "guia", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "fatura", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "obs", "TEXT NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "qtd", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "lote_fornecedor", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "localizacao", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "entrega_total", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "stock_ref", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "logistic_status", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "inspection_status", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "inspection_defect", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "inspection_decision", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "quality_status", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_linha_entregas", "quality_nc_id", "VARCHAR(30) NULL")
    if "notas_encomenda_documentos" in tables:
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "ne_numero", "VARCHAR(30) NOT NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "data_registo", "DATETIME NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "tipo", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "titulo", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "caminho", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "guia", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "fatura", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "data_entrega", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "data_documento", "DATE NULL")
        _mysql_ensure_column(cur, "notas_encomenda_documentos", "obs", "TEXT NULL")
    if "expedicoes" in tables:
        _mysql_ensure_column(cur, "expedicoes", "ano", "INT NULL")
        _mysql_ensure_column(cur, "expedicoes", "tipo", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "expedicoes", "encomenda_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "expedicoes", "cliente_codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "expedicoes", "cliente_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "expedicoes", "codigo_at", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "expedicoes", "serie_id", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "expedicoes", "seq_num", "INT NULL")
        _mysql_ensure_column(cur, "expedicoes", "at_validation_code", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "expedicoes", "atcud", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "expedicoes", "emitente_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "expedicoes", "emitente_nif", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "expedicoes", "emitente_morada", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "expedicoes", "destinatario", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "expedicoes", "dest_nif", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "expedicoes", "dest_morada", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "expedicoes", "local_carga", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "expedicoes", "local_descarga", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "expedicoes", "data_emissao", "DATETIME NULL")
        _mysql_ensure_column(cur, "expedicoes", "data_transporte", "DATETIME NULL")
        _mysql_ensure_column(cur, "expedicoes", "matricula", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "expedicoes", "transportador", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "expedicoes", "estado", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "expedicoes", "observacoes", "TEXT NULL")
        _mysql_ensure_column(cur, "expedicoes", "created_by", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "expedicoes", "anulada", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "expedicoes", "anulada_motivo", "TEXT NULL")
    if "at_series" in tables:
        _mysql_ensure_column(cur, "at_series", "doc_type", "VARCHAR(10) NOT NULL")
        _mysql_ensure_column(cur, "at_series", "serie_id", "VARCHAR(40) NOT NULL")
        _mysql_ensure_column(cur, "at_series", "inicio_sequencia", "INT NOT NULL DEFAULT 1")
        _mysql_ensure_column(cur, "at_series", "next_seq", "INT NOT NULL DEFAULT 1")
        _mysql_ensure_column(cur, "at_series", "data_inicio_prevista", "DATE NULL")
        _mysql_ensure_column(cur, "at_series", "validation_code", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "at_series", "status", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "at_series", "last_error", "TEXT NULL")
        _mysql_ensure_column(cur, "at_series", "last_sent_payload_hash", "VARCHAR(64) NULL")
        _mysql_ensure_column(cur, "at_series", "updated_at", "DATETIME NULL")
    if "expedicao_linhas" in tables:
        _mysql_ensure_column(cur, "expedicao_linhas", "encomenda_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "expedicao_linhas", "peca_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "expedicao_linhas", "ref_interna", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "expedicao_linhas", "ref_externa", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "expedicao_linhas", "descricao", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "expedicao_linhas", "qtd", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "expedicao_linhas", "unid", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "expedicao_linhas", "peso", "DECIMAL(10,3) NULL")
        _mysql_ensure_column(cur, "expedicao_linhas", "manual", "BOOLEAN NULL")
    if "transportes" in tables:
        _mysql_ensure_column(cur, "transportes", "ano", "INT NULL")
        _mysql_ensure_column(cur, "transportes", "tipo_responsavel", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "transportes", "estado", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "transportes", "data_planeada", "DATE NULL")
        _mysql_ensure_column(cur, "transportes", "hora_saida", "VARCHAR(10) NULL")
        _mysql_ensure_column(cur, "transportes", "viatura", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "transportes", "matricula", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "transportes", "motorista", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "transportes", "telefone_motorista", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "transportes", "origem", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "transportes", "observacoes", "TEXT NULL")
        _mysql_ensure_column(cur, "transportes", "created_by", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "transportes", "created_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "transportes", "updated_at", "DATETIME NULL")
        _mysql_ensure_index(cur, "transportes", "idx_transportes_data_estado", "data_planeada, estado")
    if "transportes_paragens" in tables:
        _mysql_ensure_column(cur, "transportes_paragens", "ordem", "INT NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "encomenda_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "expedicao_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "cliente_codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "cliente_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "zona_transporte", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "local_descarga", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "contacto", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "telefone", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "data_planeada", "DATETIME NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "estado", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "observacoes", "TEXT NULL")
    if "transportes_tarifarios" in tables:
        _mysql_ensure_column(cur, "transportes_tarifarios", "transportadora_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "transportadora_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "zona", "VARCHAR(120) NOT NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "valor_base", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "valor_por_palete", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "valor_por_kg", "DECIMAL(12,4) NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "valor_por_m3", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "custo_minimo", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "ativo", "TINYINT(1) NULL")
        _mysql_ensure_column(cur, "transportes_tarifarios", "observacoes", "TEXT NULL")
        _mysql_ensure_index(cur, "transportes_tarifarios", "idx_transportes_tarifarios_carrier_zone", "transportadora_id, zona")
    if "encomenda_espessuras" in tables:
        _mysql_ensure_column(cur, "encomenda_espessuras", "encomenda_numero", "VARCHAR(30) NOT NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "material", "VARCHAR(100) NOT NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "espessura", "VARCHAR(20) NOT NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "tempo_min", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "tempos_operacao_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "maquinas_operacao_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "estado", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "inicio_producao", "DATETIME NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "fim_producao", "DATETIME NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "tempo_producao_min", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "encomenda_espessuras", "lote_baixa", "VARCHAR(100) NULL")
    if "encomenda_reservas" in tables:
        _mysql_ensure_column(cur, "encomenda_reservas", "encomenda_numero", "VARCHAR(30) NOT NULL")
        _mysql_ensure_column(cur, "encomenda_reservas", "material_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "encomenda_reservas", "material", "VARCHAR(100) NOT NULL")
        _mysql_ensure_column(cur, "encomenda_reservas", "espessura", "VARCHAR(20) NOT NULL")
        _mysql_ensure_column(cur, "encomenda_reservas", "quantidade", "DECIMAL(10,2) NOT NULL")
        _mysql_ensure_column(cur, "encomenda_reservas", "created_at", "DATETIME NULL")
    if "plano" in tables:
        _mysql_ensure_column(cur, "plano", "ano", "INT NULL")
        _mysql_ensure_column(cur, "plano", "bloco_id", "VARCHAR(60) NOT NULL")
        _mysql_ensure_column(cur, "plano", "encomenda_numero", "VARCHAR(30) NOT NULL")
        _mysql_ensure_column(cur, "plano", "material", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "plano", "espessura", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "plano", "operacao", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "plano", "posto", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "plano", "data_planeada", "DATE NOT NULL")
        _mysql_ensure_column(cur, "plano", "inicio", "VARCHAR(8) NOT NULL")
        _mysql_ensure_column(cur, "plano", "duracao_min", "DECIMAL(10,2) NOT NULL")
        _mysql_ensure_column(cur, "plano", "color", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "plano", "chapa", "VARCHAR(120) NULL")
    if "plano_hist" in tables:
        _mysql_ensure_column(cur, "plano_hist", "ano", "INT NULL")
        _mysql_ensure_column(cur, "plano_hist", "bloco_id", "VARCHAR(60) NOT NULL")
        _mysql_ensure_column(cur, "plano_hist", "encomenda_numero", "VARCHAR(30) NOT NULL")
        _mysql_ensure_column(cur, "plano_hist", "material", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "plano_hist", "espessura", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "plano_hist", "operacao", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "plano_hist", "posto", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "plano_hist", "data_planeada", "DATE NOT NULL")
        _mysql_ensure_column(cur, "plano_hist", "inicio", "VARCHAR(8) NOT NULL")
        _mysql_ensure_column(cur, "plano_hist", "duracao_min", "DECIMAL(10,2) NOT NULL")
        _mysql_ensure_column(cur, "plano_hist", "color", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "plano_hist", "chapa", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "plano_hist", "movido_em", "DATETIME NULL")
        _mysql_ensure_column(cur, "plano_hist", "estado_final", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "plano_hist", "tempo_planeado_min", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "plano_hist", "tempo_real_min", "DECIMAL(10,2) NULL")
    if "ne_linhas_historico" in tables:
        _mysql_ensure_column(cur, "ne_linhas_historico", "created_at", "DATETIME NOT NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "evento", "VARCHAR(30) NOT NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "origem_menu", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "utilizador", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "guia_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "produto_codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "descricao", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "qtd", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "unid", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "destinatario", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "observacoes", "TEXT NULL")
        _mysql_ensure_column(cur, "ne_linhas_historico", "payload_json", "LONGTEXT NULL")
    if "orc_referencias_historico" in tables:
        _mysql_ensure_column(cur, "orc_referencias_historico", "ref_externa", "VARCHAR(100) NOT NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "ref_interna", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "descricao", "TEXT NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "material", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "espessura", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "preco_unit", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "operacao", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "tempo_peca_min", "DECIMAL(10,3) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "operacoes_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "tempos_operacao_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "custos_operacao_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "quote_cost_snapshot_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "origem_doc", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "origem_tipo", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "estado_origem", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "approved_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "desenho_path", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "orc_referencias_historico", "updated_at", "DATETIME NOT NULL")
    if "orcamento_linhas" in tables:
        _mysql_ensure_column(cur, "orcamento_linhas", "descricao", "TEXT NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "of_codigo", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "desenho_path", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "tempo_peca_min", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "operacoes_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "tempos_operacao_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "custos_operacao_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "quote_cost_snapshot_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "tipo_item", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "produto_codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "produto_unid", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "conjunto_codigo", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "conjunto_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "grupo_uuid", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "orcamento_linhas", "qtd_base", "DECIMAL(10,2) NULL")
    if "orcamentos" in tables:
        _mysql_ensure_column(cur, "orcamentos", "ano", "INT NULL")
        _mysql_ensure_column(cur, "orcamentos", "executado_por", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "orcamentos", "nota_transporte", "TEXT NULL")
        _mysql_ensure_column(cur, "orcamentos", "notas_pdf", "TEXT NULL")
        _mysql_ensure_column(cur, "orcamentos", "desconto_perc", "DECIMAL(6,2) NULL")
        _mysql_ensure_column(cur, "orcamentos", "desconto_valor", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "orcamentos", "subtotal_bruto", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "orcamentos", "preco_transporte", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "orcamentos", "custo_transporte", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "orcamentos", "paletes", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "orcamentos", "peso_bruto_kg", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "orcamentos", "volume_m3", "DECIMAL(12,3) NULL")
        _mysql_ensure_column(cur, "orcamentos", "transportadora_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "orcamentos", "transportadora_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "orcamentos", "meta_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "orcamentos", "referencia_transporte", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "orcamentos", "zona_transporte", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "orcamentos", "posto_trabalho", "VARCHAR(80) NULL")
    if "encomendas" in tables:
        _mysql_ensure_column(cur, "encomendas", "ano", "INT NULL")
        _mysql_ensure_column(cur, "encomendas", "data_entrega", "DATE NULL")
        _mysql_ensure_column(cur, "encomendas", "tempo_estimado", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "encomendas", "cativar", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "encomendas", "observacoes", "TEXT NULL")
        _mysql_ensure_column(cur, "encomendas", "posto_trabalho", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "encomendas", "nota_transporte", "TEXT NULL")
        _mysql_ensure_column(cur, "encomendas", "preco_transporte", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "encomendas", "custo_transporte", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "encomendas", "paletes", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "encomendas", "peso_bruto_kg", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "encomendas", "volume_m3", "DECIMAL(12,3) NULL")
        _mysql_ensure_column(cur, "encomendas", "transportadora_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "encomendas", "transportadora_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "encomendas", "referencia_transporte", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "encomendas", "zona_transporte", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "encomendas", "local_descarga", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "encomendas", "transporte_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "encomendas", "estado_transporte", "VARCHAR(50) NULL")
        _mysql_ensure_column(cur, "encomendas", "tipo_encomenda", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "encomendas", "of_codigo", "VARCHAR(30) NULL")
    if "pecas" in tables:
        _mysql_ensure_column(cur, "pecas", "tipo_material", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "pecas", "subtipo_material", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "pecas", "dimensao", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "pecas", "ficheiros_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "pecas", "perfil_tipo", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "pecas", "perfil_tamanho", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "pecas", "comprimento_mm", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "pecas", "tubo_forma", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "pecas", "lado_a", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "pecas", "lado_b", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "pecas", "tubo_espessura", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "pecas", "diametro", "DECIMAL(10,2) NULL")
    if "transportes" in tables:
        _mysql_ensure_column(cur, "transportes", "transportadora_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "transportes", "transportadora_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "transportes", "referencia_transporte", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "transportes", "custo_previsto", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes", "paletes_total_manual", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "transportes", "peso_total_manual_kg", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes", "volume_total_manual_m3", "DECIMAL(12,3) NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_transporte_estado", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_transporte_ref", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_transporte_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_transporte_by", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_transporte_obs", "TEXT NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_resposta_obs", "TEXT NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_confirmado_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_confirmado_by", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_recusado_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "transportes", "pedido_recusado_by", "VARCHAR(80) NULL")
    if "transportes_paragens" in tables:
        _mysql_ensure_column(cur, "transportes_paragens", "paletes", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "peso_bruto_kg", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "volume_m3", "DECIMAL(12,3) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "preco_transporte", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "custo_transporte", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "transportadora_id", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "transportadora_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "referencia_transporte", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "check_carga_ok", "TINYINT(1) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "check_docs_ok", "TINYINT(1) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "check_paletes_ok", "TINYINT(1) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "pod_estado", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "pod_recebido_nome", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "pod_recebido_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "transportes_paragens", "pod_obs", "TEXT NULL")
    if "pecas" in tables:
        _mysql_ensure_column(cur, "pecas", "observacoes", "TEXT NULL")
        _mysql_ensure_column(cur, "pecas", "desenho_path", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "pecas", "operacoes_fluxo_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "pecas", "hist_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "pecas", "qtd_expedida", "DECIMAL(10,2) NULL")
    if "produtos_mov" in tables:
        _mysql_ensure_column(cur, "produtos_mov", "data", "DATETIME NULL")
        _mysql_ensure_column(cur, "produtos_mov", "tipo", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "produtos_mov", "operador", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "produtos_mov", "codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "produtos_mov", "descricao", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "produtos_mov", "qtd", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "produtos_mov", "antes", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "produtos_mov", "depois", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "produtos_mov", "obs", "TEXT NULL")
        _mysql_ensure_column(cur, "produtos_mov", "origem", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "produtos_mov", "ref_doc", "VARCHAR(50) NULL")
    if "stock_log" in tables:
        _mysql_ensure_column(cur, "stock_log", "operador", "VARCHAR(120) NULL")
    tables = _mysql_existing_tables(cur, force=True)
    if "quality_nonconformities" in tables:
        for name, definition in {
            "origem": "VARCHAR(120) NULL",
            "referencia": "VARCHAR(120) NULL",
            "entidade_tipo": "VARCHAR(60) NULL",
            "entidade_id": "VARCHAR(120) NULL",
            "entidade_label": "VARCHAR(255) NULL",
            "tipo": "VARCHAR(60) NULL",
            "gravidade": "VARCHAR(40) NULL",
            "estado": "VARCHAR(40) NULL",
            "responsavel": "VARCHAR(120) NULL",
            "prazo": "DATE NULL",
            "descricao": "TEXT NULL",
            "causa": "TEXT NULL",
            "acao": "TEXT NULL",
            "eficacia": "TEXT NULL",
            "fornecedor_id": "VARCHAR(30) NULL",
            "fornecedor_nome": "VARCHAR(150) NULL",
            "material_id": "VARCHAR(30) NULL",
            "lote_fornecedor": "VARCHAR(100) NULL",
            "ne_numero": "VARCHAR(30) NULL",
            "guia": "VARCHAR(60) NULL",
            "fatura": "VARCHAR(60) NULL",
            "decisao": "VARCHAR(255) NULL",
            "movement_id": "VARCHAR(255) NULL",
            "qtd_recebida": "DECIMAL(10,2) NULL",
            "qtd_aprovada": "DECIMAL(10,2) NULL",
            "qtd_rejeitada": "DECIMAL(10,2) NULL",
            "qtd_pendente": "DECIMAL(10,2) NULL",
            "created_at": "DATETIME NULL",
            "updated_at": "DATETIME NULL",
            "created_by": "VARCHAR(120) NULL",
            "updated_by": "VARCHAR(120) NULL",
            "closed_at": "DATETIME NULL",
            "closed_by": "VARCHAR(120) NULL",
        }.items():
            _mysql_ensure_column(cur, "quality_nonconformities", name, definition)
    if "quality_documents" in tables:
        for name, definition in {
            "titulo": "VARCHAR(180) NOT NULL",
            "tipo": "VARCHAR(80) NULL",
            "entidade": "VARCHAR(80) NULL",
            "referencia": "VARCHAR(120) NULL",
            "entidade_tipo": "VARCHAR(60) NULL",
            "entidade_id": "VARCHAR(120) NULL",
            "versao": "VARCHAR(30) NULL",
            "estado": "VARCHAR(40) NULL",
            "responsavel": "VARCHAR(120) NULL",
            "caminho": "VARCHAR(512) NULL",
            "obs": "TEXT NULL",
            "created_at": "DATETIME NULL",
            "updated_at": "DATETIME NULL",
            "created_by": "VARCHAR(120) NULL",
            "updated_by": "VARCHAR(120) NULL",
        }.items():
            _mysql_ensure_column(cur, "quality_documents", name, definition)
    if "quality_audit_log" in tables:
        for name, definition in {
            "created_at": "DATETIME NULL",
            "user_name": "VARCHAR(120) NULL",
            "action": "VARCHAR(120) NULL",
            "entity_type": "VARCHAR(80) NULL",
            "entity_id": "VARCHAR(120) NULL",
            "summary": "TEXT NULL",
            "before_json": "LONGTEXT NULL",
            "after_json": "LONGTEXT NULL",
        }.items():
            _mysql_ensure_column(cur, "quality_audit_log", name, definition)
    if "app_config" in tables:
        cur.execute("DELETE FROM app_config WHERE ckey=%s", ("quality_runtime",))
    if "conjuntos_modelo" in tables:
        _mysql_ensure_column(cur, "conjuntos_modelo", "descricao", "VARCHAR(150) NOT NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo", "notas", "TEXT NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo", "ativo", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo", "template", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo", "origem", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo", "created_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo", "updated_at", "DATETIME NULL")
    if "conjuntos" in tables:
        _mysql_ensure_column(cur, "conjuntos", "descricao", "VARCHAR(150) NOT NULL")
        _mysql_ensure_column(cur, "conjuntos", "notas", "TEXT NULL")
        _mysql_ensure_column(cur, "conjuntos", "ativo", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "conjuntos", "template", "BOOLEAN NULL")
        _mysql_ensure_column(cur, "conjuntos", "origem", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "conjuntos", "margem_perc", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "conjuntos", "total_custo", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "conjuntos", "total_final", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "conjuntos", "created_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "conjuntos", "updated_at", "DATETIME NULL")
    if "conjuntos_itens" in tables:
        _mysql_ensure_column(cur, "conjuntos_itens", "linha_ordem", "INT NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "tipo_item", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "ref_externa", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "descricao", "TEXT NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "material", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "espessura", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "operacao", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "produto_codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "produto_unid", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "qtd", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "tempo_peca_min", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "preco_unit", "DECIMAL(10,4) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "desenho_path", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "conjuntos_itens", "meta_json", "LONGTEXT NULL")
        _mysql_ensure_index(cur, "conjuntos_itens", "idx_conjuntos_codigo_ord", "`conjunto_codigo`, `linha_ordem`")
    if "conjuntos_modelo_itens" in tables:
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "linha_ordem", "INT NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "tipo_item", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "ref_externa", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "descricao", "TEXT NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "material", "VARCHAR(100) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "espessura", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "operacao", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "produto_codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "produto_unid", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "qtd", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "tempo_peca_min", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "preco_unit", "DECIMAL(10,4) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "desenho_path", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "conjuntos_modelo_itens", "meta_json", "LONGTEXT NULL")
        _mysql_ensure_index(cur, "conjuntos_modelo_itens", "idx_conjuntos_itens_codigo_ord", "`conjunto_codigo`, `linha_ordem`")
    if "encomenda_montagem_itens" in tables:
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "linha_ordem", "INT NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "tipo_item", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "descricao", "TEXT NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "produto_codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "produto_unid", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "qtd_planeada", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "qtd_consumida", "DECIMAL(10,2) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "preco_unit", "DECIMAL(10,4) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "conjunto_codigo", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "conjunto_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "grupo_uuid", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "estado", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "obs", "TEXT NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "created_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "consumed_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "encomenda_montagem_itens", "consumed_by", "VARCHAR(120) NULL")
        _mysql_ensure_index(cur, "encomenda_montagem_itens", "idx_enc_montagem_num_ord", "`encomenda_numero`, `linha_ordem`")
        _mysql_ensure_index(cur, "encomenda_montagem_itens", "idx_enc_montagem_estado", "`estado`")
    if "faturacao_registos" in tables:
        _mysql_ensure_column(cur, "faturacao_registos", "ano", "INT NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "origem", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "orcamento_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "encomenda_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "cliente_codigo", "VARCHAR(20) NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "cliente_nome", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "data_venda", "DATE NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "data_vencimento", "DATE NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "valor_venda_manual", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "estado_pagamento_manual", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "obs", "TEXT NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "created_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "faturacao_registos", "updated_at", "DATETIME NULL")
        _mysql_ensure_index(cur, "faturacao_registos", "idx_faturacao_registos_orc", "`orcamento_numero`")
        _mysql_ensure_index(cur, "faturacao_registos", "idx_faturacao_registos_enc", "`encomenda_numero`")
        _mysql_ensure_index(cur, "faturacao_registos", "idx_faturacao_registos_cliente", "`cliente_codigo`")
        _mysql_ensure_index(cur, "faturacao_registos", "idx_faturacao_registos_ano", "`ano`")
    if "faturacao_faturas" in tables:
        _mysql_ensure_column(cur, "faturacao_faturas", "documento_id", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "doc_type", "VARCHAR(10) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "numero_fatura", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "serie", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "serie_id", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "seq_num", "INT NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "at_validation_code", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "atcud", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "guia_numero", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "data_emissao", "DATE NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "data_vencimento", "DATE NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "moeda", "VARCHAR(10) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "iva_perc", "DECIMAL(6,2) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "subtotal", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "valor_iva", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "valor_total", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "caminho", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "obs", "TEXT NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "estado", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "anulada", "TINYINT(1) NULL DEFAULT 0")
        _mysql_ensure_column(cur, "faturacao_faturas", "anulada_motivo", "TEXT NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "anulada_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "legal_invoice_no", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "system_entry_date", "DATETIME NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "source_id", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "source_billing", "VARCHAR(1) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "status_source_id", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "hash", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "hash_control", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "previous_hash", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "document_snapshot_json", "LONGTEXT NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "communication_status", "VARCHAR(30) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "communication_filename", "VARCHAR(255) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "communication_error", "TEXT NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "communicated_at", "DATETIME NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "communication_batch_id", "VARCHAR(80) NULL")
        _mysql_ensure_column(cur, "faturacao_faturas", "created_at", "DATETIME NULL")
        _mysql_ensure_index(cur, "faturacao_faturas", "idx_fat_faturas_reg", "`registo_numero`")
        _mysql_ensure_index(cur, "faturacao_faturas", "idx_fat_faturas_doc", "`documento_id`")
        _mysql_ensure_index(cur, "faturacao_faturas", "idx_fat_faturas_emissao", "`data_emissao`")
        _mysql_ensure_index(cur, "faturacao_faturas", "idx_fat_faturas_serie_seq", "`serie_id`, `seq_num`")
        _mysql_ensure_index(cur, "faturacao_faturas", "idx_fat_faturas_legal", "`legal_invoice_no`")
        _mysql_ensure_index(cur, "faturacao_faturas", "idx_fat_faturas_comm", "`communication_status`")
    if "faturacao_pagamentos" in tables:
        _mysql_ensure_column(cur, "faturacao_pagamentos", "pagamento_id", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "fatura_documento_id", "VARCHAR(60) NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "data_pagamento", "DATE NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "valor", "DECIMAL(12,2) NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "metodo", "VARCHAR(40) NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "referencia", "VARCHAR(120) NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "titulo_comprovativo", "VARCHAR(150) NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "caminho_comprovativo", "VARCHAR(512) NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "obs", "TEXT NULL")
        _mysql_ensure_column(cur, "faturacao_pagamentos", "created_at", "DATETIME NULL")
        _mysql_ensure_index(cur, "faturacao_pagamentos", "idx_fat_pag_reg", "`registo_numero`")
        _mysql_ensure_index(cur, "faturacao_pagamentos", "idx_fat_pag_fatura", "`fatura_documento_id`")
        _mysql_ensure_index(cur, "faturacao_pagamentos", "idx_fat_pag_data", "`data_pagamento`")

    if "encomendas" in tables:
        _mysql_ensure_index(cur, "encomendas", "idx_encomendas_estado", "`estado`")
        _mysql_ensure_index(cur, "encomendas", "idx_encomendas_ano", "`ano`")
    if "orcamentos" in tables:
        _mysql_ensure_index(cur, "orcamentos", "idx_orcamentos_estado", "`estado`")
        _mysql_ensure_index(cur, "orcamentos", "idx_orcamentos_ano", "`ano`")
    if "notas_encomenda" in tables:
        _mysql_ensure_index(cur, "notas_encomenda", "idx_ne_ano", "`ano`")
    if "expedicoes" in tables:
        _mysql_ensure_index(cur, "expedicoes", "idx_expedicoes_ano", "`ano`")
    if "plano" in tables:
        _mysql_ensure_index(cur, "plano", "idx_plano_ano", "`ano`")
    if "plano_hist" in tables:
        _mysql_ensure_index(cur, "plano_hist", "idx_plano_hist_ano", "`ano`")
    if "pecas" in tables:
        _mysql_ensure_index(cur, "pecas", "idx_pecas_ref_interna", "`ref_interna`")
    if "fornecedores" in tables:
        _mysql_ensure_index(cur, "fornecedores", "idx_fornecedores_nome", "`nome`")
        _mysql_ensure_index(cur, "fornecedores", "idx_fornecedores_nif", "`nif`")
    if "expedicoes" in tables:
        _mysql_ensure_index(cur, "expedicoes", "idx_expedicoes_serie_seq", "`serie_id`, `seq_num`")
        _mysql_ensure_index(cur, "expedicoes", "idx_expedicoes_atcud", "`atcud`")
    if "at_series" in tables:
        _mysql_ensure_index(cur, "at_series", "idx_at_series_doc_serie", "`doc_type`, `serie_id`")

    # Backfill do ano para suportar filtros/sincronizacao anual.
    try:
        if "orcamentos" in tables:
            cur.execute(
                """
                UPDATE `orcamentos`
                SET `ano` = COALESCE(`ano`, YEAR(`data`), CAST(SUBSTRING(`numero`, 5, 4) AS UNSIGNED))
                WHERE `ano` IS NULL
                """
            )
        if "encomendas" in tables:
            cur.execute(
                """
                UPDATE `encomendas`
                SET `ano` = COALESCE(`ano`, YEAR(`data_criacao`), YEAR(`data_entrega`))
                WHERE `ano` IS NULL
                """
            )
        if "notas_encomenda" in tables:
            cur.execute(
                """
                UPDATE `notas_encomenda`
                SET `ano` = COALESCE(`ano`, YEAR(`data_entrega`), CAST(SUBSTRING(`numero`, 4, 4) AS UNSIGNED))
                WHERE `ano` IS NULL
                """
            )
        if "expedicoes" in tables:
            cur.execute(
                """
                UPDATE `expedicoes`
                SET `ano` = COALESCE(`ano`, YEAR(`data_emissao`), CAST(SUBSTRING(`numero`, 4, 4) AS UNSIGNED))
                WHERE `ano` IS NULL
                """
            )
        if "plano" in tables:
            cur.execute("UPDATE `plano` SET `ano` = COALESCE(`ano`, YEAR(`data_planeada`)) WHERE `ano` IS NULL")
        if "plano_hist" in tables:
            cur.execute("UPDATE `plano_hist` SET `ano` = COALESCE(`ano`, YEAR(`data_planeada`)) WHERE `ano` IS NULL")
        if "materiais" in tables:
            cur.execute(
                """
                UPDATE `materiais`
                SET `preco_unid` = CASE
                    WHEN `formato` = 'Tubo' THEN COALESCE(`metros`, 0) * COALESCE(`p_compra`, 0)
                    WHEN `formato` IN ('Chapa', 'Perfil', 'Cantoneira', 'Barra', 'Varão nervurado') THEN COALESCE(`peso_unid`, 0) * COALESCE(`p_compra`, 0)
                    ELSE COALESCE(`p_compra`, 0)
                END
                WHERE `preco_unid` IS NULL OR `preco_unid` <= 0
                """
            )
    except Exception:
        pass
    ordered = [
        "expedicao_linhas",
        "expedicoes",
        "faturacao_pagamentos",
        "faturacao_faturas",
        "faturacao_registos",
        "at_series",
        "orc_referencias_historico",
        "conjuntos_modelo_itens",
        "conjuntos_modelo",
        "encomenda_montagem_itens",
        "orcamento_linhas",
        "conjuntos_modelo_itens",
        "conjuntos_itens",
        "plano_hist",
        "plano",
        "transportes_paragens",
        "transportes_tarifarios",
        "transportes",
        "encomenda_espessuras",
        "encomenda_reservas",
        "pecas",
        "notas_encomenda_linha_entregas",
        "notas_encomenda_documentos",
        "notas_encomenda_entregas",
        "notas_encomenda_linhas",
        "quality_audit_log",
        "quality_documents",
        "quality_nonconformities",
        "stock_log",
        "produtos_mov",
        "orcamentos",
        "encomendas",
        "notas_encomenda",
        "conjuntos_modelo",
        "conjuntos",
        "produtos",
        "materiais",
        "fornecedores",
        "clientes",
        "operadores",
        "orcamentistas",
        "users",
    ]
    cur.execute("SET FOREIGN_KEY_CHECKS=0")
    try:
        for t in ordered:
            if t in tables:
                cur.execute(f"DELETE FROM `{t}`")
    finally:
        cur.execute("SET FOREIGN_KEY_CHECKS=1")

    if "users" in tables:
        seen_users = set()
        for u in data.get("users", []):
            un = (_clip(u.get("username"), 50) or "").strip()
            if not un:
                continue
            key = un.lower()
            if key in seen_users:
                continue
            seen_users.add(key)
            cur.execute(
                "INSERT INTO users (username, password, role) VALUES (%s, %s, %s)",
                (
                    un,
                    _clip(u.get("password"), 255),
                    _clip(_normalize_role_name(u.get("role")), 50),
                ),
            )

    if "operadores" in tables:
        seen_ops = set()
        for nome in data.get("operadores", []):
            n = (_clip(nome, 120) or "").strip()
            if not n or n.lower() in seen_ops:
                continue
            seen_ops.add(n.lower())
            cur.execute("INSERT INTO operadores (nome) VALUES (%s)", (n,))

    if "orcamentistas" in tables:
        seen_orc = set()
        for nome in data.get("orcamentistas", []):
            n = (_clip(nome, 120) or "").strip()
            if not n or n.lower() in seen_orc:
                continue
            seen_orc.add(n.lower())
            cur.execute("INSERT INTO orcamentistas (nome) VALUES (%s)", (n,))

    if "at_series" in tables:
        seen_series = set()
        for s in data.get("at_series", []):
            if not isinstance(s, dict):
                continue
            doc_type = (_clip(s.get("doc_type"), 10) or "GT").strip().upper()
            serie_id = (_clip(s.get("serie_id"), 40) or "").strip()
            if not serie_id:
                continue
            key = (doc_type, serie_id)
            if key in seen_series:
                continue
            seen_series.add(key)
            inicio_seq = int(parse_float(s.get("inicio_sequencia", 1), 1) or 1)
            next_seq = int(parse_float(s.get("next_seq", inicio_seq), inicio_seq) or inicio_seq)
            if inicio_seq < 1:
                inicio_seq = 1
            if next_seq < inicio_seq:
                next_seq = inicio_seq
            cur.execute(
                """
                INSERT INTO at_series (
                    doc_type, serie_id, inicio_sequencia, next_seq, data_inicio_prevista,
                    validation_code, status, last_error, last_sent_payload_hash, updated_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    doc_type,
                    serie_id,
                    inicio_seq,
                    next_seq,
                    _to_mysql_date(s.get("data_inicio_prevista")),
                    _clip(s.get("validation_code"), 40),
                    _clip(s.get("status"), 20),
                    s.get("last_error"),
                    _clip(s.get("last_sent_payload_hash"), 64),
                    _to_mysql_datetime(s.get("updated_at") or now_iso()),
                ),
            )

    if "clientes" in tables:
        for c in data.get("clientes", []):
            codigo = _clip(c.get("codigo"), 20)
            if not codigo:
                continue
            cur.execute(
                """
                INSERT INTO clientes (codigo, nome, nif, morada, contacto, email)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (
                    codigo,
                    _clip(c.get("nome"), 150),
                    _clip(c.get("nif"), 20),
                    _clip(c.get("morada"), 255),
                    _clip(c.get("contacto"), 50),
                    _clip(c.get("email"), 150),
                ),
            )

    if "fornecedores" in tables:
        for f in data.get("fornecedores", []):
            fid = _clip(f.get("id"), 20)
            if not fid:
                continue
            cur.execute(
                """
                INSERT INTO fornecedores (
                    id, nome, nif, morada, contacto, email,
                    codigo_postal, localidade, pais, cond_pagamento, prazo_entrega_dias, website, obs
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    fid,
                    _clip(f.get("nome"), 150),
                    _clip(f.get("nif"), 20),
                    _clip(f.get("morada"), 255),
                    _clip(f.get("contacto"), 50),
                    _clip(f.get("email"), 150),
                    _clip(f.get("codigo_postal"), 20),
                    _clip(f.get("localidade"), 120),
                    _clip(f.get("pais"), 80),
                    _clip(f.get("cond_pagamento"), 120),
                    int(parse_float(f.get("prazo_entrega_dias", 0), 0) or 0),
                    _clip(f.get("website"), 255),
                    f.get("obs"),
                ),
            )

    if "materiais" in tables:
        for m in data.get("materiais", []):
            mid = _clip(m.get("id"), 20)
            if not mid:
                continue
            cur.execute(
                """
                INSERT INTO materiais (
                    id, lote_fornecedor, formato, material, material_familia, espessura, comprimento, largura, altura, diametro, metros, kg_m, peso_unid, p_compra, preco_unid,
                    quantidade, reservado, tipo, localizacao, is_sobra, atualizado_em, origem_lote, origem_encomenda, secao_tipo, logistic_status,
                    quality_status, quality_blocked, quality_pending_qty, quality_received_qty, quality_approved_qty, quality_return_document_id,
                    inspection_status, inspection_defect, inspection_decision, inspection_at, inspection_by,
                    inspection_note_number, inspection_supplier_id, inspection_supplier_name, inspection_guia, inspection_fatura, quality_nc_id, supplier_claim_id
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    mid,
                    _clip(m.get("lote_fornecedor"), 100),
                    _clip(m.get("formato"), 50),
                    _clip(m.get("material"), 100),
                    _clip(m.get("material_familia"), 40),
                    _clip(m.get("espessura"), 20),
                    _to_num(m.get("comprimento")),
                    _to_num(m.get("largura")),
                    _to_num(m.get("altura")),
                    _to_num(m.get("diametro")),
                    _to_num(m.get("metros")),
                    _to_num(m.get("kg_m")),
                    _to_num(m.get("peso_unid")),
                    _to_num(m.get("p_compra")),
                    _to_num(m.get("preco_unid", materia_preco_unitario(m))),
                    _to_num(m.get("quantidade")),
                    _to_num(m.get("reservado")),
                    _clip(m.get("tipo"), 50),
                    _clip(_get_localizacao(m), 100),
                    1 if _to_bool(m.get("is_sobra")) else 0,
                    _to_mysql_datetime(m.get("atualizado_em")),
                    _clip(m.get("origem_lote"), 100),
                    _clip(m.get("origem_encomenda"), 30),
                    _clip(m.get("secao_tipo"), 40),
                    _clip(m.get("logistic_status"), 30),
                    _clip(m.get("quality_status"), 40),
                    1 if _to_bool(m.get("quality_blocked")) else 0,
                    _to_num(m.get("quality_pending_qty")),
                    _to_num(m.get("quality_received_qty")),
                    _to_num(m.get("quality_approved_qty")),
                    _clip(m.get("quality_return_document_id"), 30),
                    _clip(m.get("inspection_status"), 40),
                    _clip(m.get("inspection_defect"), 255),
                    _clip(m.get("inspection_decision"), 255),
                    _to_mysql_datetime(m.get("inspection_at")),
                    _clip(m.get("inspection_by"), 120),
                    _clip(m.get("inspection_note_number"), 30),
                    _clip(m.get("inspection_supplier_id"), 20),
                    _clip(m.get("inspection_supplier_name"), 150),
                    _clip(m.get("inspection_guia"), 60),
                    _clip(m.get("inspection_fatura"), 60),
                    _clip(m.get("quality_nc_id"), 30),
                    _clip(m.get("supplier_claim_id"), 30),
                ),
            )

    if "produtos" in tables:
        for p in data.get("produtos", []):
            codigo = _clip(p.get("codigo"), 20)
            if not codigo:
                continue
            cur.execute(
                """
                INSERT INTO produtos (
                    codigo, descricao, categoria, category_id, subcat, subcategory_id, tipo, type_id, unid, qty, alerta, p_compra, atualizado_em,
                    logistic_status, quality_status, quality_blocked, quality_pending_qty, quality_received_qty, quality_approved_qty, quality_return_document_id,
                    inspection_defect, inspection_decision, inspection_note_number, quality_nc_id
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    codigo,
                    _clip(p.get("descricao"), 255),
                    _clip(p.get("categoria"), 100),
                    _clip(p.get("category_id"), 80),
                    _clip(p.get("subcat"), 100),
                    _clip(p.get("subcategory_id"), 80),
                    _clip(p.get("tipo"), 100),
                    _clip(p.get("type_id"), 80),
                    _clip(p.get("unid"), 20),
                    _to_num(p.get("qty")),
                    _to_num(p.get("alerta")),
                    _to_num(p.get("p_compra")),
                    _to_mysql_datetime(p.get("atualizado_em")),
                    _clip(p.get("logistic_status"), 30),
                    _clip(p.get("quality_status"), 40),
                    1 if _to_bool(p.get("quality_blocked")) else 0,
                    _to_num(p.get("quality_pending_qty")),
                    _to_num(p.get("quality_received_qty")),
                    _to_num(p.get("quality_approved_qty")),
                    _clip(p.get("quality_return_document_id"), 30),
                    _clip(p.get("inspection_defect"), 255),
                    _clip(p.get("inspection_decision"), 255),
                    _clip(p.get("inspection_note_number"), 30),
                    _clip(p.get("quality_nc_id"), 30),
                ),
            )

    if "conjuntos_modelo" in tables:
        seen_modelos = set()
        for model in data.get("conjuntos_modelo", []):
            if not isinstance(model, dict):
                continue
            codigo = _clip(model.get("codigo"), 40)
            if not codigo or codigo in seen_modelos:
                continue
            seen_modelos.add(codigo)
            cur.execute(
                """
                INSERT INTO conjuntos_modelo (
                    codigo, descricao, notas, ativo, template, origem, created_at, updated_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    codigo,
                    _clip(model.get("descricao"), 150) or codigo,
                    model.get("notas"),
                    1 if _to_bool(model.get("ativo", True)) else 0,
                    1 if _to_bool(model.get("template", False)) else 0,
                    _clip(model.get("origem"), 80),
                    _to_mysql_datetime(model.get("created_at") or now_iso()),
                    _to_mysql_datetime(model.get("updated_at") or now_iso()),
                ),
            )
            if "conjuntos_modelo_itens" in tables:
                for index, item in enumerate(list(model.get("itens", []) or []), start=1):
                    if not isinstance(item, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO conjuntos_modelo_itens (
                            conjunto_codigo, linha_ordem, tipo_item, ref_externa, descricao, material, espessura,
                            operacao, produto_codigo, produto_unid, qtd, tempo_peca_min, preco_unit, desenho_path, meta_json
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            codigo,
                            int(_to_num(item.get("linha_ordem")) or index),
                            _clip(normalize_orc_line_type(item.get("tipo_item")), 30),
                            _clip(item.get("ref_externa"), 100),
                            item.get("descricao"),
                            _clip(item.get("material"), 100),
                            _clip(item.get("espessura"), 20),
                            _clip(item.get("operacao"), 150),
                            _clip(item.get("produto_codigo"), 20),
                            _clip(item.get("produto_unid"), 20),
                            _to_num(item.get("qtd")),
                            _to_num(item.get("tempo_peca_min", item.get("tempo_pecas_min"))),
                            _to_num(item.get("preco_unit")),
                            _clip(item.get("desenho"), 512),
                            json.dumps(
                                {
                                    key: item.get(key)
                                    for key in (
                                        "calc_mode",
                                        "descricao_base",
                                        "weight_total",
                                        "total_cost",
                                        "quantity_units",
                                        "price_per_kg",
                                        "price_base_value",
                                        "price_markup_pct",
                                        "stock_metric_value",
                                        "meters_per_unit",
                                        "kg_per_m",
                                        "length_mm",
                                        "width_mm",
                                        "thickness_mm",
                                        "density",
                                        "diameter_mm",
                                        "manual_unit_price",
                                        "profile_section",
                                        "profile_size",
                                        "tube_section",
                                        "quality",
                                        "stock_material_id",
                                        "hint",
                                        "price_base_label",
                                        "material_family",
                                        "material_subtype",
                                    )
                                    if item.get(key) not in (None, "", [], {})
                                },
                                ensure_ascii=False,
                            ),
                        ),
                    )

    if "conjuntos" in tables:
        seen_conjuntos = set()
        for model in data.get("conjuntos", []):
            if not isinstance(model, dict):
                continue
            codigo = _clip(model.get("codigo"), 40)
            if not codigo or codigo in seen_conjuntos:
                continue
            seen_conjuntos.add(codigo)
            cur.execute(
                """
                INSERT INTO conjuntos (
                    codigo, descricao, notas, ativo, template, origem, margem_perc, total_custo, total_final, created_at, updated_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    codigo,
                    _clip(model.get("descricao"), 150) or codigo,
                    model.get("notas"),
                    1 if _to_bool(model.get("ativo", True)) else 0,
                    1 if _to_bool(model.get("template", False)) else 0,
                    _clip(model.get("origem"), 80),
                    _to_num(model.get("margem_perc")),
                    _to_num(model.get("total_custo")),
                    _to_num(model.get("total_final")),
                    _to_mysql_datetime(model.get("created_at") or now_iso()),
                    _to_mysql_datetime(model.get("updated_at") or now_iso()),
                ),
            )
            if "conjuntos_itens" in tables:
                for index, item in enumerate(list(model.get("itens", []) or []), start=1):
                    if not isinstance(item, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO conjuntos_itens (
                            conjunto_codigo, linha_ordem, tipo_item, ref_externa, descricao, material, espessura,
                            operacao, produto_codigo, produto_unid, qtd, tempo_peca_min, preco_unit, desenho_path, meta_json
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            codigo,
                            int(_to_num(item.get("linha_ordem")) or index),
                            _clip(normalize_orc_line_type(item.get("tipo_item")), 30),
                            _clip(item.get("ref_externa"), 100),
                            item.get("descricao"),
                            _clip(item.get("material"), 100),
                            _clip(item.get("espessura"), 20),
                            _clip(item.get("operacao"), 150),
                            _clip(item.get("produto_codigo"), 20),
                            _clip(item.get("produto_unid"), 20),
                            _to_num(item.get("qtd")),
                            _to_num(item.get("tempo_peca_min", item.get("tempo_pecas_min"))),
                            _to_num(item.get("preco_unit")),
                            _clip(item.get("desenho"), 512),
                            json.dumps(
                                {
                                    key: item.get(key)
                                    for key in (
                                        "calc_mode",
                                        "descricao_base",
                                        "weight_total",
                                        "total_cost",
                                        "quantity_units",
                                        "price_per_kg",
                                        "price_base_value",
                                        "price_markup_pct",
                                        "stock_metric_value",
                                        "meters_per_unit",
                                        "kg_per_m",
                                        "length_mm",
                                        "width_mm",
                                        "thickness_mm",
                                        "density",
                                        "diameter_mm",
                                        "manual_unit_price",
                                        "profile_section",
                                        "profile_size",
                                        "tube_section",
                                        "quality",
                                        "stock_material_id",
                                        "hint",
                                        "price_base_label",
                                        "material_family",
                                        "material_subtype",
                                    )
                                    if item.get(key) not in (None, "", [], {})
                                },
                                ensure_ascii=False,
                            ),
                        ),
                    )

    if "orcamentos" in tables:
        for o in data.get("orcamentos", []):
            if not isinstance(o, dict):
                continue
            num = _clip(o.get("numero"), 30)
            if not num:
                continue
            cliente_cod = _extract_cliente_codigo(o.get("cliente"), data)
            cur.execute(
                """
                INSERT INTO orcamentos (
                    numero, ano, data, estado, cliente_codigo, iva_perc, subtotal, total, numero_encomenda, nota_cliente,
                    executado_por, nota_transporte, notas_pdf, desconto_perc, desconto_valor, subtotal_bruto,
                    preco_transporte, custo_transporte, paletes,
                    peso_bruto_kg, volume_m3, transportadora_id, transportadora_nome, referencia_transporte, zona_transporte, posto_trabalho, meta_json
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    num,
                    _derive_year_from_values(o.get("data"), o.get("numero"), default=datetime.now().year),
                    _to_mysql_datetime(o.get("data")),
                    _clip(o.get("estado"), 50),
                    _clip(cliente_cod, 20),
                    _to_num(o.get("iva_perc")),
                    _to_num(o.get("subtotal")),
                    _to_num(o.get("total")),
                    _clip(o.get("numero_encomenda"), 30),
                    o.get("nota_cliente"),
                    _clip(o.get("executado_por"), 120),
                    o.get("nota_transporte"),
                    o.get("notas_pdf"),
                    _to_num(o.get("desconto_perc")),
                    _to_num(o.get("desconto_valor")),
                    _to_num(o.get("subtotal_bruto")),
                    _to_num(o.get("preco_transporte")),
                    _to_num(o.get("custo_transporte")),
                    _to_num(o.get("paletes")),
                    _to_num(o.get("peso_bruto_kg")),
                    _to_num(o.get("volume_m3")),
                    _clip(o.get("transportadora_id"), 30),
                    _clip(o.get("transportadora_nome"), 150),
                    _clip(o.get("referencia_transporte"), 80),
                    _clip(o.get("zona_transporte"), 120),
                    _clip(o.get("posto_trabalho"), 80),
                    json.dumps(
                        {
                            key: o.get(key)
                            for key in (
                                "desconto_modo",
                                "desconto_grupos",
                            )
                            if o.get(key) not in (None, "", [], {})
                        },
                        ensure_ascii=False,
                    ),
                ),
            )
            if "orcamento_linhas" in tables:
                for l in o.get("linhas", []):
                    if not isinstance(l, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO orcamento_linhas (
                            orcamento_numero, ref_interna, ref_externa, descricao, material, espessura, operacao,
                            of_codigo, qtd, preco_unit, tempo_peca_min, operacoes_json, tempos_operacao_json,
                            custos_operacao_json, quote_cost_snapshot_json, total, desenho_path, tipo_item,
                            produto_codigo, produto_unid, conjunto_codigo, conjunto_nome, grupo_uuid, qtd_base
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            num,
                            _clip(l.get("ref_interna"), 50),
                            _clip(l.get("ref_externa"), 100),
                            l.get("descricao"),
                            _clip(l.get("material"), 100),
                            _clip(l.get("espessura"), 20),
                            _clip(l.get("operacao", l.get("Operacoes", l.get("Operações", ""))), 150),
                            _clip(l.get("of"), 30),
                            _to_num(l.get("qtd")),
                            _to_num(l.get("preco_unit")),
                            _to_num(l.get("tempo_peca_min", l.get("tempo_pecas_min"))),
                            json.dumps(
                                {
                                    "operacoes_lista": list(l.get("operacoes_lista", []) or []),
                                    "operacoes_fluxo": [dict(item or {}) for item in list(l.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                                    "operacoes_detalhe": [dict(item or {}) for item in list(l.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                                },
                                ensure_ascii=False,
                            ),
                            json.dumps(dict(l.get("tempos_operacao", {}) or {}), ensure_ascii=False),
                            json.dumps(dict(l.get("custos_operacao", {}) or {}), ensure_ascii=False),
                            json.dumps(dict(l.get("quote_cost_snapshot", {}) or {}), ensure_ascii=False),
                            _to_num(l.get("total")),
                            _clip(l.get("desenho"), 512),
                            _clip(normalize_orc_line_type(l.get("tipo_item")), 30),
                            _clip(l.get("produto_codigo"), 20),
                            _clip(l.get("produto_unid"), 20),
                            _clip(l.get("conjunto_codigo"), 40),
                            _clip(l.get("conjunto_nome"), 150),
                            _clip(l.get("grupo_uuid"), 60),
                            _to_num(l.get("qtd_base", l.get("qtd"))),
                        ),
                    )

    if "orc_referencias_historico" in tables:
        refs_db = data.get("orc_refs", {}) or {}
        for ref_ext, ref_data in refs_db.items():
            ref_ext_txt = _clip(ref_ext, 100)
            if not ref_ext_txt:
                continue
            row = ref_data if isinstance(ref_data, dict) else {}
            cur.execute(
                """
                INSERT INTO orc_referencias_historico (
                    ref_externa, ref_interna, descricao, material, espessura, preco_unit, operacao, tempo_peca_min,
                    operacoes_json, tempos_operacao_json, custos_operacao_json, quote_cost_snapshot_json,
                    origem_doc, origem_tipo, estado_origem, approved_at, desenho_path, updated_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    ref_ext_txt,
                    _clip(row.get("ref_interna"), 50),
                    row.get("descricao"),
                    _clip(row.get("material"), 100),
                    _clip(row.get("espessura"), 20),
                    _to_num(row.get("preco_unit")),
                    _clip(row.get("operacao"), 150),
                    _to_num(row.get("tempo_peca_min")),
                    json.dumps(
                        {
                            "operacoes_lista": list(row.get("operacoes_lista", []) or []),
                            "operacoes_fluxo": [dict(item or {}) for item in list(row.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                            "operacoes_detalhe": [dict(item or {}) for item in list(row.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                        },
                        ensure_ascii=False,
                    ),
                    json.dumps(dict(row.get("tempos_operacao", {}) or {}), ensure_ascii=False),
                    json.dumps(dict(row.get("custos_operacao", {}) or {}), ensure_ascii=False),
                    json.dumps(dict(row.get("quote_cost_snapshot", {}) or {}), ensure_ascii=False),
                    _clip(row.get("origem_doc"), 30),
                    _clip(row.get("origem_tipo"), 80),
                    _clip(row.get("estado_origem"), 80),
                    _to_mysql_datetime(row.get("approved_at")),
                    _clip(row.get("desenho"), 512),
                    _to_mysql_datetime(now_iso()),
                ),
            )

    if "encomendas" in tables:
        for e in data.get("encomendas", []):
            enc_num = _clip(e.get("numero"), 30)
            if not enc_num:
                continue
            cliente_cod = _extract_cliente_codigo(e.get("cliente"), data)
            cur.execute(
                """
                INSERT INTO encomendas (
                    numero, ano, cliente_codigo, nota_cliente, data_criacao, data_entrega, tempo_estimado,
                    cativar, observacoes, estado, numero_orcamento, posto_trabalho,
                    nota_transporte, preco_transporte, custo_transporte, paletes, peso_bruto_kg, volume_m3,
                    transportadora_id, transportadora_nome, referencia_transporte,
                    zona_transporte, local_descarga, transporte_numero, estado_transporte, tipo_encomenda, of_codigo
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    enc_num,
                    _derive_year_from_values(e.get("data_criacao"), e.get("data_entrega"), e.get("numero"), default=datetime.now().year),
                    _clip(cliente_cod, 20),
                    _clip(e.get("nota_cliente"), 255),
                    _to_mysql_datetime(e.get("data_criacao")),
                    _to_mysql_date(e.get("data_entrega")),
                    _to_num(e.get("tempo_estimado")),
                    1 if _to_bool(e.get("cativar")) else 0,
                    e.get("observacoes", e.get("Observacoes", e.get("Observações", ""))),
                    _clip(e.get("estado"), 50),
                    _clip(e.get("numero_orcamento"), 30),
                    _clip(e.get("posto_trabalho", e.get("posto")), 80),
                    e.get("nota_transporte"),
                    _to_num(e.get("preco_transporte")),
                    _to_num(e.get("custo_transporte")),
                    _to_num(e.get("paletes")),
                    _to_num(e.get("peso_bruto_kg")),
                    _to_num(e.get("volume_m3")),
                    _clip(e.get("transportadora_id"), 30),
                    _clip(e.get("transportadora_nome"), 150),
                    _clip(e.get("referencia_transporte"), 80),
                    _clip(e.get("zona_transporte"), 120),
                    _clip(e.get("local_descarga"), 255),
                    _clip(e.get("transporte_numero"), 30),
                    _clip(e.get("estado_transporte"), 50),
                    _clip(e.get("tipo_encomenda"), 30),
                    _clip(e.get("of_codigo") or (e.get("ordem_fabrico", {}) or {}).get("id"), 30),
                ),
            )
            if "encomenda_espessuras" in tables:
                for m in e.get("materiais", []):
                    if not isinstance(m, dict):
                        continue
                    mat = _clip(m.get("material"), 100)
                    if not mat:
                        continue
                    for esp_obj in m.get("espessuras", []):
                        if not isinstance(esp_obj, dict):
                            continue
                        esp = _clip(esp_obj.get("espessura"), 20)
                        if not esp:
                            continue
                        cur.execute(
                            """
                            INSERT INTO encomenda_espessuras (
                                encomenda_numero, material, espessura, tempo_min, tempos_operacao_json, maquinas_operacao_json, estado,
                                inicio_producao, fim_producao, tempo_producao_min, lote_baixa
                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                            """,
                            (
                                enc_num,
                                mat,
                                esp,
                                _to_num(esp_obj.get("tempo_min")),
                                json.dumps(dict(esp_obj.get("tempos_operacao", {}) or {}), ensure_ascii=False),
                                json.dumps(dict(esp_obj.get("maquinas_operacao", esp_obj.get("recursos_operacao", {})) or {}), ensure_ascii=False),
                                _clip(esp_obj.get("estado"), 50),
                                _to_mysql_datetime(esp_obj.get("inicio_producao")),
                                _to_mysql_datetime(esp_obj.get("fim_producao")),
                                _to_num(esp_obj.get("tempo_producao_min")),
                                _clip(esp_obj.get("lote_baixa"), 100),
                            ),
                        )
            if "encomenda_reservas" in tables:
                for r in e.get("reservas", []):
                    if not isinstance(r, dict):
                        continue
                    mat_r = _clip(r.get("material"), 100)
                    esp_r = _clip(r.get("espessura"), 20)
                    qtd_r = _to_num(r.get("quantidade"))
                    if not mat_r or not esp_r or qtd_r is None:
                        continue
                    cur.execute(
                        """
                        INSERT INTO encomenda_reservas (
                            encomenda_numero, material_id, material, espessura, quantidade, created_at
                        ) VALUES (%s, %s, %s, %s, %s, %s)
                        """,
                        (
                            enc_num,
                            _clip(r.get("material_id"), 30),
                            mat_r,
                            esp_r,
                            qtd_r,
                            _to_mysql_datetime(now_iso()),
                        ),
                    )
            if "encomenda_montagem_itens" in tables:
                for index, item in enumerate(encomenda_montagem_itens(e), start=1):
                    cur.execute(
                        """
                        INSERT INTO encomenda_montagem_itens (
                            encomenda_numero, linha_ordem, tipo_item, descricao, produto_codigo, produto_unid,
                            qtd_planeada, qtd_consumida, preco_unit, conjunto_codigo, conjunto_nome, grupo_uuid,
                            estado, obs, created_at, consumed_at, consumed_by
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            enc_num,
                            int(_to_num(item.get("linha_ordem")) or index),
                            _clip(normalize_orc_line_type(item.get("tipo_item")), 30),
                            item.get("descricao"),
                            _clip(item.get("produto_codigo"), 20),
                            _clip(item.get("produto_unid"), 20),
                            _to_num(item.get("qtd_planeada", item.get("qtd"))),
                            _to_num(item.get("qtd_consumida")),
                            _to_num(item.get("preco_unit")),
                            _clip(item.get("conjunto_codigo"), 40),
                            _clip(item.get("conjunto_nome"), 150),
                            _clip(item.get("grupo_uuid"), 60),
                            _clip(item.get("estado"), 30),
                            item.get("obs"),
                            _to_mysql_datetime(item.get("created_at") or now_iso()),
                            _to_mysql_datetime(item.get("consumed_at")),
                            _clip(item.get("consumed_by"), 120),
                        ),
                    )

    if "plano" in tables:
        for item in data.get("plano", []):
            if not isinstance(item, dict):
                continue
            enc_num = _clip(item.get("encomenda"), 30)
            data_planeada = _to_mysql_date(item.get("data"))
            inicio = _clip(item.get("inicio"), 8)
            duracao_min = _to_num(item.get("duracao_min"))
            if not enc_num or not data_planeada or not inicio or duracao_min is None:
                continue
            bloco_id = _clip(item.get("id"), 60) or _clip(
                f"PL-{enc_num}-{_clip(item.get('material'), 20)}-{_clip(item.get('espessura'), 10)}-{data_planeada}-{inicio}-{int(duracao_min)}",
                60,
            )
            cur.execute(
                """
                INSERT INTO plano (
                    bloco_id, encomenda_numero, ano, material, espessura, operacao, posto, data_planeada, inicio, duracao_min, color, chapa
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    bloco_id,
                    enc_num,
                    _derive_year_from_values(item.get("data"), enc_num, default=datetime.now().year),
                    _clip(item.get("material"), 100),
                    _clip(item.get("espessura"), 20),
                    _clip(item.get("operacao"), 80),
                    _clip(item.get("posto") or item.get("posto_trabalho") or item.get("maquina"), 80),
                    data_planeada,
                    inicio,
                    duracao_min,
                    _clip(item.get("color"), 20),
                    _clip(item.get("chapa"), 120),
                ),
            )

    if "plano_hist" in tables:
        for item in data.get("plano_hist", []):
            if not isinstance(item, dict):
                continue
            enc_num = _clip(item.get("encomenda"), 30)
            data_planeada = _to_mysql_date(item.get("data"))
            inicio = _clip(item.get("inicio"), 8)
            duracao_min = _to_num(item.get("duracao_min"))
            if not enc_num or not data_planeada or not inicio or duracao_min is None:
                continue
            bloco_id = _clip(item.get("id"), 60) or _clip(
                f"PLH-{enc_num}-{_clip(item.get('material'), 20)}-{_clip(item.get('espessura'), 10)}-{data_planeada}-{inicio}-{int(duracao_min)}",
                60,
            )
            cur.execute(
                """
                INSERT INTO plano_hist (
                    bloco_id, encomenda_numero, ano, material, espessura, operacao, posto, data_planeada, inicio, duracao_min, color, chapa,
                    movido_em, estado_final, tempo_planeado_min, tempo_real_min
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    bloco_id,
                    enc_num,
                    _derive_year_from_values(item.get("data"), enc_num, default=datetime.now().year),
                    _clip(item.get("material"), 100),
                    _clip(item.get("espessura"), 20),
                    _clip(item.get("operacao"), 80),
                    _clip(item.get("posto") or item.get("posto_trabalho") or item.get("maquina"), 80),
                    data_planeada,
                    inicio,
                    duracao_min,
                    _clip(item.get("color"), 20),
                    _clip(item.get("chapa"), 120),
                    _to_mysql_datetime(item.get("movido_em")),
                    _clip(item.get("estado_final"), 50),
                    _to_num(item.get("tempo_planeado_min")),
                    _to_num(item.get("tempo_real_min")),
                ),
            )

    if "pecas" in tables:
        used_ids = set()
        for e in data.get("encomendas", []):
            enc_num = _clip(e.get("numero"), 30)
            if not enc_num:
                continue
            for idx, p in enumerate(encomenda_pecas(e), start=1):
                pid = _clip(p.get("id"), 30) or _clip(f"{enc_num}-{idx}", 30)
                base_pid = pid
                n = 2
                while pid in used_ids:
                    suffix = f"-{n}"
                    pid = (base_pid[: max(1, 30 - len(suffix))] + suffix)[:30]
                    n += 1
                used_ids.add(pid)
                cur.execute(
                    """
                    INSERT INTO pecas (
                        id, encomenda_numero, ref_interna, ref_externa, material, espessura, quantidade_pedida, operacoes,
                        of_codigo, opp_codigo, estado, produzido_ok, produzido_nok, inicio_producao, fim_producao,
                        tempo_producao_min, lote_baixa, observacoes, desenho_path, operacoes_fluxo_json, hist_json, qtd_expedida,
                        tipo_material, subtipo_material, dimensao, ficheiros_json,
                        perfil_tipo, perfil_tamanho, comprimento_mm, tubo_forma, lado_a, lado_b, tubo_espessura, diametro
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        pid,
                        enc_num,
                        _clip(p.get("ref_interna"), 50),
                        _clip(p.get("ref_externa"), 100),
                        _clip(p.get("material"), 100),
                        _clip(p.get("espessura"), 20),
                        _to_num(p.get("quantidade_pedida")),
                        _clip(p.get("Operacoes", p.get("operacoes", p.get("Operações", ""))), 150),
                        _clip(p.get("of"), 30),
                        _clip(p.get("opp"), 30),
                        _clip(p.get("estado"), 50),
                        _to_num(p.get("produzido_ok")),
                        _to_num(p.get("produzido_nok")),
                        _to_mysql_datetime(p.get("inicio_producao")),
                        _to_mysql_datetime(p.get("fim_producao")),
                        _to_num(p.get("tempo_producao_min")),
                        _clip(p.get("lote_baixa"), 100),
                        p.get("Observacoes", p.get("Observações")),
                        _clip(p.get("desenho"), 512),
                        json.dumps(ensure_peca_operacoes(p), ensure_ascii=False),
                        json.dumps(p.get("hist", []), ensure_ascii=False),
                        _to_num(p.get("qtd_expedida")),
                        _clip(p.get("tipo_material"), 30),
                        _clip(p.get("subtipo_material"), 100),
                        _clip(p.get("dimensao", p.get("dimensoes")), 120),
                        json.dumps(list(p.get("ficheiros", []) or []), ensure_ascii=False),
                        _clip(p.get("perfil_tipo"), 40),
                        _clip(p.get("perfil_tamanho"), 80),
                        _to_num(p.get("comprimento_mm")),
                        _clip(p.get("tubo_forma"), 40),
                        _to_num(p.get("lado_a")),
                        _to_num(p.get("lado_b")),
                        _to_num(p.get("tubo_espessura")),
                        _to_num(p.get("diametro")),
                    ),
                )

    if "notas_encomenda" in tables:
        for ne in data.get("notas_encomenda", []):
            ne_num = _clip(ne.get("numero"), 30)
            if not ne_num:
                continue
            fornecedor_id = _extract_fornecedor_id(ne.get("fornecedor_id"), data)
            if not fornecedor_id:
                fornecedor_id = _extract_fornecedor_id(ne.get("fornecedor"), data)
            ne_geradas_val = ne.get("ne_geradas", [])
            if isinstance(ne_geradas_val, (list, tuple, set)):
                ne_geradas_txt = ",".join(str(x).strip() for x in ne_geradas_val if str(x).strip())
            else:
                ne_geradas_txt = str(ne_geradas_val or "").strip()
            cur.execute(
                """
                INSERT INTO notas_encomenda (
                    numero, ano, fornecedor_id, contacto, data_entrega, estado, total, obs, local_descarga, meio_transporte,
                    oculta, is_draft, data_ultima_entrega, guia_ultima, fatura_ultima, fatura_caminho_ultima, data_doc_ultima,
                    origem_cotacao, ne_geradas
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    ne_num,
                    _derive_year_from_values(ne.get("data_entrega"), ne.get("numero"), default=datetime.now().year),
                    _clip(fornecedor_id, 20),
                    _clip(ne.get("contacto"), 80),
                    _to_mysql_date(ne.get("data_entrega")),
                    _clip(ne.get("estado"), 50),
                    _to_num(ne.get("total")),
                    ne.get("obs"),
                    _clip(ne.get("local_descarga"), 255),
                    _clip(ne.get("meio_transporte"), 100),
                    1 if _to_bool(ne.get("oculta")) else 0,
                    1 if _to_bool(ne.get("_draft")) else 0,
                    _to_mysql_date(ne.get("data_ultima_entrega")),
                    _clip(ne.get("guia_ultima"), 60),
                    _clip(ne.get("fatura_ultima"), 60),
                    _clip(ne.get("fatura_caminho_ultima"), 512),
                    _to_mysql_date(ne.get("data_doc_ultima")),
                    _clip(ne.get("origem_cotacao"), 30),
                    ne_geradas_txt,
                ),
            )
            if "notas_encomenda_entregas" in tables:
                for ent in ne.get("entregas", []):
                    if not isinstance(ent, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO notas_encomenda_entregas (
                            ne_numero, data_registo, data_entrega, data_documento, guia, fatura, obs
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            ne_num,
                            _to_mysql_datetime(ent.get("data_registo")),
                            _to_mysql_date(ent.get("data_entrega")),
                            _to_mysql_date(ent.get("data_documento")),
                            _clip(ent.get("guia"), 60),
                            _clip(ent.get("fatura"), 60),
                            ent.get("obs"),
                        ),
                    )
            if "notas_encomenda_documentos" in tables:
                for doc in ne.get("documentos", []):
                    if not isinstance(doc, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO notas_encomenda_documentos (
                            ne_numero, data_registo, tipo, titulo, caminho, guia, fatura, data_entrega, data_documento, obs
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            ne_num,
                            _to_mysql_datetime(doc.get("data_registo")),
                            _clip(doc.get("tipo"), 40),
                            _clip(doc.get("titulo"), 150),
                            _clip(doc.get("caminho"), 512),
                            _clip(doc.get("guia"), 60),
                            _clip(doc.get("fatura"), 60),
                            _to_mysql_date(doc.get("data_entrega")),
                            _to_mysql_date(doc.get("data_documento")),
                            doc.get("obs"),
                        ),
                    )
            if "notas_encomenda_linhas" in tables:
                for idx_l, l in enumerate(ne.get("linhas", []), start=1):
                    if not isinstance(l, dict):
                        continue
                    qtd_l = _to_num(l.get("qtd"))
                    qtd_ent_l = _to_num(l.get("qtd_entregue"))
                    if qtd_ent_l is None:
                        qtd_ent_l = qtd_l if _to_bool(l.get("entregue")) else 0.0
                    cur.execute(
                        """
                        INSERT INTO notas_encomenda_linhas (
                            ne_numero, linha_ordem, ref_material, descricao, fornecedor_linha, origem, qtd, unid,
                            preco, desconto, iva, total, entregue, qtd_entregue, lote_fornecedor, material, espessura, comprimento,
                            largura, altura, diametro, metros, kg_m, localizacao, peso_unid, p_compra, formato, material_familia, secao_tipo, stock_in, guia_entrega,
                            fatura_entrega, data_doc_entrega, data_entrega_real, obs_entrega, logistic_status, inspection_status, inspection_defect, inspection_decision, quality_status, quality_nc_id
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            ne_num,
                            idx_l,
                            _clip(l.get("ref") or l.get("material"), 20),
                            _clip(l.get("descricao"), 255),
                            _clip(l.get("fornecedor_linha"), 150),
                            _clip(l.get("origem"), 50),
                            qtd_l,
                            _clip(l.get("unid"), 20),
                            _to_num(l.get("preco")),
                            _to_num(l.get("desconto")),
                            _to_num(l.get("iva")),
                            _to_num(l.get("total")),
                            1 if _to_bool(l.get("entregue")) else 0,
                            qtd_ent_l,
                            _clip(l.get("lote_fornecedor"), 100),
                            _clip(l.get("material"), 100),
                            _clip(l.get("espessura"), 20),
                            _to_num(l.get("comprimento")),
                            _to_num(l.get("largura")),
                            _to_num(l.get("altura")),
                            _to_num(l.get("diametro")),
                            _to_num(l.get("metros")),
                            _to_num(l.get("kg_m")),
                            _clip(l.get("localizacao"), 100),
                            _to_num(l.get("peso_unid")),
                            _to_num(l.get("p_compra")),
                            _clip(l.get("formato"), 50),
                            _clip(l.get("material_familia"), 40),
                            _clip(l.get("secao_tipo"), 40),
                            1 if _to_bool(l.get("_stock_in")) else 0,
                            _clip(l.get("guia_entrega"), 60),
                            _clip(l.get("fatura_entrega"), 60),
                            _to_mysql_date(l.get("data_doc_entrega")),
                            _to_mysql_date(l.get("data_entrega_real")),
                            l.get("obs_entrega"),
                            _clip(l.get("logistic_status"), 30),
                            _clip(l.get("inspection_status"), 40),
                            _clip(l.get("inspection_defect"), 255),
                            _clip(l.get("inspection_decision"), 255),
                            _clip(l.get("quality_status"), 40),
                            _clip(l.get("quality_nc_id"), 30),
                        ),
                    )
                    if "notas_encomenda_linha_entregas" in tables:
                        for ent_l in l.get("entregas_linha", []):
                            if not isinstance(ent_l, dict):
                                continue
                            cur.execute(
                                """
                                INSERT INTO notas_encomenda_linha_entregas (
                                    ne_numero, linha_ordem, data_registo, data_entrega, data_documento, guia, fatura, obs, qtd,
                                    lote_fornecedor, localizacao, entrega_total, stock_ref, logistic_status, inspection_status, inspection_defect,
                                    inspection_decision, quality_status, quality_nc_id
                                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                """,
                                (
                                    ne_num,
                                    idx_l,
                                    _to_mysql_datetime(ent_l.get("data_registo")),
                                    _to_mysql_date(ent_l.get("data_entrega")),
                                    _to_mysql_date(ent_l.get("data_documento")),
                                    _clip(ent_l.get("guia"), 60),
                                    _clip(ent_l.get("fatura"), 60),
                                    ent_l.get("obs"),
                                    _to_num(ent_l.get("qtd")),
                                    _clip(ent_l.get("lote_fornecedor"), 100),
                                    _clip(ent_l.get("localizacao"), 100),
                                    1 if _to_bool(ent_l.get("entrega_total")) else 0,
                                    _clip(ent_l.get("stock_ref"), 30),
                                    _clip(ent_l.get("logistic_status"), 30),
                                    _clip(ent_l.get("inspection_status"), 40),
                                    _clip(ent_l.get("inspection_defect"), 255),
                                    _clip(ent_l.get("inspection_decision"), 255),
                                    _clip(ent_l.get("quality_status"), 40),
                                    _clip(ent_l.get("quality_nc_id"), 30),
                                ),
                            )

    if "expedicoes" in tables:
        for ex in data.get("expedicoes", []):
            if not isinstance(ex, dict):
                continue
            num = _clip(ex.get("numero"), 30)
            if not num:
                continue
            cur.execute(
                """
                INSERT INTO expedicoes (
                    numero, ano, tipo, encomenda_numero, cliente_codigo, cliente_nome, codigo_at, serie_id, seq_num, at_validation_code, atcud,
                    emitente_nome, emitente_nif, emitente_morada,
                    destinatario, dest_nif, dest_morada,
                    local_carga, local_descarga, data_emissao, data_transporte, matricula, transportador, estado,
                    observacoes, created_by, anulada, anulada_motivo
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    num,
                    _derive_year_from_values(ex.get("data_emissao"), ex.get("data_transporte"), ex.get("numero"), default=datetime.now().year),
                    _clip(ex.get("tipo"), 30),
                    _clip(ex.get("encomenda"), 30),
                    _clip(ex.get("cliente"), 20),
                    _clip(ex.get("cliente_nome"), 150),
                    _clip(ex.get("codigo_at"), 80),
                    _clip(ex.get("serie_id"), 40),
                    int(parse_float(ex.get("seq_num", 0), 0) or 0) or None,
                    _clip(ex.get("at_validation_code"), 40),
                    _clip(ex.get("atcud"), 120),
                    _clip(ex.get("emitente_nome"), 150),
                    _clip(ex.get("emitente_nif"), 20),
                    _clip(ex.get("emitente_morada"), 255),
                    _clip(ex.get("destinatario"), 150),
                    _clip(ex.get("dest_nif"), 20),
                    _clip(ex.get("dest_morada"), 255),
                    _clip(ex.get("local_carga"), 255),
                    _clip(ex.get("local_descarga"), 255),
                    _to_mysql_datetime(ex.get("data_emissao")),
                    _to_mysql_datetime(ex.get("data_transporte")),
                    _clip(ex.get("matricula"), 30),
                    _clip(ex.get("transportador"), 150),
                    _clip(ex.get("estado"), 50),
                    ex.get("observacoes"),
                    _clip(ex.get("created_by"), 80),
                    1 if _to_bool(ex.get("anulada")) else 0,
                    ex.get("anulada_motivo"),
                ),
            )
            if "expedicao_linhas" in tables:
                for l in ex.get("linhas", []):
                    if not isinstance(l, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO expedicao_linhas (
                            expedicao_numero, encomenda_numero, peca_id, ref_interna, ref_externa, descricao,
                            qtd, unid, peso, manual
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            num,
                            _clip(l.get("encomenda"), 30),
                            _clip(l.get("peca_id"), 30),
                            _clip(l.get("ref_interna"), 50),
                            _clip(l.get("ref_externa"), 100),
                            _clip(l.get("descricao"), 255),
                            _to_num(l.get("qtd")),
                            _clip(l.get("unid"), 20),
                            _to_num(l.get("peso")),
                            1 if _to_bool(l.get("manual")) else 0,
                        ),
                    )

    if "transportes" in tables:
        for tr in data.get("transportes", []):
            if not isinstance(tr, dict):
                continue
            num = _clip(tr.get("numero"), 30)
            if not num:
                continue
            cur.execute(
                """
                INSERT INTO transportes (
                    numero, ano, tipo_responsavel, estado, data_planeada, hora_saida, viatura, matricula,
                    motorista, telefone_motorista, origem, transportadora_id, transportadora_nome,
                    referencia_transporte, custo_previsto, paletes_total_manual, peso_total_manual_kg,
                    volume_total_manual_m3, pedido_transporte_estado, pedido_transporte_ref,
                    pedido_transporte_at, pedido_transporte_by, pedido_transporte_obs,
                    pedido_resposta_obs, pedido_confirmado_at, pedido_confirmado_by,
                    pedido_recusado_at, pedido_recusado_by,
                    observacoes, created_by, created_at, updated_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    num,
                    _derive_year_from_values(tr.get("data_planeada"), tr.get("numero"), default=datetime.now().year),
                    _clip(tr.get("tipo_responsavel"), 40),
                    _clip(tr.get("estado"), 40),
                    _to_mysql_date(tr.get("data_planeada")),
                    _clip(tr.get("hora_saida"), 10),
                    _clip(tr.get("viatura"), 120),
                    _clip(tr.get("matricula"), 30),
                    _clip(tr.get("motorista"), 120),
                    _clip(tr.get("telefone_motorista"), 40),
                    _clip(tr.get("origem"), 255),
                    _clip(tr.get("transportadora_id"), 30),
                    _clip(tr.get("transportadora_nome"), 150),
                    _clip(tr.get("referencia_transporte"), 80),
                    _to_num(tr.get("custo_previsto")),
                    _to_num(tr.get("paletes_total_manual")),
                    _to_num(tr.get("peso_total_manual_kg")),
                    _to_num(tr.get("volume_total_manual_m3")),
                    _clip(tr.get("pedido_transporte_estado"), 40),
                    _clip(tr.get("pedido_transporte_ref"), 80),
                    _to_mysql_datetime(tr.get("pedido_transporte_at")),
                    _clip(tr.get("pedido_transporte_by"), 80),
                    tr.get("pedido_transporte_obs"),
                    tr.get("pedido_resposta_obs"),
                    _to_mysql_datetime(tr.get("pedido_confirmado_at")),
                    _clip(tr.get("pedido_confirmado_by"), 80),
                    _to_mysql_datetime(tr.get("pedido_recusado_at")),
                    _clip(tr.get("pedido_recusado_by"), 80),
                    tr.get("observacoes"),
                    _clip(tr.get("created_by"), 80),
                    _to_mysql_datetime(tr.get("created_at")),
                    _to_mysql_datetime(tr.get("updated_at")),
                ),
            )
            if "transportes_paragens" in tables:
                for index, stop in enumerate(list(tr.get("paragens", []) or []), start=1):
                    if not isinstance(stop, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO transportes_paragens (
                            transporte_numero, ordem, encomenda_numero, expedicao_numero, cliente_codigo, cliente_nome,
                            zona_transporte, local_descarga, contacto, telefone, data_planeada, paletes, peso_bruto_kg, volume_m3,
                            preco_transporte, custo_transporte, transportadora_id, transportadora_nome,
                            referencia_transporte, check_carga_ok, check_docs_ok, check_paletes_ok,
                            pod_estado, pod_recebido_nome, pod_recebido_at, pod_obs,
                            estado, observacoes
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            num,
                            int(_to_num(stop.get("ordem")) or index),
                            _clip(stop.get("encomenda_numero", stop.get("encomenda")), 30),
                            _clip(stop.get("expedicao_numero", stop.get("guia_numero")), 30),
                            _clip(stop.get("cliente_codigo"), 20),
                            _clip(stop.get("cliente_nome"), 150),
                            _clip(stop.get("zona_transporte"), 120),
                            _clip(stop.get("local_descarga"), 255),
                            _clip(stop.get("contacto"), 120),
                            _clip(stop.get("telefone"), 40),
                            _to_mysql_datetime(stop.get("data_planeada")),
                            _to_num(stop.get("paletes")),
                            _to_num(stop.get("peso_bruto_kg")),
                            _to_num(stop.get("volume_m3")),
                            _to_num(stop.get("preco_transporte")),
                            _to_num(stop.get("custo_transporte")),
                            _clip(stop.get("transportadora_id"), 30),
                            _clip(stop.get("transportadora_nome"), 150),
                            _clip(stop.get("referencia_transporte"), 80),
                            1 if _to_bool(stop.get("check_carga_ok")) else 0,
                            1 if _to_bool(stop.get("check_docs_ok")) else 0,
                            1 if _to_bool(stop.get("check_paletes_ok")) else 0,
                            _clip(stop.get("pod_estado"), 40),
                            _clip(stop.get("pod_recebido_nome"), 120),
                            _to_mysql_datetime(stop.get("pod_recebido_at")),
                            stop.get("pod_obs"),
                            _clip(stop.get("estado"), 40),
                            stop.get("observacoes"),
                        ),
                    )
    if "transportes_tarifarios" in tables:
        for row in data.get("transportes_tarifarios", []):
            if not isinstance(row, dict):
                continue
            zona_txt = _clip(row.get("zona"), 120)
            if not zona_txt:
                continue
            cur.execute(
                """
                INSERT INTO transportes_tarifarios (
                    id, transportadora_id, transportadora_nome, zona, valor_base, valor_por_palete,
                    valor_por_kg, valor_por_m3, custo_minimo, ativo, observacoes
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    _to_num(row.get("id")),
                    _clip(row.get("transportadora_id"), 30),
                    _clip(row.get("transportadora_nome"), 150),
                    zona_txt,
                    _to_num(row.get("valor_base")),
                    _to_num(row.get("valor_por_palete")),
                    _to_num(row.get("valor_por_kg")),
                    _to_num(row.get("valor_por_m3")),
                    _to_num(row.get("custo_minimo")),
                    1 if _to_bool(row.get("ativo", True)) else 0,
                    row.get("observacoes"),
                ),
            )

    if "faturacao_registos" in tables:
        for reg in list(data.get("faturacao", []) or []):
            if not isinstance(reg, dict):
                continue
            reg_num = _clip(reg.get("numero"), 30)
            if not reg_num:
                continue
            cur.execute(
                """
                INSERT INTO faturacao_registos (
                    numero, ano, origem, orcamento_numero, encomenda_numero, cliente_codigo, cliente_nome,
                    data_venda, data_vencimento, valor_venda_manual, estado_pagamento_manual, obs, created_at, updated_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    reg_num,
                    _derive_year_from_values(reg.get("data_venda"), reg.get("numero"), default=datetime.now().year),
                    _clip(reg.get("origem"), 30),
                    _clip(reg.get("orcamento_numero"), 30),
                    _clip(reg.get("encomenda_numero"), 30),
                    _clip(reg.get("cliente_codigo"), 20),
                    _clip(reg.get("cliente_nome"), 150),
                    _to_mysql_date(reg.get("data_venda")),
                    _to_mysql_date(reg.get("data_vencimento")),
                    _to_num(reg.get("valor_venda_manual")),
                    _clip(reg.get("estado_pagamento_manual"), 30),
                    reg.get("obs"),
                    _to_mysql_datetime(reg.get("created_at") or now_iso()),
                    _to_mysql_datetime(reg.get("updated_at") or now_iso()),
                ),
            )
            if "faturacao_faturas" in tables:
                for row in list(reg.get("faturas", []) or []):
                    if not isinstance(row, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO faturacao_faturas (
                            registo_numero, documento_id, doc_type, numero_fatura, serie, serie_id, seq_num,
                            at_validation_code, atcud, guia_numero, data_emissao, data_vencimento,
                            moeda, iva_perc, subtotal, valor_iva, valor_total, caminho, obs,
                            estado, anulada, anulada_motivo, anulada_at, legal_invoice_no,
                            system_entry_date, source_id, source_billing, status_source_id,
                            hash, hash_control, previous_hash, document_snapshot_json,
                            communication_status, communication_filename, communication_error,
                            communicated_at, communication_batch_id, created_at
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            reg_num,
                            _clip(row.get("id"), 60),
                            _clip(row.get("doc_type"), 10),
                            _clip(row.get("numero_fatura"), 60),
                            _clip(row.get("serie"), 40),
                            _clip(row.get("serie_id"), 40),
                            int(_to_num(row.get("seq_num")) or 0) or None,
                            _clip(row.get("at_validation_code"), 40),
                            _clip(row.get("atcud"), 80),
                            _clip(row.get("guia_numero"), 30),
                            _to_mysql_date(row.get("data_emissao")),
                            _to_mysql_date(row.get("data_vencimento")),
                            _clip(row.get("moeda"), 10),
                            _to_num(row.get("iva_perc")),
                            _to_num(row.get("subtotal")),
                            _to_num(row.get("valor_iva")),
                            _to_num(row.get("valor_total")),
                            _clip(row.get("caminho"), 512),
                            row.get("obs"),
                            _clip(row.get("estado"), 30),
                            1 if bool(row.get("anulada")) else 0,
                            row.get("anulada_motivo"),
                            _to_mysql_datetime(row.get("anulada_at")),
                            _clip(row.get("legal_invoice_no"), 80),
                            _to_mysql_datetime(row.get("system_entry_date")),
                            _clip(row.get("source_id"), 80),
                            _clip(row.get("source_billing"), 1),
                            _clip(row.get("status_source_id"), 80),
                            _clip(row.get("hash"), 512),
                            _clip(row.get("hash_control"), 80),
                            _clip(row.get("previous_hash"), 512),
                            row.get("document_snapshot_json"),
                            _clip(row.get("communication_status"), 30),
                            _clip(row.get("communication_filename"), 255),
                            row.get("communication_error"),
                            _to_mysql_datetime(row.get("communicated_at")),
                            _clip(row.get("communication_batch_id"), 80),
                            _to_mysql_datetime(row.get("created_at") or now_iso()),
                        ),
                    )
            if "faturacao_pagamentos" in tables:
                for row in list(reg.get("pagamentos", []) or []):
                    if not isinstance(row, dict):
                        continue
                    cur.execute(
                        """
                        INSERT INTO faturacao_pagamentos (
                            registo_numero, pagamento_id, fatura_documento_id, data_pagamento, valor, metodo,
                            referencia, titulo_comprovativo, caminho_comprovativo, obs, created_at
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            reg_num,
                            _clip(row.get("id"), 60),
                            _clip(row.get("fatura_id"), 60),
                            _to_mysql_date(row.get("data_pagamento")),
                            _to_num(row.get("valor")),
                            _clip(row.get("metodo"), 40),
                            _clip(row.get("referencia"), 120),
                            _clip(row.get("titulo_comprovativo"), 150),
                            _clip(row.get("caminho_comprovativo"), 512),
                            row.get("obs"),
                            _to_mysql_datetime(row.get("created_at") or now_iso()),
                        ),
                    )

    if "stock_log" in tables:
        for s in data.get("stock_log", []):
            cur.execute(
                "INSERT INTO stock_log (data, acao, operador, detalhes) VALUES (%s, %s, %s, %s)",
                (
                    _to_mysql_datetime(s.get("data")),
                    _clip(s.get("acao"), 50),
                    _clip(s.get("operador"), 120),
                    s.get("detalhes"),
                ),
            )

    if "produtos_mov" in tables:
        for r in data.get("produtos_mov", []):
            mov = normalize_produto_mov_row(r)
            if not mov:
                continue
            cur.execute(
                """
                INSERT INTO produtos_mov (
                    data, tipo, operador, codigo, descricao, qtd, antes, depois, obs, origem, ref_doc
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    _to_mysql_datetime(mov.get("data")),
                    _clip(mov.get("tipo"), 40),
                    _clip(mov.get("operador"), 120),
                    _clip(mov.get("codigo"), 20),
                    _clip(mov.get("descricao"), 255),
                    _to_num(mov.get("qtd")),
                    _to_num(mov.get("antes")),
                    _to_num(mov.get("depois")),
                    mov.get("obs"),
                    _clip(mov.get("origem"), 80),
                    _clip(mov.get("ref_doc"), 50),
                ),
            )

    if "quality_nonconformities" in tables:
        for row in data.get("quality_nonconformities", []):
            if not isinstance(row, dict) or not str(row.get("id", "") or "").strip():
                continue
            cur.execute(
                """
                INSERT INTO quality_nonconformities (
                    id, origem, referencia, entidade_tipo, entidade_id, entidade_label, tipo, gravidade,
                    estado, responsavel, prazo, descricao, causa, acao, eficacia, fornecedor_id,
                    fornecedor_nome, material_id, lote_fornecedor, ne_numero, guia, fatura, decisao,
                    movement_id, qtd_recebida, qtd_aprovada, qtd_rejeitada, qtd_pendente,
                    created_at, updated_at, created_by, updated_by, closed_at, closed_by
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
                """,
                (
                    _clip(row.get("id"), 30),
                    _clip(row.get("origem"), 120),
                    _clip(row.get("referencia"), 120),
                    _clip(row.get("entidade_tipo"), 60),
                    _clip(row.get("entidade_id"), 120),
                    _clip(row.get("entidade_label"), 255),
                    _clip(row.get("tipo"), 60),
                    _clip(row.get("gravidade"), 40),
                    _clip(row.get("estado"), 40),
                    _clip(row.get("responsavel"), 120),
                    _to_mysql_date(row.get("prazo")),
                    row.get("descricao"),
                    row.get("causa"),
                    row.get("acao"),
                    row.get("eficacia"),
                    _clip(row.get("fornecedor_id"), 30),
                    _clip(row.get("fornecedor_nome"), 150),
                    _clip(row.get("material_id"), 30),
                    _clip(row.get("lote_fornecedor"), 100),
                    _clip(row.get("ne_numero"), 30),
                    _clip(row.get("guia"), 60),
                    _clip(row.get("fatura"), 60),
                    _clip(row.get("decisao"), 255),
                    _clip(row.get("movement_id"), 255),
                    _to_num(row.get("qtd_recebida")),
                    _to_num(row.get("qtd_aprovada")),
                    _to_num(row.get("qtd_rejeitada")),
                    _to_num(row.get("qtd_pendente")),
                    _to_mysql_datetime(row.get("created_at")),
                    _to_mysql_datetime(row.get("updated_at")),
                    _clip(row.get("created_by"), 120),
                    _clip(row.get("updated_by"), 120),
                    _to_mysql_datetime(row.get("closed_at")),
                    _clip(row.get("closed_by"), 120),
                ),
            )

    if "quality_documents" in tables:
        for row in data.get("quality_documents", []):
            if not isinstance(row, dict) or not str(row.get("id", "") or "").strip():
                continue
            cur.execute(
                """
                INSERT INTO quality_documents (
                    id, titulo, tipo, entidade, referencia, entidade_tipo, entidade_id, versao,
                    estado, responsavel, caminho, obs, created_at, updated_at, created_by, updated_by
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    _clip(row.get("id"), 30),
                    _clip(row.get("titulo"), 180),
                    _clip(row.get("tipo"), 80),
                    _clip(row.get("entidade"), 80),
                    _clip(row.get("referencia"), 120),
                    _clip(row.get("entidade_tipo"), 60),
                    _clip(row.get("entidade_id"), 120),
                    _clip(row.get("versao"), 30),
                    _clip(row.get("estado"), 40),
                    _clip(row.get("responsavel"), 120),
                    _clip(row.get("caminho"), 512),
                    row.get("obs"),
                    _to_mysql_datetime(row.get("created_at")),
                    _to_mysql_datetime(row.get("updated_at")),
                    _clip(row.get("created_by"), 120),
                    _clip(row.get("updated_by"), 120),
                ),
            )

    if "quality_audit_log" in tables:
        for row in data.get("audit_log", []):
            if not isinstance(row, dict) or not str(row.get("id", "") or "").strip():
                continue
            cur.execute(
                """
                INSERT INTO quality_audit_log (
                    id, created_at, user_name, action, entity_type, entity_id, summary, before_json, after_json
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    _clip(row.get("id"), 80),
                    _to_mysql_datetime(row.get("created_at")),
                    _clip(row.get("user"), 120),
                    _clip(row.get("action"), 120),
                    _clip(row.get("entity_type"), 80),
                    _clip(row.get("entity_id"), 120),
                    row.get("summary"),
                    json.dumps(row.get("before"), ensure_ascii=False, default=str) if "before" in row else None,
                    json.dumps(row.get("after"), ensure_ascii=False, default=str) if "after" in row else None,
                ),
            )


def _db_to_iso(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%dT%H:%M:%S")
    txt = str(value).strip()
    if not txt:
        return ""
    try:
        return datetime.fromisoformat(txt.replace("Z", "+00:00")).strftime("%Y-%m-%dT%H:%M:%S")
    except Exception:
        return txt


def _next_seq_from_pattern(values, pattern, default_next=1):
    max_n = 0
    rx = re.compile(pattern)
    for v in values:
        m = rx.match(str(v or ""))
        if m:
            try:
                max_n = max(max_n, int(m.group(1)))
            except Exception:
                pass
    return max(max_n + 1, default_next)


def _rebuild_runtime_sequences(data):
    seq = data.setdefault("seq", {})
    seq.setdefault("ref_interna", {})

    seq["cliente"] = _next_seq_from_pattern(
        [c.get("codigo") for c in data.get("clientes", [])],
        r"^CL(\d{4,})$",
        1,
    )
    seq["produto"] = _next_seq_from_pattern(
        [p.get("codigo") for p in data.get("produtos", [])],
        r"^PRD-(\d{4,})$",
        1,
    )
    seq["fornecedor"] = max(
        _next_seq_from_pattern(
            [f.get("id") for f in data.get("fornecedores", [])],
            r"^FOR-(\d{4,})$",
            1,
        ),
        _load_fornecedor_sequence_next(data),
    )
    seq["encomenda"] = _next_seq_from_pattern(
        [e.get("numero") for e in data.get("encomendas", [])],
        r"^BARCELBAL(\d{4,})$",
        1,
    )
    seq["ne"] = _next_seq_from_pattern(
        [n.get("numero") for n in data.get("notas_encomenda", [])],
        r"^NE-\d{4}-(\d{4,})$",
        1,
    )
    data["exp_seq"] = _next_seq_from_pattern(
        [x.get("numero") for x in data.get("expedicoes", []) if isinstance(x, dict)],
        r"^GT-\d{4}-(\d{4,})$",
        1,
    )

    ref_map = {}
    for e in data.get("encomendas", []):
        cli = _extract_cliente_codigo(e.get("cliente"), data) or str(e.get("cliente", "")).strip()
        if not cli:
            continue
        for p in encomenda_pecas(e):
            ref = str(p.get("ref_interna", "")).strip()
            m = re.match(rf"^{re.escape(cli)}-(\d+)$", ref)
            if m:
                ref_map[cli] = max(ref_map.get(cli, 0), int(m.group(1)))
    for o in data.get("orcamentos", []):
        if not isinstance(o, dict):
            continue
        cli = _extract_cliente_codigo(o.get("cliente"), data) or str(o.get("cliente", "")).strip()
        if not cli:
            continue
        for l in o.get("linhas", []):
            if not isinstance(l, dict):
                continue
            ref = str(l.get("ref_interna", "")).strip()
            m = re.match(rf"^{re.escape(cli)}-(\d+)$", ref)
            if m:
                ref_map[cli] = max(ref_map.get(cli, 0), int(m.group(1)))
    seq["ref_interna"] = ref_map

    data["orc_seq"] = _next_seq_from_pattern(
        [o.get("numero") for o in data.get("orcamentos", []) if isinstance(o, dict)],
        r"^ORC-\d{4}-(\d{4,})$",
        1,
    )

    of_values = []
    opp_values = []
    for e in data.get("encomendas", []):
        for p in encomenda_pecas(e):
            of_values.append(p.get("of"))
            opp_values.append(p.get("opp"))
    for o in data.get("orcamentos", []):
        if not isinstance(o, dict):
            continue
        for l in o.get("linhas", []):
            if not isinstance(l, dict):
                continue
            of_values.append(l.get("of"))
            opp_values.append(l.get("opp"))
    data["of_seq"] = _next_seq_from_pattern(of_values, r"^OF-\d{4}-(\d{4,})$", 1)
    data["opp_seq"] = _next_seq_from_pattern(opp_values, r"^OPP-\d{4}-(\d{4,})$", 1)

    refs = []
    seen = set()
    for e in data.get("encomendas", []):
        for p in encomenda_pecas(e):
            for ref in (p.get("ref_interna"), p.get("ref_externa")):
                r = str(ref or "").strip()
                if r and r not in seen:
                    seen.add(r)
                    refs.append(r)
    for o in data.get("orcamentos", []):
        if not isinstance(o, dict):
            continue
        for l in o.get("linhas", []):
            if not isinstance(l, dict):
                continue
            for ref in (l.get("ref_interna"), l.get("ref_externa")):
                r = str(ref or "").strip()
                if r and r not in seen:
                    seen.add(r)
                    refs.append(r)
    data["refs"] = refs
    data["materiais_hist"] = sorted({str(m.get("material", "")).strip() for m in data.get("materiais", []) if str(m.get("material", "")).strip()})
    data["espessuras_hist"] = sorted({str(m.get("espessura", "")).strip() for m in data.get("materiais", []) if str(m.get("espessura", "")).strip()})


def _mysql_load_relational_data():
    conn = _mysql_connect()
    try:
        data = _copy_default_data()
        with conn.cursor() as cur:
            tables = _mysql_existing_tables(cur)

            def fetch_all(table, order_by=None):
                if table not in tables:
                    return []
                sql = f"SELECT * FROM `{table}`"
                if order_by:
                    sql += f" ORDER BY {order_by}"
                cur.execute(sql)
                return cur.fetchall() or []

            users_rows = fetch_all("users", "id")
            if users_rows:
                data["users"] = [
                    {
                        "username": str(r.get("username", "") or ""),
                        "password": str(r.get("password", "") or ""),
                        "role": str(r.get("role", "") or ""),
                    }
                    for r in users_rows
                ]
            ops_rows = fetch_all("operadores", "nome")
            if ops_rows:
                data["operadores"] = [str(r.get("nome", "") or "").strip() for r in ops_rows if str(r.get("nome", "") or "").strip()]
            orc_exec_rows = fetch_all("orcamentistas", "nome")
            if orc_exec_rows:
                data["orcamentistas"] = [str(r.get("nome", "") or "").strip() for r in orc_exec_rows if str(r.get("nome", "") or "").strip()]

            for r in fetch_all("at_series", "id"):
                doc_type = str(r.get("doc_type", "") or "GT").strip().upper() or "GT"
                serie_id = str(r.get("serie_id", "") or "").strip()
                if not serie_id:
                    continue
                inicio_seq = int(_to_num(r.get("inicio_sequencia")) or 1)
                next_seq = int(_to_num(r.get("next_seq")) or inicio_seq or 1)
                if inicio_seq < 1:
                    inicio_seq = 1
                if next_seq < inicio_seq:
                    next_seq = inicio_seq
                data["at_series"].append(
                    {
                        "doc_type": doc_type,
                        "serie_id": serie_id,
                        "inicio_sequencia": inicio_seq,
                        "next_seq": next_seq,
                        "data_inicio_prevista": _db_to_iso(r.get("data_inicio_prevista")),
                        "validation_code": str(r.get("validation_code", "") or ""),
                        "status": str(r.get("status", "") or ""),
                        "last_error": str(r.get("last_error", "") or ""),
                        "last_sent_payload_hash": str(r.get("last_sent_payload_hash", "") or ""),
                        "updated_at": _db_to_iso(r.get("updated_at")),
                    }
                )

            for r in fetch_all("clientes", "codigo"):
                data["clientes"].append(
                    {
                        "codigo": str(r.get("codigo", "") or ""),
                        "nome": str(r.get("nome", "") or ""),
                        "nif": str(r.get("nif", "") or ""),
                        "morada": str(r.get("morada", "") or ""),
                        "contacto": str(r.get("contacto", "") or ""),
                        "email": str(r.get("email", "") or ""),
                        "observacoes": "",
                        "prazo_entrega": "",
                        "cond_pagamento": "",
                        "obs_tecnicas": "",
                    }
                )

            for r in fetch_all("fornecedores", "id"):
                data["fornecedores"].append(
                    {
                        "id": str(r.get("id", "") or ""),
                        "nome": str(r.get("nome", "") or ""),
                        "nif": str(r.get("nif", "") or ""),
                        "morada": str(r.get("morada", "") or ""),
                        "contacto": str(r.get("contacto", "") or ""),
                        "email": str(r.get("email", "") or ""),
                        "codigo_postal": str(r.get("codigo_postal", "") or ""),
                        "localidade": str(r.get("localidade", "") or ""),
                        "pais": str(r.get("pais", "") or ""),
                        "cond_pagamento": str(r.get("cond_pagamento", "") or ""),
                        "prazo_entrega_dias": str(r.get("prazo_entrega_dias", "") or ""),
                        "website": str(r.get("website", "") or ""),
                        "obs": str(r.get("obs", "") or ""),
                    }
                )

            for r in fetch_all("materiais", "id"):
                local = str(r.get("localizacao", "") or "")
                data["materiais"].append(
                    {
                        "id": str(r.get("id", "") or ""),
                        "lote_fornecedor": str(r.get("lote_fornecedor", "") or ""),
                        "formato": str(r.get("formato", "") or ""),
                        "material": str(r.get("material", "") or ""),
                        "material_familia": str(r.get("material_familia", "") or ""),
                        "espessura": str(r.get("espessura", "") or ""),
                        "comprimento": _to_num(r.get("comprimento")) or 0.0,
                        "largura": _to_num(r.get("largura")) or 0.0,
                        "altura": _to_num(r.get("altura")) or 0.0,
                        "diametro": _to_num(r.get("diametro")) or 0.0,
                        "metros": _to_num(r.get("metros")) or 0.0,
                        "kg_m": _to_num(r.get("kg_m")) or 0.0,
                        "peso_unid": _to_num(r.get("peso_unid")) or 0.0,
                        "p_compra": _to_num(r.get("p_compra")) or 0.0,
                        "preco_unid": _to_num(r.get("preco_unid")) or 0.0,
                        "quantidade": _to_num(r.get("quantidade")) or 0.0,
                        "reservado": _to_num(r.get("reservado")) or 0.0,
                        "tipo": str(r.get("tipo", "") or ""),
                        "secao_tipo": str(r.get("secao_tipo", "") or ""),
                        "Localizacao": local,
                        "Localizacao": local,
                        "is_sobra": bool(r.get("is_sobra")),
                        "atualizado_em": _db_to_iso(r.get("atualizado_em")),
                        "origem_lote": str(r.get("origem_lote", "") or ""),
                        "origem_encomenda": str(r.get("origem_encomenda", "") or ""),
                        "logistic_status": str(r.get("logistic_status", "") or ""),
                        "quality_status": str(r.get("quality_status", "") or ""),
                        "quality_blocked": bool(r.get("quality_blocked")),
                        "quality_pending_qty": _to_num(r.get("quality_pending_qty")) or 0.0,
                        "quality_received_qty": _to_num(r.get("quality_received_qty")) or 0.0,
                        "quality_approved_qty": _to_num(r.get("quality_approved_qty")) or 0.0,
                        "quality_return_document_id": str(r.get("quality_return_document_id", "") or ""),
                        "inspection_status": str(r.get("inspection_status", "") or ""),
                        "inspection_defect": str(r.get("inspection_defect", "") or ""),
                        "inspection_decision": str(r.get("inspection_decision", "") or ""),
                        "inspection_at": _db_to_iso(r.get("inspection_at")),
                        "inspection_by": str(r.get("inspection_by", "") or ""),
                        "inspection_note_number": str(r.get("inspection_note_number", "") or ""),
                        "inspection_supplier_id": str(r.get("inspection_supplier_id", "") or ""),
                        "inspection_supplier_name": str(r.get("inspection_supplier_name", "") or ""),
                        "inspection_guia": str(r.get("inspection_guia", "") or ""),
                        "inspection_fatura": str(r.get("inspection_fatura", "") or ""),
                        "quality_nc_id": str(r.get("quality_nc_id", "") or ""),
                        "supplier_claim_id": str(r.get("supplier_claim_id", "") or ""),
                    }
                )

            for r in fetch_all("produtos", "codigo"):
                data["produtos"].append(
                    {
                        "codigo": str(r.get("codigo", "") or ""),
                        "descricao": str(r.get("descricao", "") or ""),
                        "categoria": str(r.get("categoria", "") or ""),
                        "category_id": str(r.get("category_id", "") or ""),
                        "subcat": str(r.get("subcat", "") or ""),
                        "subcategory_id": str(r.get("subcategory_id", "") or ""),
                        "tipo": str(r.get("tipo", "") or ""),
                        "type_id": str(r.get("type_id", "") or ""),
                        "dimensoes": "",
                        "comprimento": 0.0,
                        "largura": 0.0,
                        "espessura": 0.0,
                        "metros_unidade": 0.0,
                        "metros": 0.0,
                        "peso_unid": 0.0,
                        "fabricante": "",
                        "modelo": "",
                        "unid": str(r.get("unid", "") or "UN"),
                        "qty": _to_num(r.get("qty")) or 0.0,
                        "alerta": _to_num(r.get("alerta")) or 0.0,
                        "p_compra": _to_num(r.get("p_compra")) or 0.0,
                        "pvp1": 0.0,
                        "pvp2": 0.0,
                        "obs": "",
                        "atualizado_em": _db_to_iso(r.get("atualizado_em")),
                        "logistic_status": str(r.get("logistic_status", "") or ""),
                        "quality_status": str(r.get("quality_status", "") or ""),
                        "quality_blocked": bool(r.get("quality_blocked")),
                        "quality_pending_qty": _to_num(r.get("quality_pending_qty")) or 0.0,
                        "quality_received_qty": _to_num(r.get("quality_received_qty")) or 0.0,
                        "quality_approved_qty": _to_num(r.get("quality_approved_qty")) or 0.0,
                        "quality_return_document_id": str(r.get("quality_return_document_id", "") or ""),
                        "inspection_defect": str(r.get("inspection_defect", "") or ""),
                        "inspection_decision": str(r.get("inspection_decision", "") or ""),
                        "inspection_note_number": str(r.get("inspection_note_number", "") or ""),
                        "quality_nc_id": str(r.get("quality_nc_id", "") or ""),
                    }
                )

            modelos_map = {}
            for row in fetch_all("conjuntos_modelo", "codigo"):
                codigo = str(row.get("codigo", "") or "").strip()
                if not codigo:
                    continue
                modelos_map[codigo] = {
                    "codigo": codigo,
                    "descricao": str(row.get("descricao", "") or "").strip() or codigo,
                    "notas": str(row.get("notas", "") or "").strip(),
                    "ativo": bool(row.get("ativo")) if row.get("ativo") is not None else True,
                    "template": bool(row.get("template")) if row.get("template") is not None else False,
                    "origem": str(row.get("origem", "") or "").strip(),
                    "created_at": _db_to_iso(row.get("created_at")),
                    "updated_at": _db_to_iso(row.get("updated_at")),
                    "itens": [],
                }
            for row in fetch_all("conjuntos_modelo_itens", "conjunto_codigo, linha_ordem, id"):
                codigo = str(row.get("conjunto_codigo", "") or "").strip()
                if not codigo:
                    continue
                meta_payload = {}
                try:
                    raw_meta_json = row.get("meta_json")
                    if raw_meta_json:
                        parsed_meta = json.loads(raw_meta_json)
                        if isinstance(parsed_meta, dict):
                            meta_payload = parsed_meta
                except Exception:
                    meta_payload = {}
                modelos_map.setdefault(
                    codigo,
                    {
                        "codigo": codigo,
                        "descricao": codigo,
                        "notas": "",
                        "ativo": True,
                        "template": False,
                        "origem": "",
                        "created_at": "",
                        "updated_at": "",
                        "itens": [],
                    },
                )
                modelos_map[codigo]["itens"].append(
                    {
                        "linha_ordem": int(_to_num(row.get("linha_ordem")) or len(modelos_map[codigo]["itens"]) + 1),
                        "tipo_item": normalize_orc_line_type(row.get("tipo_item")),
                        "ref_externa": str(row.get("ref_externa", "") or ""),
                        "descricao": str(row.get("descricao", "") or ""),
                        "material": str(row.get("material", "") or ""),
                        "espessura": str(row.get("espessura", "") or ""),
                        "operacao": str(row.get("operacao", "") or ""),
                        "produto_codigo": str(row.get("produto_codigo", "") or ""),
                        "produto_unid": str(row.get("produto_unid", "") or ""),
                        "qtd": _to_num(row.get("qtd")) or 0.0,
                        "tempo_peca_min": _to_num(row.get("tempo_peca_min")) or 0.0,
                        "preco_unit": _to_num(row.get("preco_unit")) or 0.0,
                        "desenho": str(row.get("desenho_path", "") or ""),
                        **dict(meta_payload or {}),
                    }
                )
            if modelos_map:
                data["conjuntos_modelo"] = list(modelos_map.values())

            conjuntos_map = {}
            for row in fetch_all("conjuntos", "codigo"):
                codigo = str(row.get("codigo", "") or "").strip()
                if not codigo:
                    continue
                conjuntos_map[codigo] = {
                    "codigo": codigo,
                    "descricao": str(row.get("descricao", "") or "").strip() or codigo,
                    "notas": str(row.get("notas", "") or "").strip(),
                    "ativo": bool(row.get("ativo")) if row.get("ativo") is not None else True,
                    "template": bool(row.get("template")) if row.get("template") is not None else False,
                    "origem": str(row.get("origem", "") or "").strip(),
                    "margem_perc": _to_num(row.get("margem_perc")) or 0.0,
                    "total_custo": _to_num(row.get("total_custo")) or 0.0,
                    "total_final": _to_num(row.get("total_final")) or 0.0,
                    "created_at": _db_to_iso(row.get("created_at")),
                    "updated_at": _db_to_iso(row.get("updated_at")),
                    "itens": [],
                }
            for row in fetch_all("conjuntos_itens", "conjunto_codigo, linha_ordem, id"):
                codigo = str(row.get("conjunto_codigo", "") or "").strip()
                if not codigo:
                    continue
                meta_payload = {}
                try:
                    raw_meta_json = row.get("meta_json")
                    if raw_meta_json:
                        parsed_meta = json.loads(raw_meta_json)
                        if isinstance(parsed_meta, dict):
                            meta_payload = parsed_meta
                except Exception:
                    meta_payload = {}
                conjuntos_map.setdefault(
                    codigo,
                    {
                        "codigo": codigo,
                        "descricao": codigo,
                        "notas": "",
                        "ativo": True,
                        "template": False,
                        "origem": "",
                        "margem_perc": 0.0,
                        "total_custo": 0.0,
                        "total_final": 0.0,
                        "created_at": "",
                        "updated_at": "",
                        "itens": [],
                    },
                )
                conjuntos_map[codigo]["itens"].append(
                    {
                        "linha_ordem": int(_to_num(row.get("linha_ordem")) or len(conjuntos_map[codigo]["itens"]) + 1),
                        "tipo_item": normalize_orc_line_type(row.get("tipo_item")),
                        "ref_externa": str(row.get("ref_externa", "") or ""),
                        "descricao": str(row.get("descricao", "") or ""),
                        "material": str(row.get("material", "") or ""),
                        "espessura": str(row.get("espessura", "") or ""),
                        "operacao": str(row.get("operacao", "") or ""),
                        "produto_codigo": str(row.get("produto_codigo", "") or ""),
                        "produto_unid": str(row.get("produto_unid", "") or ""),
                        "qtd": _to_num(row.get("qtd")) or 0.0,
                        "tempo_peca_min": _to_num(row.get("tempo_peca_min")) or 0.0,
                        "preco_unit": _to_num(row.get("preco_unit")) or 0.0,
                        "desenho": str(row.get("desenho_path", "") or ""),
                        **dict(meta_payload or {}),
                    }
                )
            if conjuntos_map:
                data["conjuntos"] = list(conjuntos_map.values())

            orc_linhas_map = {}
            for l in fetch_all("orcamento_linhas", "id"):
                num = str(l.get("orcamento_numero", "") or "")
                operacoes_json = {}
                tempos_operacao = {}
                custos_operacao = {}
                quote_cost_snapshot = {}
                try:
                    raw_operacoes_json = l.get("operacoes_json")
                    if raw_operacoes_json:
                        parsed_operacoes = json.loads(raw_operacoes_json)
                        if isinstance(parsed_operacoes, dict):
                            operacoes_json = parsed_operacoes
                except Exception:
                    operacoes_json = {}
                try:
                    raw_tempos_json = l.get("tempos_operacao_json")
                    if raw_tempos_json:
                        parsed_tempos = json.loads(raw_tempos_json)
                        if isinstance(parsed_tempos, dict):
                            tempos_operacao = parsed_tempos
                except Exception:
                    tempos_operacao = {}
                try:
                    raw_custos_json = l.get("custos_operacao_json")
                    if raw_custos_json:
                        parsed_custos = json.loads(raw_custos_json)
                        if isinstance(parsed_custos, dict):
                            custos_operacao = parsed_custos
                except Exception:
                    custos_operacao = {}
                try:
                    raw_snapshot_json = l.get("quote_cost_snapshot_json")
                    if raw_snapshot_json:
                        parsed_snapshot = json.loads(raw_snapshot_json)
                        if isinstance(parsed_snapshot, dict):
                            quote_cost_snapshot = parsed_snapshot
                except Exception:
                    quote_cost_snapshot = {}
                meta_payload = {}
                try:
                    raw_meta_json = l.get("meta_json")
                    if raw_meta_json:
                        parsed_meta = json.loads(raw_meta_json)
                        if isinstance(parsed_meta, dict):
                            meta_payload = parsed_meta
                except Exception:
                    meta_payload = {}
                orc_linhas_map.setdefault(num, []).append(
                    {
                        "ref_interna": str(l.get("ref_interna", "") or ""),
                        "ref_externa": str(l.get("ref_externa", "") or ""),
                        "descricao": str(l.get("descricao", "") or ""),
                        "tipo_item": normalize_orc_line_type(l.get("tipo_item")),
                        "material": str(l.get("material", "") or ""),
                        "espessura": str(l.get("espessura", "") or ""),
                        "operacao": str(l.get("operacao", "") or ""),
                        "of": str(l.get("of_codigo", "") or ""),
                        "produto_codigo": str(l.get("produto_codigo", "") or ""),
                        "produto_unid": str(l.get("produto_unid", "") or ""),
                        "conjunto_codigo": str(l.get("conjunto_codigo", "") or ""),
                        "conjunto_nome": str(l.get("conjunto_nome", "") or ""),
                        "grupo_uuid": str(l.get("grupo_uuid", "") or ""),
                        "qtd_base": _to_num(l.get("qtd_base")) or (_to_num(l.get("qtd")) or 0.0),
                        "tempo_peca_min": _to_num(l.get("tempo_peca_min")) or 0.0,
                        "qtd": _to_num(l.get("qtd")) or 0.0,
                        "preco_unit": _to_num(l.get("preco_unit")) or 0.0,
                        "total": _to_num(l.get("total")) or 0.0,
                        "desenho": str(l.get("desenho_path", "") or ""),
                        "operacoes_lista": list(operacoes_json.get("operacoes_lista", []) or []),
                        "operacoes_fluxo": [dict(item or {}) for item in list(operacoes_json.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                        "operacoes_detalhe": [dict(item or {}) for item in list(operacoes_json.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                        "tempos_operacao": dict(tempos_operacao or {}),
                        "custos_operacao": dict(custos_operacao or {}),
                        "quote_cost_snapshot": dict(quote_cost_snapshot or {}),
                        **dict(meta_payload or {}),
                    }
                )

            for o in fetch_all("orcamentos", "numero"):
                num = str(o.get("numero", "") or "")
                meta_payload = {}
                try:
                    raw_meta_json = o.get("meta_json")
                    if raw_meta_json:
                        parsed_meta = json.loads(raw_meta_json)
                        if isinstance(parsed_meta, dict):
                            meta_payload = parsed_meta
                except Exception:
                    meta_payload = {}
                data["orcamentos"].append(
                    {
                        "numero": num,
                        "data": _db_to_iso(o.get("data")),
                        "estado": str(o.get("estado", "") or ""),
                        "cliente": _normalize_orc_cliente(str(o.get("cliente_codigo", "") or ""), data),
                        "posto_trabalho": str(o.get("posto_trabalho", "") or ""),
                        "linhas": orc_linhas_map.get(num, []),
                        "iva_perc": _to_num(o.get("iva_perc")) or 0.0,
                        "desconto_perc": _to_num(o.get("desconto_perc")) or 0.0,
                        "desconto_valor": _to_num(o.get("desconto_valor")) or 0.0,
                        "subtotal_bruto": _to_num(o.get("subtotal_bruto")) or 0.0,
                        "preco_transporte": _to_num(o.get("preco_transporte")) or 0.0,
                        "custo_transporte": _to_num(o.get("custo_transporte")) or 0.0,
                        "paletes": _to_num(o.get("paletes")) or 0.0,
                        "peso_bruto_kg": _to_num(o.get("peso_bruto_kg")) or 0.0,
                        "volume_m3": _to_num(o.get("volume_m3")) or 0.0,
                        "transportadora_id": str(o.get("transportadora_id", "") or ""),
                        "transportadora_nome": str(o.get("transportadora_nome", "") or ""),
                        "referencia_transporte": str(o.get("referencia_transporte", "") or ""),
                        "zona_transporte": str(o.get("zona_transporte", "") or ""),
                        "subtotal": _to_num(o.get("subtotal")) or 0.0,
                        "total": _to_num(o.get("total")) or 0.0,
                        "numero_encomenda": str(o.get("numero_encomenda", "") or ""),
                        "executado_por": str(o.get("executado_por", "") or ""),
                        "nota_transporte": str(o.get("nota_transporte", "") or ""),
                        "notas_pdf": str(o.get("notas_pdf", "") or ""),
                        "nota_cliente": str(o.get("nota_cliente", "") or ""),
                        **dict(meta_payload or {}),
                    }
                )

            refs_rows = fetch_all("orc_referencias_historico", "updated_at DESC")
            if refs_rows:
                refs_db = {}
                for r in refs_rows:
                    ref_ext = str(r.get("ref_externa", "") or "").strip()
                    if not ref_ext or ref_ext in refs_db:
                        continue
                    operacoes_json = {}
                    tempos_operacao = {}
                    custos_operacao = {}
                    quote_cost_snapshot = {}
                    try:
                        raw_operacoes_json = r.get("operacoes_json")
                        if raw_operacoes_json:
                            parsed_operacoes = json.loads(raw_operacoes_json)
                            if isinstance(parsed_operacoes, dict):
                                operacoes_json = parsed_operacoes
                    except Exception:
                        operacoes_json = {}
                    try:
                        raw_tempos_json = r.get("tempos_operacao_json")
                        if raw_tempos_json:
                            parsed_tempos = json.loads(raw_tempos_json)
                            if isinstance(parsed_tempos, dict):
                                tempos_operacao = parsed_tempos
                    except Exception:
                        tempos_operacao = {}
                    try:
                        raw_custos_json = r.get("custos_operacao_json")
                        if raw_custos_json:
                            parsed_custos = json.loads(raw_custos_json)
                            if isinstance(parsed_custos, dict):
                                custos_operacao = parsed_custos
                    except Exception:
                        custos_operacao = {}
                    try:
                        raw_snapshot_json = r.get("quote_cost_snapshot_json")
                        if raw_snapshot_json:
                            parsed_snapshot = json.loads(raw_snapshot_json)
                            if isinstance(parsed_snapshot, dict):
                                quote_cost_snapshot = parsed_snapshot
                    except Exception:
                        quote_cost_snapshot = {}
                    refs_db[ref_ext] = {
                        "ref_interna": str(r.get("ref_interna", "") or ""),
                        "ref_externa": ref_ext,
                        "descricao": str(r.get("descricao", "") or ""),
                        "material": str(r.get("material", "") or ""),
                        "espessura": str(r.get("espessura", "") or ""),
                        "preco_unit": _to_num(r.get("preco_unit")) or 0.0,
                        "operacao": str(r.get("operacao", "") or ""),
                        "tempo_peca_min": _to_num(r.get("tempo_peca_min")) or 0.0,
                        "operacoes_lista": list(operacoes_json.get("operacoes_lista", []) or []),
                        "operacoes_fluxo": [dict(item or {}) for item in list(operacoes_json.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                        "operacoes_detalhe": [dict(item or {}) for item in list(operacoes_json.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                        "tempos_operacao": dict(tempos_operacao or {}),
                        "custos_operacao": dict(custos_operacao or {}),
                        "quote_cost_snapshot": dict(quote_cost_snapshot or {}),
                        "origem_doc": str(r.get("origem_doc", "") or ""),
                        "origem_tipo": str(r.get("origem_tipo", "") or ""),
                        "estado_origem": str(r.get("estado_origem", "") or ""),
                        "approved_at": _db_to_iso(r.get("approved_at")),
                        "desenho": str(r.get("desenho_path", "") or ""),
                    }
                data["orc_refs"] = refs_db

            peca_map = {}
            for p in fetch_all("pecas", "id"):
                enc_num = str(p.get("encomenda_numero", "") or "")
                fluxo = []
                hist_rows = []
                try:
                    raw_fluxo = p.get("operacoes_fluxo_json")
                    if raw_fluxo:
                        parsed = json.loads(raw_fluxo)
                        if isinstance(parsed, list):
                            fluxo = parsed
                except Exception:
                    fluxo = []
                try:
                    raw_hist = p.get("hist_json")
                    if raw_hist:
                        parsed_h = json.loads(raw_hist)
                        if isinstance(parsed_h, list):
                            hist_rows = parsed_h
                except Exception:
                    hist_rows = []
                ficheiros_rows = []
                try:
                    raw_files = p.get("ficheiros_json")
                    if raw_files:
                        parsed_files = json.loads(raw_files)
                        if isinstance(parsed_files, list):
                            ficheiros_rows = [str(item or "") for item in parsed_files if str(item or "")]
                except Exception:
                    ficheiros_rows = []
                peca = {
                    "id": str(p.get("id", "") or ""),
                    "ref_interna": str(p.get("ref_interna", "") or ""),
                    "ref_externa": str(p.get("ref_externa", "") or ""),
                    "material": str(p.get("material", "") or ""),
                    "tipo_material": str(p.get("tipo_material", "") or "CHAPA"),
                    "subtipo_material": str(p.get("subtipo_material", "") or p.get("material", "") or ""),
                    "espessura": str(p.get("espessura", "") or ""),
                    "dimensao": str(p.get("dimensao", "") or ""),
                    "dimensoes": str(p.get("dimensao", "") or ""),
                    "perfil_tipo": str(p.get("perfil_tipo", "") or ""),
                    "perfil_tamanho": str(p.get("perfil_tamanho", "") or ""),
                    "comprimento_mm": _to_num(p.get("comprimento_mm")) or 0.0,
                    "tubo_forma": str(p.get("tubo_forma", "") or ""),
                    "lado_a": _to_num(p.get("lado_a")) or 0.0,
                    "lado_b": _to_num(p.get("lado_b")) or 0.0,
                    "tubo_espessura": _to_num(p.get("tubo_espessura")) or 0.0,
                    "diametro": _to_num(p.get("diametro")) or 0.0,
                    "quantidade_pedida": _to_num(p.get("quantidade_pedida")) or 0.0,
                    "Operacoes": str(p.get("operacoes", "") or ""),
                    "Observacoes": str(p.get("observacoes", "") or ""),
                    "of": str(p.get("of_codigo", "") or ""),
                    "opp": str(p.get("opp_codigo", "") or ""),
                    "estado": str(p.get("estado", "") or ""),
                    "produzido_ok": _to_num(p.get("produzido_ok")) or 0.0,
                    "produzido_nok": _to_num(p.get("produzido_nok")) or 0.0,
                    "inicio_producao": _db_to_iso(p.get("inicio_producao")),
                    "fim_producao": _db_to_iso(p.get("fim_producao")),
                    "produzido_qualidade": 0.0,
                    "tempo_producao_min": _to_num(p.get("tempo_producao_min")) or 0.0,
                    "hist": hist_rows,
                    "lote_baixa": str(p.get("lote_baixa", "") or ""),
                    "desenho": str(p.get("desenho_path", "") or ""),
                    "ficheiros": ficheiros_rows,
                    "operacoes_fluxo": fluxo,
                    "qtd_expedida": _to_num(p.get("qtd_expedida")) or 0.0,
                    "expedicoes": [],
                }
                peca_map.setdefault(enc_num, []).append(peca)

            esp_meta_map = {}
            for r in fetch_all("encomenda_espessuras", "id"):
                enc_num = str(r.get("encomenda_numero", "") or "")
                mat = str(r.get("material", "") or "").strip()
                esp = str(r.get("espessura", "") or "").strip()
                if not enc_num or not mat or not esp:
                    continue
                tempos_operacao = {}
                maquinas_operacao = {}
                raw_tempos_operacao = r.get("tempos_operacao_json")
                if raw_tempos_operacao:
                    try:
                        parsed_tempos = json.loads(str(raw_tempos_operacao))
                    except Exception:
                        parsed_tempos = {}
                    if isinstance(parsed_tempos, dict):
                        for op_name, op_value in parsed_tempos.items():
                            op_txt = normalize_planeamento_operacao(op_name)
                            if not op_txt:
                                continue
                            op_raw = str(op_value if op_value is not None else "").strip()
                            if op_raw:
                                tempos_operacao[op_txt] = op_raw
                raw_maquinas_operacao = r.get("maquinas_operacao_json")
                if raw_maquinas_operacao:
                    try:
                        parsed_maquinas = json.loads(str(raw_maquinas_operacao))
                    except Exception:
                        parsed_maquinas = {}
                    if isinstance(parsed_maquinas, dict):
                        for op_name, resource_value in parsed_maquinas.items():
                            op_txt = normalize_planeamento_operacao(op_name)
                            if not op_txt:
                                continue
                            resource_txt = str(resource_value if resource_value is not None else "").strip()
                            if resource_txt:
                                maquinas_operacao[op_txt] = resource_txt
                esp_meta_map.setdefault(enc_num, {})[(mat, esp)] = {
                    "tempo_min": _to_num(r.get("tempo_min")),
                    "tempos_operacao": tempos_operacao,
                    "maquinas_operacao": maquinas_operacao,
                    "estado": str(r.get("estado", "") or "Preparacao"),
                    "inicio_producao": _db_to_iso(r.get("inicio_producao")),
                    "fim_producao": _db_to_iso(r.get("fim_producao")),
                    "tempo_producao_min": _to_num(r.get("tempo_producao_min")) or 0.0,
                    "lote_baixa": str(r.get("lote_baixa", "") or ""),
                }

            reservas_map = {}
            for r in fetch_all("encomenda_reservas", "id"):
                enc_num = str(r.get("encomenda_numero", "") or "").strip()
                mat = str(r.get("material", "") or "").strip()
                esp = str(r.get("espessura", "") or "").strip()
                qtd = _to_num(r.get("quantidade")) or 0.0
                if not enc_num or not mat or not esp or qtd <= 0:
                    continue
                reservas_map.setdefault(enc_num, []).append(
                    {
                        "material_id": str(r.get("material_id", "") or ""),
                        "material": mat,
                        "espessura": esp,
                        "quantidade": qtd,
                    }
                )

            montagem_map = {}
            for r in fetch_all("encomenda_montagem_itens", "encomenda_numero, linha_ordem, id"):
                enc_num = str(r.get("encomenda_numero", "") or "").strip()
                if not enc_num:
                    continue
                montagem_map.setdefault(enc_num, []).append(
                    {
                        "linha_ordem": int(_to_num(r.get("linha_ordem")) or len(montagem_map.get(enc_num, [])) + 1),
                        "tipo_item": normalize_orc_line_type(r.get("tipo_item")),
                        "descricao": str(r.get("descricao", "") or ""),
                        "produto_codigo": str(r.get("produto_codigo", "") or ""),
                        "produto_unid": str(r.get("produto_unid", "") or ""),
                        "qtd_planeada": _to_num(r.get("qtd_planeada")) or 0.0,
                        "qtd_consumida": _to_num(r.get("qtd_consumida")) or 0.0,
                        "preco_unit": _to_num(r.get("preco_unit")) or 0.0,
                        "conjunto_codigo": str(r.get("conjunto_codigo", "") or ""),
                        "conjunto_nome": str(r.get("conjunto_nome", "") or ""),
                        "grupo_uuid": str(r.get("grupo_uuid", "") or ""),
                        "estado": str(r.get("estado", "") or ""),
                        "obs": str(r.get("obs", "") or ""),
                        "created_at": _db_to_iso(r.get("created_at")),
                        "consumed_at": _db_to_iso(r.get("consumed_at")),
                        "consumed_by": str(r.get("consumed_by", "") or ""),
                    }
                )

            for p in fetch_all("plano", "data_planeada, inicio, id"):
                enc_num = str(p.get("encomenda_numero", "") or "").strip()
                data_planeada = _db_to_iso(p.get("data_planeada"))[:10]
                inicio_txt = str(p.get("inicio", "") or "").strip()
                if len(inicio_txt) >= 5:
                    inicio_txt = inicio_txt[:5]
                duracao_num = _to_num(p.get("duracao_min"))
                if not enc_num or not data_planeada or not inicio_txt or duracao_num is None:
                    continue
                data["plano"].append(
                    {
                        "id": str(p.get("bloco_id", "") or ""),
                        "encomenda": enc_num,
                        "material": str(p.get("material", "") or ""),
                        "espessura": str(p.get("espessura", "") or ""),
                        "operacao": str(p.get("operacao", "") or ""),
                        "posto": str(p.get("posto", "") or ""),
                        "data": data_planeada,
                        "inicio": inicio_txt,
                        "duracao_min": int(round(duracao_num)),
                        "color": str(p.get("color", "") or ""),
                        "chapa": str(p.get("chapa", "") or ""),
                    }
                )

            for p in fetch_all("plano_hist", "id"):
                enc_num = str(p.get("encomenda_numero", "") or "").strip()
                data_planeada = _db_to_iso(p.get("data_planeada"))[:10]
                inicio_txt = str(p.get("inicio", "") or "").strip()
                if len(inicio_txt) >= 5:
                    inicio_txt = inicio_txt[:5]
                duracao_num = _to_num(p.get("duracao_min"))
                if not enc_num or not data_planeada or not inicio_txt or duracao_num is None:
                    continue
                data["plano_hist"].append(
                    {
                        "id": str(p.get("bloco_id", "") or ""),
                        "encomenda": enc_num,
                        "material": str(p.get("material", "") or ""),
                        "espessura": str(p.get("espessura", "") or ""),
                        "operacao": str(p.get("operacao", "") or ""),
                        "posto": str(p.get("posto", "") or ""),
                        "data": data_planeada,
                        "inicio": inicio_txt,
                        "duracao_min": int(round(duracao_num)),
                        "color": str(p.get("color", "") or ""),
                        "chapa": str(p.get("chapa", "") or ""),
                        "movido_em": _db_to_iso(p.get("movido_em")),
                        "estado_final": str(p.get("estado_final", "") or ""),
                        "tempo_planeado_min": _to_num(p.get("tempo_planeado_min")) or 0.0,
                        "tempo_real_min": _to_num(p.get("tempo_real_min")) or 0.0,
                    }
                )

            for e in fetch_all("encomendas", "numero"):
                num = str(e.get("numero", "") or "")
                mats = {}
                for p in peca_map.get(num, []):
                    mat = str(p.get("material", "") or "")
                    esp = str(p.get("espessura", "") or "")
                    mats.setdefault(mat, {"material": mat, "estado": "Preparacao", "esp_map": {}})
                    mats[mat]["esp_map"].setdefault(
                        esp,
                        {
                            "espessura": esp,
                            "tempo_min": "",
                            "tempos_operacao": {},
                            "estado": "Preparacao",
                            "pecas": [],
                            "inicio_producao": "",
                            "fim_producao": "",
                            "tempo_producao_min": 0.0,
                            "lote_baixa": "",
                        },
                    )
                    mats[mat]["esp_map"][esp]["pecas"].append(p)
                for (mat, esp), meta in esp_meta_map.get(num, {}).items():
                    mats.setdefault(mat, {"material": mat, "estado": "Preparacao", "esp_map": {}})
                    mats[mat]["esp_map"].setdefault(
                        esp,
                        {
                            "espessura": esp,
                            "tempo_min": "",
                            "tempos_operacao": {},
                            "estado": "Preparacao",
                            "pecas": [],
                            "inicio_producao": "",
                            "fim_producao": "",
                            "tempo_producao_min": 0.0,
                            "lote_baixa": "",
                        },
                    )
                    slot = mats[mat]["esp_map"][esp]
                    if meta.get("tempo_min") is not None:
                        slot["tempo_min"] = meta.get("tempo_min")
                    slot["tempos_operacao"] = dict(meta.get("tempos_operacao", {}) or slot.get("tempos_operacao", {}))
                    slot["maquinas_operacao"] = dict(meta.get("maquinas_operacao", {}) or slot.get("maquinas_operacao", slot.get("recursos_operacao", {})))
                    slot["estado"] = meta.get("estado") or slot.get("estado", "Preparacao")
                    slot["inicio_producao"] = meta.get("inicio_producao", slot.get("inicio_producao", ""))
                    slot["fim_producao"] = meta.get("fim_producao", slot.get("fim_producao", ""))
                    slot["tempo_producao_min"] = meta.get("tempo_producao_min", slot.get("tempo_producao_min", 0.0))
                    slot["lote_baixa"] = meta.get("lote_baixa", slot.get("lote_baixa", ""))
                materiais = []
                for m in mats.values():
                    m["espessuras"] = list(m.pop("esp_map").values())
                    materiais.append(m)
                data["encomendas"].append(
                    {
                        "numero": num,
                        "cliente": str(e.get("cliente_codigo", "") or ""),
                        "posto_trabalho": str(e.get("posto_trabalho", "") or ""),
                        "nota_cliente": str(e.get("nota_cliente", "") or ""),
                        "data_criacao": _db_to_iso(e.get("data_criacao")),
                        "data_entrega": _db_to_iso(e.get("data_entrega"))[:10],
                        "tempo": 0.0,
                        "tempo_estimado": _to_num(e.get("tempo_estimado")) or 0.0,
                        "cativar": bool(e.get("cativar")) or bool(reservas_map.get(num)),
                        "observacoes": str(e.get("observacoes", "") or ""),
                        "estado": str(e.get("estado", "") or "Preparacao"),
                        "nota_transporte": str(e.get("nota_transporte", "") or ""),
                        "preco_transporte": _to_num(e.get("preco_transporte")) or 0.0,
                        "custo_transporte": _to_num(e.get("custo_transporte")) or 0.0,
                        "paletes": _to_num(e.get("paletes")) or 0.0,
                        "peso_bruto_kg": _to_num(e.get("peso_bruto_kg")) or 0.0,
                        "volume_m3": _to_num(e.get("volume_m3")) or 0.0,
                        "transportadora_id": str(e.get("transportadora_id", "") or ""),
                        "transportadora_nome": str(e.get("transportadora_nome", "") or ""),
                        "referencia_transporte": str(e.get("referencia_transporte", "") or ""),
                        "zona_transporte": str(e.get("zona_transporte", "") or ""),
                        "local_descarga": str(e.get("local_descarga", "") or ""),
                        "transporte_numero": str(e.get("transporte_numero", "") or ""),
                        "estado_transporte": str(e.get("estado_transporte", "") or ""),
                        "tipo_encomenda": str(e.get("tipo_encomenda", "") or "Cliente"),
                        "of_codigo": str(e.get("of_codigo", "") or ""),
                        "ordem_fabrico": {
                            "id": str(e.get("of_codigo", "") or ""),
                            "encomenda_id": num,
                            "estado": str(e.get("estado", "") or "Preparacao"),
                            "data": _db_to_iso(e.get("data_criacao"))[:10],
                        },
                        "materiais": materiais,
                        "reservas": reservas_map.get(num, []),
                        "montagem_itens": montagem_map.get(num, []),
                        "numero_orcamento": str(e.get("numero_orcamento", "") or ""),
                        "Observacoes": "",
                        "inicio_producao": "",
                        "tempo_pecas_min": 0.0,
                        "tempo_espessuras_min": 0.0,
                        "fim_producao": "",
                        "tempo_producao_min": 0.0,
                        "inicio_encomenda": "",
                        "fim_encomenda": "",
                        "estado_operador": "",
                        "obs_inicio": "",
                        "obs_interrupcao": "",
                        "tempo_por_espessura": {},
                        "espessuras": [],
                    }
                )

            ne_entregas_map = {}
            if "notas_encomenda_entregas" in tables:
                for ent in fetch_all("notas_encomenda_entregas", "id"):
                    ne_num = str(ent.get("ne_numero", "") or "")
                    if not ne_num:
                        continue
                    ne_entregas_map.setdefault(ne_num, []).append(
                        {
                            "data_registo": _db_to_iso(ent.get("data_registo")),
                            "data_entrega": _db_to_iso(ent.get("data_entrega"))[:10],
                            "data_documento": _db_to_iso(ent.get("data_documento"))[:10],
                            "guia": str(ent.get("guia", "") or ""),
                            "fatura": str(ent.get("fatura", "") or ""),
                            "obs": str(ent.get("obs", "") or ""),
                        }
                    )

            ne_docs_map = {}
            if "notas_encomenda_documentos" in tables:
                for d in fetch_all("notas_encomenda_documentos", "id"):
                    ne_num = str(d.get("ne_numero", "") or "")
                    if not ne_num:
                        continue
                    ne_docs_map.setdefault(ne_num, []).append(
                        {
                            "data_registo": _db_to_iso(d.get("data_registo")),
                            "tipo": str(d.get("tipo", "") or ""),
                            "titulo": str(d.get("titulo", "") or ""),
                            "caminho": str(d.get("caminho", "") or ""),
                            "guia": str(d.get("guia", "") or ""),
                            "fatura": str(d.get("fatura", "") or ""),
                            "data_entrega": _db_to_iso(d.get("data_entrega"))[:10],
                            "data_documento": _db_to_iso(d.get("data_documento"))[:10],
                            "obs": str(d.get("obs", "") or ""),
                        }
                    )

            ne_linha_entregas_map = {}
            if "notas_encomenda_linha_entregas" in tables:
                for ent_l in fetch_all("notas_encomenda_linha_entregas", "id"):
                    ne_num = str(ent_l.get("ne_numero", "") or "")
                    if not ne_num:
                        continue
                    try:
                        linha_ordem = int(_to_num(ent_l.get("linha_ordem")) or 0)
                    except Exception:
                        linha_ordem = 0
                    if linha_ordem <= 0:
                        continue
                    key = f"{ne_num}|{linha_ordem}"
                    ne_linha_entregas_map.setdefault(key, []).append(
                        {
                            "data_registo": _db_to_iso(ent_l.get("data_registo")),
                            "data_entrega": _db_to_iso(ent_l.get("data_entrega"))[:10],
                            "data_documento": _db_to_iso(ent_l.get("data_documento"))[:10],
                            "guia": str(ent_l.get("guia", "") or ""),
                            "fatura": str(ent_l.get("fatura", "") or ""),
                            "obs": str(ent_l.get("obs", "") or ""),
                            "qtd": _to_num(ent_l.get("qtd")) or 0.0,
                            "lote_fornecedor": str(ent_l.get("lote_fornecedor", "") or ""),
                            "localizacao": str(ent_l.get("localizacao", "") or ""),
                            "entrega_total": bool(ent_l.get("entrega_total")),
                            "stock_ref": str(ent_l.get("stock_ref", "") or ""),
                            "logistic_status": str(ent_l.get("logistic_status", "") or ""),
                            "inspection_status": str(ent_l.get("inspection_status", "") or ""),
                            "inspection_defect": str(ent_l.get("inspection_defect", "") or ""),
                            "inspection_decision": str(ent_l.get("inspection_decision", "") or ""),
                            "quality_status": str(ent_l.get("quality_status", "") or ""),
                            "quality_nc_id": str(ent_l.get("quality_nc_id", "") or ""),
                        }
                    )

            ne_linhas_map = {}
            ne_ord_counter = {}
            for l in fetch_all("notas_encomenda_linhas", "id"):
                ne_num = str(l.get("ne_numero", "") or "")
                if not ne_num:
                    continue
                ord_raw = _to_num(l.get("linha_ordem"))
                if ord_raw is None or int(ord_raw) <= 0:
                    ord_val = ne_ord_counter.get(ne_num, 0) + 1
                else:
                    ord_val = int(ord_raw)
                ne_ord_counter[ne_num] = ord_val
                q = _to_num(l.get("qtd")) or 0.0
                qtd_ent = _to_num(l.get("qtd_entregue"))
                if qtd_ent is None:
                    qtd_ent = q if bool(l.get("entregue")) else 0.0
                key = f"{ne_num}|{ord_val}"
                ne_linhas_map.setdefault(ne_num, []).append(
                    {
                        "ref": str(l.get("ref_material", "") or ""),
                        "descricao": str(l.get("descricao", "") or ""),
                        "fornecedor_linha": str(l.get("fornecedor_linha", "") or ""),
                        "origem": str(l.get("origem", "") or ""),
                        "qtd": q,
                        "unid": str(l.get("unid", "") or "UN"),
                        "preco": _to_num(l.get("preco")) or 0.0,
                        "desconto": _to_num(l.get("desconto")) or 0.0,
                        "iva": _to_num(l.get("iva")) or 23.0,
                        "total": _to_num(l.get("total")) or 0.0,
                        "entregue": bool(l.get("entregue")),
                        "qtd_entregue": qtd_ent,
                        "material": str(l.get("material", "") or ""),
                        "espessura": str(l.get("espessura", "") or ""),
                        "comprimento": _to_num(l.get("comprimento")) or 0.0,
                        "largura": _to_num(l.get("largura")) or 0.0,
                        "altura": _to_num(l.get("altura")) or 0.0,
                        "diametro": _to_num(l.get("diametro")) or 0.0,
                        "metros": _to_num(l.get("metros")) or 0.0,
                        "kg_m": _to_num(l.get("kg_m")) or 0.0,
                        "localizacao": str(l.get("localizacao", "") or ""),
                        "lote_fornecedor": str(l.get("lote_fornecedor", "") or ""),
                        "peso_unid": _to_num(l.get("peso_unid")) or 0.0,
                        "p_compra": _to_num(l.get("p_compra")) or 0.0,
                        "formato": str(l.get("formato", "") or ""),
                        "material_familia": str(l.get("material_familia", "") or ""),
                        "secao_tipo": str(l.get("secao_tipo", "") or ""),
                        "_stock_in": bool(l.get("stock_in")),
                        "guia_entrega": str(l.get("guia_entrega", "") or ""),
                        "fatura_entrega": str(l.get("fatura_entrega", "") or ""),
                        "data_doc_entrega": _db_to_iso(l.get("data_doc_entrega"))[:10],
                        "data_entrega_real": _db_to_iso(l.get("data_entrega_real"))[:10],
                        "obs_entrega": str(l.get("obs_entrega", "") or ""),
                        "logistic_status": str(l.get("logistic_status", "") or ""),
                        "inspection_status": str(l.get("inspection_status", "") or ""),
                        "inspection_defect": str(l.get("inspection_defect", "") or ""),
                        "inspection_decision": str(l.get("inspection_decision", "") or ""),
                        "quality_status": str(l.get("quality_status", "") or ""),
                        "quality_nc_id": str(l.get("quality_nc_id", "") or ""),
                        "entregas_linha": ne_linha_entregas_map.get(key, []),
                    }
                )

            for ne in fetch_all("notas_encomenda", "numero"):
                num = str(ne.get("numero", "") or "")
                if not num:
                    continue
                ne_geradas_txt = str(ne.get("ne_geradas", "") or "").strip()
                ne_geradas = [x.strip() for x in ne_geradas_txt.split(",") if x.strip()] if ne_geradas_txt else []
                entregas_rows = ne_entregas_map.get(num, [])
                guia_ult = str(ne.get("guia_ultima", "") or "")
                fatura_ult = str(ne.get("fatura_ultima", "") or "")
                fatura_path_ult = str(ne.get("fatura_caminho_ultima", "") or "")
                data_doc_ult = _db_to_iso(ne.get("data_doc_ultima"))[:10]
                data_ult_ent = _db_to_iso(ne.get("data_ultima_entrega"))[:10]
                docs_rows = ne_docs_map.get(num, [])
                if not docs_rows and fatura_path_ult:
                    docs_rows = [
                        {
                            "data_registo": "",
                            "tipo": "FATURA",
                            "titulo": fatura_ult or os.path.basename(fatura_path_ult),
                            "caminho": fatura_path_ult,
                            "guia": guia_ult,
                            "fatura": fatura_ult,
                            "data_documento": data_doc_ult,
                            "obs": "",
                        }
                    ]
                if not entregas_rows and (guia_ult or fatura_ult or data_doc_ult or data_ult_ent):
                    entregas_rows = [
                        {
                            "data_registo": "",
                            "data_entrega": data_ult_ent,
                            "data_documento": data_doc_ult,
                            "guia": guia_ult,
                            "fatura": fatura_ult,
                            "obs": "",
                        }
                    ]
                data["notas_encomenda"].append(
                    {
                        "numero": num,
                        "fornecedor": str(ne.get("fornecedor_id", "") or ""),
                        "fornecedor_id": str(ne.get("fornecedor_id", "") or ""),
                        "contacto": str(ne.get("contacto", "") or ""),
                        "data_entrega": _db_to_iso(ne.get("data_entrega"))[:10],
                        "obs": str(ne.get("obs", "") or ""),
                        "local_descarga": str(ne.get("local_descarga", "") or ""),
                        "meio_transporte": str(ne.get("meio_transporte", "") or ""),
                        "linhas": ne_linhas_map.get(num, []),
                        "total": _to_num(ne.get("total")) or 0.0,
                        "estado": str(ne.get("estado", "") or "Em edicao"),
                        "oculta": bool(ne.get("oculta")),
                        "_draft": bool(ne.get("is_draft")),
                        "entregas": entregas_rows,
                        "guia_ultima": guia_ult,
                        "fatura_ultima": fatura_ult,
                        "fatura_caminho_ultima": fatura_path_ult,
                        "data_doc_ultima": data_doc_ult,
                        "data_ultima_entrega": data_ult_ent,
                        "documentos": docs_rows,
                        "origem_cotacao": str(ne.get("origem_cotacao", "") or ""),
                        "ne_geradas": ne_geradas,
                    }
                )

            exp_linhas_map = {}
            for l in fetch_all("expedicao_linhas", "id"):
                ex_num = str(l.get("expedicao_numero", "") or "")
                exp_linhas_map.setdefault(ex_num, []).append(
                    {
                        "encomenda": str(l.get("encomenda_numero", "") or ""),
                        "peca_id": str(l.get("peca_id", "") or ""),
                        "ref_interna": str(l.get("ref_interna", "") or ""),
                        "ref_externa": str(l.get("ref_externa", "") or ""),
                        "descricao": str(l.get("descricao", "") or ""),
                        "qtd": _to_num(l.get("qtd")) or 0.0,
                        "unid": str(l.get("unid", "") or "UN"),
                        "peso": _to_num(l.get("peso")) or 0.0,
                        "manual": bool(l.get("manual")),
                    }
                )

            for ex in fetch_all("expedicoes", "id"):
                num = str(ex.get("numero", "") or "")
                if not num:
                    continue
                data["expedicoes"].append(
                    {
                        "numero": num,
                        "tipo": str(ex.get("tipo", "") or "OFF"),
                        "encomenda": str(ex.get("encomenda_numero", "") or ""),
                        "cliente": str(ex.get("cliente_codigo", "") or ""),
                        "cliente_nome": str(ex.get("cliente_nome", "") or ""),
                        "codigo_at": str(ex.get("codigo_at", "") or ""),
                        "serie_id": str(ex.get("serie_id", "") or ""),
                        "seq_num": int(_to_num(ex.get("seq_num")) or 0),
                        "at_validation_code": str(ex.get("at_validation_code", "") or ""),
                        "atcud": str(ex.get("atcud", "") or ""),
                        "emitente_nome": str(ex.get("emitente_nome", "") or ""),
                        "emitente_nif": str(ex.get("emitente_nif", "") or ""),
                        "emitente_morada": str(ex.get("emitente_morada", "") or ""),
                        "destinatario": str(ex.get("destinatario", "") or ""),
                        "dest_nif": str(ex.get("dest_nif", "") or ""),
                        "dest_morada": str(ex.get("dest_morada", "") or ""),
                        "local_carga": str(ex.get("local_carga", "") or ""),
                        "local_descarga": str(ex.get("local_descarga", "") or ""),
                        "data_emissao": _db_to_iso(ex.get("data_emissao")),
                        "data_transporte": _db_to_iso(ex.get("data_transporte")),
                        "matricula": str(ex.get("matricula", "") or ""),
                        "transportador": str(ex.get("transportador", "") or ""),
                        "estado": str(ex.get("estado", "") or "Emitida"),
                        "observacoes": str(ex.get("observacoes", "") or ""),
                        "created_by": str(ex.get("created_by", "") or ""),
                        "anulada": bool(ex.get("anulada")),
                        "anulada_motivo": str(ex.get("anulada_motivo", "") or ""),
                        "linhas": exp_linhas_map.get(num, []),
                    }
                )

            transport_stops_map = {}
            if "transportes_paragens" in tables:
                for stop in fetch_all("transportes_paragens", "transporte_numero, ordem, id"):
                    tr_num = str(stop.get("transporte_numero", "") or "").strip()
                    if not tr_num:
                        continue
                    transport_stops_map.setdefault(tr_num, []).append(
                        {
                            "ordem": int(_to_num(stop.get("ordem")) or 0),
                            "encomenda_numero": str(stop.get("encomenda_numero", "") or ""),
                            "expedicao_numero": str(stop.get("expedicao_numero", "") or ""),
                            "cliente_codigo": str(stop.get("cliente_codigo", "") or ""),
                            "cliente_nome": str(stop.get("cliente_nome", "") or ""),
                            "zona_transporte": str(stop.get("zona_transporte", "") or ""),
                            "local_descarga": str(stop.get("local_descarga", "") or ""),
                            "contacto": str(stop.get("contacto", "") or ""),
                            "telefone": str(stop.get("telefone", "") or ""),
                            "data_planeada": _db_to_iso(stop.get("data_planeada")),
                            "paletes": _to_num(stop.get("paletes")) or 0.0,
                            "peso_bruto_kg": _to_num(stop.get("peso_bruto_kg")) or 0.0,
                            "volume_m3": _to_num(stop.get("volume_m3")) or 0.0,
                            "preco_transporte": _to_num(stop.get("preco_transporte")) or 0.0,
                            "custo_transporte": _to_num(stop.get("custo_transporte")) or 0.0,
                            "transportadora_id": str(stop.get("transportadora_id", "") or ""),
                            "transportadora_nome": str(stop.get("transportadora_nome", "") or ""),
                            "referencia_transporte": str(stop.get("referencia_transporte", "") or ""),
                            "check_carga_ok": bool(stop.get("check_carga_ok")),
                            "check_docs_ok": bool(stop.get("check_docs_ok")),
                            "check_paletes_ok": bool(stop.get("check_paletes_ok")),
                            "pod_estado": str(stop.get("pod_estado", "") or ""),
                            "pod_recebido_nome": str(stop.get("pod_recebido_nome", "") or ""),
                            "pod_recebido_at": _db_to_iso(stop.get("pod_recebido_at")),
                            "pod_obs": str(stop.get("pod_obs", "") or ""),
                            "estado": str(stop.get("estado", "") or ""),
                            "observacoes": str(stop.get("observacoes", "") or ""),
                        }
                    )

            if "transportes" in tables:
                data["transportes"] = []
                for tr in fetch_all("transportes", "data_planeada, numero"):
                    num = str(tr.get("numero", "") or "").strip()
                    if not num:
                        continue
                    data["transportes"].append(
                        {
                            "numero": num,
                            "tipo_responsavel": str(tr.get("tipo_responsavel", "") or ""),
                            "estado": str(tr.get("estado", "") or ""),
                            "data_planeada": _db_to_iso(tr.get("data_planeada"))[:10],
                            "hora_saida": str(tr.get("hora_saida", "") or ""),
                            "viatura": str(tr.get("viatura", "") or ""),
                            "matricula": str(tr.get("matricula", "") or ""),
                            "motorista": str(tr.get("motorista", "") or ""),
                            "telefone_motorista": str(tr.get("telefone_motorista", "") or ""),
                            "origem": str(tr.get("origem", "") or ""),
                            "transportadora_id": str(tr.get("transportadora_id", "") or ""),
                            "transportadora_nome": str(tr.get("transportadora_nome", "") or ""),
                            "referencia_transporte": str(tr.get("referencia_transporte", "") or ""),
                            "custo_previsto": _to_num(tr.get("custo_previsto")) or 0.0,
                            "paletes_total_manual": _to_num(tr.get("paletes_total_manual")) or 0.0,
                            "peso_total_manual_kg": _to_num(tr.get("peso_total_manual_kg")) or 0.0,
                            "volume_total_manual_m3": _to_num(tr.get("volume_total_manual_m3")) or 0.0,
                            "pedido_transporte_estado": str(tr.get("pedido_transporte_estado", "") or ""),
                            "pedido_transporte_ref": str(tr.get("pedido_transporte_ref", "") or ""),
                            "pedido_transporte_at": _db_to_iso(tr.get("pedido_transporte_at")),
                            "pedido_transporte_by": str(tr.get("pedido_transporte_by", "") or ""),
                            "pedido_transporte_obs": str(tr.get("pedido_transporte_obs", "") or ""),
                            "pedido_resposta_obs": str(tr.get("pedido_resposta_obs", "") or ""),
                            "pedido_confirmado_at": _db_to_iso(tr.get("pedido_confirmado_at")),
                            "pedido_confirmado_by": str(tr.get("pedido_confirmado_by", "") or ""),
                            "pedido_recusado_at": _db_to_iso(tr.get("pedido_recusado_at")),
                            "pedido_recusado_by": str(tr.get("pedido_recusado_by", "") or ""),
                            "observacoes": str(tr.get("observacoes", "") or ""),
                            "created_by": str(tr.get("created_by", "") or ""),
                            "created_at": _db_to_iso(tr.get("created_at")),
                            "updated_at": _db_to_iso(tr.get("updated_at")),
                            "paragens": sorted(
                                list(transport_stops_map.get(num, []) or []),
                                key=lambda row: (int(_to_num(row.get("ordem")) or 0), str(row.get("encomenda_numero", "") or "")),
                            ),
                        }
                    )
            if "transportes_tarifarios" in tables:
                data["transportes_tarifarios"] = []
                for row in fetch_all("transportes_tarifarios", "transportadora_nome, zona, id"):
                    data["transportes_tarifarios"].append(
                        {
                            "id": int(_to_num(row.get("id")) or 0),
                            "transportadora_id": str(row.get("transportadora_id", "") or ""),
                            "transportadora_nome": str(row.get("transportadora_nome", "") or ""),
                            "zona": str(row.get("zona", "") or ""),
                            "valor_base": _to_num(row.get("valor_base")) or 0.0,
                            "valor_por_palete": _to_num(row.get("valor_por_palete")) or 0.0,
                            "valor_por_kg": _to_num(row.get("valor_por_kg")) or 0.0,
                            "valor_por_m3": _to_num(row.get("valor_por_m3")) or 0.0,
                            "custo_minimo": _to_num(row.get("custo_minimo")) or 0.0,
                            "ativo": bool(row.get("ativo", 1)),
                            "observacoes": str(row.get("observacoes", "") or ""),
                        }
                    )

            fat_invoice_map = {}
            if "faturacao_faturas" in tables:
                for row in fetch_all("faturacao_faturas", "data_emissao, id"):
                    reg_num = str(row.get("registo_numero", "") or "").strip()
                    if not reg_num:
                        continue
                    fat_invoice_map.setdefault(reg_num, []).append(
                        {
                            "id": str(row.get("documento_id", "") or "").strip(),
                            "doc_type": str(row.get("doc_type", "") or "").strip(),
                            "numero_fatura": str(row.get("numero_fatura", "") or "").strip(),
                            "serie": str(row.get("serie", "") or "").strip(),
                            "serie_id": str(row.get("serie_id", "") or "").strip(),
                            "seq_num": int(_to_num(row.get("seq_num")) or 0),
                            "at_validation_code": str(row.get("at_validation_code", "") or "").strip(),
                            "atcud": str(row.get("atcud", "") or "").strip(),
                            "guia_numero": str(row.get("guia_numero", "") or "").strip(),
                            "data_emissao": _db_to_iso(row.get("data_emissao"))[:10],
                            "data_vencimento": _db_to_iso(row.get("data_vencimento"))[:10],
                            "moeda": str(row.get("moeda", "") or "").strip(),
                            "iva_perc": _to_num(row.get("iva_perc")) or 0.0,
                            "subtotal": _to_num(row.get("subtotal")) or 0.0,
                            "valor_iva": _to_num(row.get("valor_iva")) or 0.0,
                            "valor_total": _to_num(row.get("valor_total")) or 0.0,
                            "caminho": str(row.get("caminho", "") or "").strip(),
                            "obs": str(row.get("obs", "") or "").strip(),
                            "estado": str(row.get("estado", "") or "").strip(),
                            "anulada": bool(_to_num(row.get("anulada")) or 0),
                            "anulada_motivo": str(row.get("anulada_motivo", "") or "").strip(),
                            "anulada_at": _db_to_iso(row.get("anulada_at")),
                            "legal_invoice_no": str(row.get("legal_invoice_no", "") or "").strip(),
                            "system_entry_date": _db_to_iso(row.get("system_entry_date")),
                            "source_id": str(row.get("source_id", "") or "").strip(),
                            "source_billing": str(row.get("source_billing", "") or "").strip(),
                            "status_source_id": str(row.get("status_source_id", "") or "").strip(),
                            "hash": str(row.get("hash", "") or "").strip(),
                            "hash_control": str(row.get("hash_control", "") or "").strip(),
                            "previous_hash": str(row.get("previous_hash", "") or "").strip(),
                            "document_snapshot_json": str(row.get("document_snapshot_json", "") or ""),
                            "communication_status": str(row.get("communication_status", "") or "").strip(),
                            "communication_filename": str(row.get("communication_filename", "") or "").strip(),
                            "communication_error": str(row.get("communication_error", "") or "").strip(),
                            "communicated_at": _db_to_iso(row.get("communicated_at")),
                            "communication_batch_id": str(row.get("communication_batch_id", "") or "").strip(),
                            "created_at": _db_to_iso(row.get("created_at")),
                        }
                    )

            fat_payment_map = {}
            if "faturacao_pagamentos" in tables:
                for row in fetch_all("faturacao_pagamentos", "data_pagamento, id"):
                    reg_num = str(row.get("registo_numero", "") or "").strip()
                    if not reg_num:
                        continue
                    fat_payment_map.setdefault(reg_num, []).append(
                        {
                            "id": str(row.get("pagamento_id", "") or "").strip(),
                            "fatura_id": str(row.get("fatura_documento_id", "") or "").strip(),
                            "data_pagamento": _db_to_iso(row.get("data_pagamento"))[:10],
                            "valor": _to_num(row.get("valor")) or 0.0,
                            "metodo": str(row.get("metodo", "") or "").strip(),
                            "referencia": str(row.get("referencia", "") or "").strip(),
                            "titulo_comprovativo": str(row.get("titulo_comprovativo", "") or "").strip(),
                            "caminho_comprovativo": str(row.get("caminho_comprovativo", "") or "").strip(),
                            "obs": str(row.get("obs", "") or "").strip(),
                            "created_at": _db_to_iso(row.get("created_at")),
                        }
                    )

            for row in fetch_all("faturacao_registos", "numero"):
                reg_num = str(row.get("numero", "") or "").strip()
                if not reg_num:
                    continue
                data["faturacao"].append(
                    {
                        "numero": reg_num,
                        "origem": str(row.get("origem", "") or "").strip(),
                        "orcamento_numero": str(row.get("orcamento_numero", "") or "").strip(),
                        "encomenda_numero": str(row.get("encomenda_numero", "") or "").strip(),
                        "cliente_codigo": str(row.get("cliente_codigo", "") or "").strip(),
                        "cliente_nome": str(row.get("cliente_nome", "") or "").strip(),
                        "data_venda": _db_to_iso(row.get("data_venda"))[:10],
                        "data_vencimento": _db_to_iso(row.get("data_vencimento"))[:10],
                        "valor_venda_manual": _to_num(row.get("valor_venda_manual")) or 0.0,
                        "estado_pagamento_manual": str(row.get("estado_pagamento_manual", "") or "").strip(),
                        "obs": str(row.get("obs", "") or "").strip(),
                        "created_at": _db_to_iso(row.get("created_at")),
                        "updated_at": _db_to_iso(row.get("updated_at")),
                        "faturas": fat_invoice_map.get(reg_num, []),
                        "pagamentos": fat_payment_map.get(reg_num, []),
                    }
                )

            for s in fetch_all("stock_log", "id"):
                data["stock_log"].append(
                    {
                        "data": _db_to_iso(s.get("data")),
                        "acao": str(s.get("acao", "") or ""),
                        "operador": str(s.get("operador", "") or ""),
                        "detalhes": str(s.get("detalhes", "") or ""),
                    }
                )

            for r in fetch_all("produtos_mov", "id"):
                mov = normalize_produto_mov_row(
                    {
                        "data": _db_to_iso(r.get("data")),
                        "tipo": str(r.get("tipo", "") or ""),
                        "operador": str(r.get("operador", "") or ""),
                        "codigo": str(r.get("codigo", "") or ""),
                        "descricao": str(r.get("descricao", "") or ""),
                        "qtd": _to_num(r.get("qtd")) or 0.0,
                        "antes": _to_num(r.get("antes")) or 0.0,
                        "depois": _to_num(r.get("depois")) or 0.0,
                        "obs": str(r.get("obs", "") or ""),
                        "origem": str(r.get("origem", "") or ""),
                        "ref_doc": str(r.get("ref_doc", "") or ""),
                    }
                )
                if mov:
                    data["produtos_mov"].append(mov)

            for row in fetch_all("quality_nonconformities", "created_at, id"):
                data["quality_nonconformities"].append(
                    {
                        "id": str(row.get("id", "") or ""),
                        "origem": str(row.get("origem", "") or ""),
                        "referencia": str(row.get("referencia", "") or ""),
                        "entidade_tipo": str(row.get("entidade_tipo", "") or ""),
                        "entidade_id": str(row.get("entidade_id", "") or ""),
                        "entidade_label": str(row.get("entidade_label", "") or ""),
                        "tipo": str(row.get("tipo", "") or ""),
                        "gravidade": str(row.get("gravidade", "") or ""),
                        "estado": str(row.get("estado", "") or ""),
                        "responsavel": str(row.get("responsavel", "") or ""),
                        "prazo": _db_to_iso(row.get("prazo"))[:10],
                        "descricao": str(row.get("descricao", "") or ""),
                        "causa": str(row.get("causa", "") or ""),
                        "acao": str(row.get("acao", "") or ""),
                        "eficacia": str(row.get("eficacia", "") or ""),
                        "fornecedor_id": str(row.get("fornecedor_id", "") or ""),
                        "fornecedor_nome": str(row.get("fornecedor_nome", "") or ""),
                        "material_id": str(row.get("material_id", "") or ""),
                        "lote_fornecedor": str(row.get("lote_fornecedor", "") or ""),
                        "ne_numero": str(row.get("ne_numero", "") or ""),
                        "guia": str(row.get("guia", "") or ""),
                        "fatura": str(row.get("fatura", "") or ""),
                        "decisao": str(row.get("decisao", "") or ""),
                        "movement_id": str(row.get("movement_id", "") or ""),
                        "qtd_recebida": _to_num(row.get("qtd_recebida")) or 0.0,
                        "qtd_aprovada": _to_num(row.get("qtd_aprovada")) or 0.0,
                        "qtd_rejeitada": _to_num(row.get("qtd_rejeitada")) or 0.0,
                        "qtd_pendente": _to_num(row.get("qtd_pendente")) or 0.0,
                        "created_at": _db_to_iso(row.get("created_at")),
                        "updated_at": _db_to_iso(row.get("updated_at")),
                        "created_by": str(row.get("created_by", "") or ""),
                        "updated_by": str(row.get("updated_by", "") or ""),
                        "closed_at": _db_to_iso(row.get("closed_at")),
                        "closed_by": str(row.get("closed_by", "") or ""),
                    }
                )

            for row in fetch_all("quality_documents", "tipo, titulo, id"):
                data["quality_documents"].append(
                    {
                        "id": str(row.get("id", "") or ""),
                        "titulo": str(row.get("titulo", "") or ""),
                        "tipo": str(row.get("tipo", "") or ""),
                        "entidade": str(row.get("entidade", "") or ""),
                        "referencia": str(row.get("referencia", "") or ""),
                        "entidade_tipo": str(row.get("entidade_tipo", "") or ""),
                        "entidade_id": str(row.get("entidade_id", "") or ""),
                        "versao": str(row.get("versao", "") or ""),
                        "estado": str(row.get("estado", "") or ""),
                        "responsavel": str(row.get("responsavel", "") or ""),
                        "caminho": str(row.get("caminho", "") or ""),
                        "obs": str(row.get("obs", "") or ""),
                        "created_at": _db_to_iso(row.get("created_at")),
                        "updated_at": _db_to_iso(row.get("updated_at")),
                        "created_by": str(row.get("created_by", "") or ""),
                        "updated_by": str(row.get("updated_by", "") or ""),
                    }
                )

            for row in fetch_all("quality_audit_log", "created_at, id"):
                event = {
                    "id": str(row.get("id", "") or ""),
                    "created_at": _db_to_iso(row.get("created_at")),
                    "user": str(row.get("user_name", "") or ""),
                    "action": str(row.get("action", "") or ""),
                    "entity_type": str(row.get("entity_type", "") or ""),
                    "entity_id": str(row.get("entity_id", "") or ""),
                    "summary": str(row.get("summary", "") or ""),
                }
                for source_key, target_key in (("before_json", "before"), ("after_json", "after")):
                    raw = row.get(source_key)
                    if raw:
                        try:
                            event[target_key] = json.loads(str(raw))
                        except Exception:
                            event[target_key] = str(raw)
                data["audit_log"].append(event)

            for ev in fetch_all("op_eventos", "id"):
                data["op_eventos"].append(
                    {
                        "created_at": _db_to_iso(ev.get("created_at")),
                        "evento": str(ev.get("evento", "") or ""),
                        "encomenda_numero": str(ev.get("encomenda_numero", "") or ""),
                        "peca_id": str(ev.get("peca_id", "") or ""),
                        "ref_interna": str(ev.get("ref_interna", "") or ""),
                        "material": str(ev.get("material", "") or ""),
                        "espessura": str(ev.get("espessura", "") or ""),
                        "operacao": str(ev.get("operacao", "") or ""),
                        "operador": str(ev.get("operador", "") or ""),
                        "qtd_ok": _to_num(ev.get("qtd_ok")) or 0.0,
                        "qtd_nok": _to_num(ev.get("qtd_nok")) or 0.0,
                        "info": str(ev.get("info", "") or ""),
                    }
                )

            for pa in fetch_all("op_paragens", "id"):
                data["op_paragens"].append(
                    {
                        "created_at": _db_to_iso(pa.get("created_at")),
                        "fechada_at": _db_to_iso(pa.get("fechada_at")),
                        "encomenda_numero": str(pa.get("encomenda_numero", "") or ""),
                        "peca_id": str(pa.get("peca_id", "") or ""),
                        "ref_interna": str(pa.get("ref_interna", "") or ""),
                        "material": str(pa.get("material", "") or ""),
                        "espessura": str(pa.get("espessura", "") or ""),
                        "operador": str(pa.get("operador", "") or ""),
                        "origem": str(pa.get("origem", "") or ""),
                        "estado": str(pa.get("estado", "") or ""),
                        "causa": str(pa.get("causa", "") or ""),
                        "detalhe": str(pa.get("detalhe", "") or ""),
                        "grupo_id": str(pa.get("grupo_id", "") or ""),
                        "duracao_min": _to_num(pa.get("duracao_min")) or 0.0,
                    }
                )

            if "qualidade" in tables:
                for q in fetch_all("qualidade", "id"):
                    data["qualidade"].append(
                        {
                            "encomenda": str(q.get("encomenda", "") or ""),
                            "peca": str(q.get("peca", "") or ""),
                            "ok": _to_num(q.get("ok_qtd") if "ok_qtd" in q else q.get("ok")) or 0.0,
                            "nok": _to_num(q.get("nok_qtd") if "nok_qtd" in q else q.get("nok")) or 0.0,
                            "motivo": str(q.get("motivo", "") or ""),
                            "data": _db_to_iso(q.get("data")),
                        }
                    )

        _apply_runtime_state_payload(
            data,
            _app_config_load_json(RUNTIME_STATE_CONFIG_KEY, RUNTIME_STATE_CONFIG_FILE, conn=conn),
        )
        _rebuild_runtime_sequences(data)
        return data
    finally:
        conn.close()


def _mysql_save_relational_data(data, conn=None):
    global _MYSQL_SCHEMA_SYNCED
    attempts = 1 if conn is not None else _MYSQL_SAVE_RETRY_COUNT
    last_error = None
    for attempt in range(max(1, attempts)):
        own_conn = conn is None
        active_conn = conn if conn is not None else _mysql_connect()
        lock_acquired = False
        try:
            with active_conn.cursor() as cur:
                try:
                    cur.execute(f"SET SESSION innodb_lock_wait_timeout = {_MYSQL_SAVE_LOCK_TIMEOUT_SEC}")
                except Exception:
                    pass
                lock_acquired = _mysql_named_lock_acquire(cur, _mysql_save_lock_name(), timeout_sec=_MYSQL_SAVE_LOCK_TIMEOUT_SEC)
                if not lock_acquired:
                    raise RuntimeError("Nao foi possivel obter o lock global de gravacao do LuGEST.")
                # NOTE: despite the name, this routine syncs both schema and data snapshot.
                # It must run on every save to persist changes.
                _mysql_sync_relational_schema(cur, data)
                _app_config_save_json(
                    RUNTIME_STATE_CONFIG_KEY,
                    RUNTIME_STATE_CONFIG_FILE,
                    _runtime_state_payload(data),
                    conn=active_conn,
                )
                _MYSQL_SCHEMA_SYNCED = True
            active_conn.commit()
            return
        except Exception as ex:
            last_error = ex
            _MYSQL_SCHEMA_SYNCED = False
            _mysql_schema_cache_reset()
            try:
                active_conn.rollback()
            except Exception:
                pass
            if (attempt + 1) >= max(1, attempts) or not _mysql_retryable_write_error(ex):
                raise
            time.sleep(0.35 * (attempt + 1))
        finally:
            if lock_acquired:
                try:
                    with active_conn.cursor() as cur:
                        _mysql_named_lock_release(cur, _mysql_save_lock_name())
                except Exception:
                    pass
            if own_conn and active_conn is not None:
                active_conn.close()
    if last_error is not None:
        raise last_error


def mysql_refresh_runtime_impulse_data(data, cleanup_orphans=True):
    if not USE_MYSQL_STORAGE or not MYSQL_AVAILABLE or not isinstance(data, dict):
        return
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            tables = _mysql_existing_tables(cur, force=True)
            if cleanup_orphans:
                if "op_eventos" in tables and "encomendas" in tables and "pecas" in tables:
                    cur.execute(
                        """
                        DELETE oe
                        FROM op_eventos oe
                        LEFT JOIN encomendas e ON e.numero = oe.encomenda_numero
                        LEFT JOIN pecas p ON p.id = oe.peca_id AND p.encomenda_numero = oe.encomenda_numero
                        WHERE (
                            COALESCE(oe.encomenda_numero, '') <> ''
                            AND e.numero IS NULL
                        ) OR (
                            COALESCE(oe.peca_id, '') <> ''
                            AND p.id IS NULL
                        )
                        """
                    )
                if "op_paragens" in tables and "encomendas" in tables and "pecas" in tables:
                    cur.execute(
                        """
                        DELETE op
                        FROM op_paragens op
                        LEFT JOIN encomendas e ON e.numero = op.encomenda_numero
                        LEFT JOIN pecas p ON p.id = op.peca_id AND p.encomenda_numero = op.encomenda_numero
                        WHERE (
                            COALESCE(op.encomenda_numero, '') <> ''
                            AND e.numero IS NULL
                        ) OR (
                            COALESCE(op.peca_id, '') <> ''
                            AND p.id IS NULL
                        )
                        """
                    )
                if "peca_operacoes_execucao" in tables and "encomendas" in tables and "pecas" in tables:
                    cur.execute(
                        """
                        DELETE px
                        FROM peca_operacoes_execucao px
                        LEFT JOIN encomendas e ON e.numero = px.encomenda_numero
                        LEFT JOIN pecas p ON p.id = px.peca_id AND p.encomenda_numero = px.encomenda_numero
                        WHERE (
                            COALESCE(px.encomenda_numero, '') <> ''
                            AND e.numero IS NULL
                        ) OR (
                            COALESCE(px.peca_id, '') <> ''
                            AND p.id IS NULL
                        )
                        """
                    )
            conn.commit()

            data["op_eventos"] = []
            if "op_eventos" in tables:
                cur.execute("SELECT * FROM op_eventos ORDER BY id")
                for ev in cur.fetchall() or []:
                    data["op_eventos"].append(
                        {
                            "created_at": _db_to_iso(ev.get("created_at")),
                            "evento": str(ev.get("evento", "") or ""),
                            "encomenda_numero": str(ev.get("encomenda_numero", "") or ""),
                            "peca_id": str(ev.get("peca_id", "") or ""),
                            "ref_interna": str(ev.get("ref_interna", "") or ""),
                            "material": str(ev.get("material", "") or ""),
                            "espessura": str(ev.get("espessura", "") or ""),
                            "operacao": str(ev.get("operacao", "") or ""),
                            "operador": str(ev.get("operador", "") or ""),
                            "qtd_ok": _to_num(ev.get("qtd_ok")) or 0.0,
                            "qtd_nok": _to_num(ev.get("qtd_nok")) or 0.0,
                            "info": str(ev.get("info", "") or ""),
                        }
                    )

            data["op_paragens"] = []
            if "op_paragens" in tables:
                cur.execute("SELECT * FROM op_paragens ORDER BY id")
                for pa in cur.fetchall() or []:
                    data["op_paragens"].append(
                        {
                            "created_at": _db_to_iso(pa.get("created_at")),
                            "fechada_at": _db_to_iso(pa.get("fechada_at")),
                            "encomenda_numero": str(pa.get("encomenda_numero", "") or ""),
                            "peca_id": str(pa.get("peca_id", "") or ""),
                            "ref_interna": str(pa.get("ref_interna", "") or ""),
                            "material": str(pa.get("material", "") or ""),
                            "espessura": str(pa.get("espessura", "") or ""),
                            "operador": str(pa.get("operador", "") or ""),
                            "origem": str(pa.get("origem", "") or ""),
                            "estado": str(pa.get("estado", "") or ""),
                            "causa": str(pa.get("causa", "") or ""),
                            "detalhe": str(pa.get("detalhe", "") or ""),
                            "grupo_id": str(pa.get("grupo_id", "") or ""),
                            "duracao_min": _to_num(pa.get("duracao_min")) or 0.0,
                        }
                    )
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


def mysql_log_ne_linha_historico(
    *,
    evento,
    origem_menu="",
    utilizador="",
    guia_numero="",
    produto_codigo="",
    descricao="",
    qtd=None,
    unid="",
    destinatario="",
    observacoes="",
    payload=None,
):
    if not USE_MYSQL_STORAGE or not MYSQL_AVAILABLE:
        return
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            if not _mysql_runtime_schema_ready("ne_linhas_historico"):
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS ne_linhas_historico (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        created_at DATETIME NOT NULL,
                        evento VARCHAR(30) NOT NULL,
                        origem_menu VARCHAR(80) NULL,
                        utilizador VARCHAR(80) NULL,
                        guia_numero VARCHAR(30) NULL,
                        produto_codigo VARCHAR(20) NULL,
                        descricao VARCHAR(255) NULL,
                        qtd DECIMAL(10,2) NULL,
                        unid VARCHAR(20) NULL,
                        destinatario VARCHAR(150) NULL,
                        observacoes TEXT NULL,
                        payload_json LONGTEXT NULL,
                        INDEX idx_ne_linhas_hist_created_at (created_at),
                        INDEX idx_ne_linhas_hist_evento (evento),
                        INDEX idx_ne_linhas_hist_guia (guia_numero)
                    )
                    """
                )
                _mysql_runtime_schema_mark("ne_linhas_historico")
            cur.execute(
                """
                INSERT INTO ne_linhas_historico (
                    created_at, evento, origem_menu, utilizador, guia_numero, produto_codigo, descricao, qtd, unid,
                    destinatario, observacoes, payload_json
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    _to_mysql_datetime(now_iso()),
                    _clip(evento, 30),
                    _clip(origem_menu, 80),
                    _clip(utilizador, 80),
                    _clip(guia_numero, 30),
                    _clip(produto_codigo, 20),
                    _clip(descricao, 255),
                    _to_num(qtd),
                    _clip(unid, 20),
                    _clip(destinatario, 150),
                    observacoes,
                    json.dumps(payload or {}, ensure_ascii=False),
                ),
            )
        conn.commit()
    except Exception:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        if conn:
            conn.close()


def mysql_log_production_event(
    *,
    evento,
    encomenda_numero="",
    peca_id="",
    ref_interna="",
    material="",
    espessura="",
    operacao="",
    operador="",
    qtd_ok=None,
    qtd_nok=None,
    info="",
    causa_paragem="",
    duracao_min=None,
    created_at="",
    grupo_id="",
):
    global _RUNTIME_DATA_REF
    ts_now = str(created_at or now_iso()).strip() or now_iso()
    grupo_key = _clip(grupo_id, 80)
    ev_row = {
        "created_at": ts_now,
        "evento": _clip(evento, 30),
        "encomenda_numero": _clip(encomenda_numero, 30),
        "peca_id": _clip(peca_id, 30),
        "ref_interna": _clip(ref_interna, 60),
        "material": _clip(material, 100),
        "espessura": _clip(espessura, 20),
        "operacao": _clip(operacao, 80),
        "operador": _clip(operador, 80),
        "qtd_ok": _to_num(qtd_ok),
        "qtd_nok": _to_num(qtd_nok),
        "info": info,
    }
    try:
        if isinstance(_RUNTIME_DATA_REF, dict):
            _RUNTIME_DATA_REF.setdefault("op_eventos", []).append(dict(ev_row))
            evento_norm_mem = str(evento or "").strip().upper()
            causa_txt_mem = str(causa_paragem or "").strip()
            if evento_norm_mem in ("PARAGEM", "STOP") and causa_txt_mem:
                _RUNTIME_DATA_REF.setdefault("op_paragens", []).append(
                    {
                        "created_at": ts_now,
                        "fechada_at": "",
                        "encomenda_numero": _clip(encomenda_numero, 30),
                        "peca_id": _clip(peca_id, 30),
                        "ref_interna": _clip(ref_interna, 60),
                        "material": _clip(material, 100),
                        "espessura": _clip(espessura, 20),
                        "operador": _clip(operador, 80),
                        "origem": "AVARIA",
                        "estado": "ABERTA",
                        "causa": _clip(causa_txt_mem, 120),
                        "detalhe": info,
                        "grupo_id": grupo_key,
                        "duracao_min": _to_num(duracao_min),
                    }
                )
            elif evento_norm_mem in ("RESUME_PIECE", "START_OP", "FINISH_OP", "CLOSE_AVARIA"):
                try:
                    p_id = _clip(peca_id, 30)
                    e_num = _clip(encomenda_numero, 30)
                    if p_id and e_num:
                        now_dt = iso_to_dt(ts_now)
                        for row in reversed(_RUNTIME_DATA_REF.get("op_paragens", []) or []):
                            if not isinstance(row, dict):
                                continue
                            if str(row.get("encomenda_numero", "") or "") != e_num:
                                continue
                            if str(row.get("peca_id", "") or "") != p_id:
                                continue
                            if row.get("duracao_min") not in (None, "", 0, 0.0):
                                continue
                            created_dt = iso_to_dt(row.get("created_at"))
                            if created_dt and now_dt and now_dt >= created_dt:
                                mins = max(0.0, (now_dt - created_dt).total_seconds() / 60.0)
                                row["duracao_min"] = round(mins, 2)
                                row["fechada_at"] = ts_now
                                row["estado"] = "FECHADA"
                                break
                except Exception:
                    pass
    except Exception:
        pass

    if not USE_MYSQL_STORAGE or not MYSQL_AVAILABLE:
        return
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            if not _mysql_runtime_schema_ready("op_eventos"):
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS op_eventos (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        created_at DATETIME NOT NULL,
                        evento VARCHAR(30) NOT NULL,
                        encomenda_numero VARCHAR(30) NULL,
                        peca_id VARCHAR(30) NULL,
                        ref_interna VARCHAR(60) NULL,
                        material VARCHAR(100) NULL,
                        espessura VARCHAR(20) NULL,
                        operacao VARCHAR(80) NULL,
                        operador VARCHAR(80) NULL,
                        qtd_ok DECIMAL(10,2) NULL,
                        qtd_nok DECIMAL(10,2) NULL,
                        info TEXT NULL,
                        INDEX idx_op_eventos_created_at (created_at),
                        INDEX idx_op_eventos_evento (evento),
                        INDEX idx_op_eventos_enc (encomenda_numero),
                        INDEX idx_op_eventos_peca (peca_id),
                        INDEX idx_op_eventos_operador (operador)
                    )
                    """
                )
                _mysql_runtime_schema_mark("op_eventos")
            cur.execute(
                """
                INSERT INTO op_eventos (
                    created_at, evento, encomenda_numero, peca_id, ref_interna, material, espessura,
                    operacao, operador, qtd_ok, qtd_nok, info
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    _to_mysql_datetime(ts_now),
                    _clip(evento, 30),
                    _clip(encomenda_numero, 30),
                    _clip(peca_id, 30),
                    _clip(ref_interna, 60),
                    _clip(material, 100),
                    _clip(espessura, 20),
                    _clip(operacao, 80),
                    _clip(operador, 80),
                    _to_num(qtd_ok),
                    _to_num(qtd_nok),
                    info,
                ),
            )
            evento_norm = str(evento or "").strip().upper()
            causa_txt = str(causa_paragem or "").strip()
            if evento_norm in ("PARAGEM", "STOP") and causa_txt:
                if not _mysql_runtime_schema_ready("op_paragens"):
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS op_paragens (
                            id INT AUTO_INCREMENT PRIMARY KEY,
                            created_at DATETIME NOT NULL,
                            fechada_at DATETIME NULL,
                            encomenda_numero VARCHAR(30) NULL,
                            peca_id VARCHAR(30) NULL,
                            ref_interna VARCHAR(60) NULL,
                            material VARCHAR(100) NULL,
                            espessura VARCHAR(20) NULL,
                            operador VARCHAR(80) NULL,
                            origem VARCHAR(20) NULL,
                            estado VARCHAR(20) NULL,
                            causa VARCHAR(120) NULL,
                            detalhe TEXT NULL,
                            grupo_id VARCHAR(80) NULL,
                            duracao_min DECIMAL(10,2) NULL,
                            INDEX idx_op_paragens_created_at (created_at),
                            INDEX idx_op_paragens_causa (causa),
                            INDEX idx_op_paragens_enc (encomenda_numero),
                            INDEX idx_op_paragens_peca (peca_id)
                        )
                        """
                    )
                    _mysql_runtime_schema_mark("op_paragens")
                _mysql_ensure_column(cur, "op_paragens", "grupo_id", "VARCHAR(80) NULL")
                cur.execute(
                    """
                    INSERT INTO op_paragens (
                        created_at, fechada_at, encomenda_numero, peca_id, ref_interna, material, espessura,
                        operador, origem, estado, causa, detalhe, grupo_id, duracao_min
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        _to_mysql_datetime(ts_now),
                        None,
                        _clip(encomenda_numero, 30),
                        _clip(peca_id, 30),
                        _clip(ref_interna, 60),
                        _clip(material, 100),
                        _clip(espessura, 20),
                        _clip(operador, 80),
                        "AVARIA",
                        "ABERTA",
                        _clip(causa_txt, 120),
                        info,
                        grupo_key,
                        _to_num(duracao_min),
                    ),
                )
            elif evento_norm in ("RESUME_PIECE", "START_OP", "FINISH_OP", "CLOSE_AVARIA"):
                try:
                    p_id = _clip(peca_id, 30)
                    e_num = _clip(encomenda_numero, 30)
                    if p_id and e_num:
                        if not _mysql_runtime_schema_ready("op_paragens"):
                            cur.execute(
                                """
                                CREATE TABLE IF NOT EXISTS op_paragens (
                                    id INT AUTO_INCREMENT PRIMARY KEY,
                                    created_at DATETIME NOT NULL,
                                    fechada_at DATETIME NULL,
                                    encomenda_numero VARCHAR(30) NULL,
                                    peca_id VARCHAR(30) NULL,
                                    ref_interna VARCHAR(60) NULL,
                                    material VARCHAR(100) NULL,
                                    espessura VARCHAR(20) NULL,
                                    operador VARCHAR(80) NULL,
                                    origem VARCHAR(20) NULL,
                                    estado VARCHAR(20) NULL,
                                    causa VARCHAR(120) NULL,
                                    detalhe TEXT NULL,
                                    grupo_id VARCHAR(80) NULL,
                                    duracao_min DECIMAL(10,2) NULL,
                                    INDEX idx_op_paragens_created_at (created_at),
                                    INDEX idx_op_paragens_causa (causa),
                                    INDEX idx_op_paragens_enc (encomenda_numero),
                                    INDEX idx_op_paragens_peca (peca_id)
                                )
                                """
                            )
                            _mysql_runtime_schema_mark("op_paragens")
                        _mysql_ensure_column(cur, "op_paragens", "grupo_id", "VARCHAR(80) NULL")
                        cur.execute(
                            """
                            UPDATE op_paragens p
                            JOIN (
                                SELECT id
                                FROM op_paragens
                                WHERE encomenda_numero=%s
                                  AND peca_id=%s
                                  AND (duracao_min IS NULL OR duracao_min=0)
                                ORDER BY created_at DESC, id DESC
                                LIMIT 1
                            ) u ON u.id = p.id
                            SET p.duracao_min = ROUND(TIMESTAMPDIFF(SECOND, p.created_at, %s)/60.0, 2),
                                p.fechada_at = %s,
                                p.estado = 'FECHADA'
                            """,
                            (
                                e_num,
                                p_id,
                                _to_mysql_datetime(ts_now),
                                _to_mysql_datetime(ts_now),
                            ),
                        )
                except Exception:
                    pass
        conn.commit()
    except Exception:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        if conn:
            conn.close()


def mysql_upsert_orc_referencia(
    *,
    ref_externa,
    ref_interna="",
    descricao="",
    material="",
    espessura="",
    preco_unit=None,
    operacao="",
    desenho_path="",
    tempo_peca_min=None,
    operacoes_lista=None,
    operacoes_fluxo=None,
    operacoes_detalhe=None,
    tempos_operacao=None,
    custos_operacao=None,
    quote_cost_snapshot=None,
    origem_doc="",
    origem_tipo="",
    estado_origem="",
    approved_at="",
):
    if not USE_MYSQL_STORAGE or not MYSQL_AVAILABLE:
        return
    ref_ext = _clip(ref_externa, 100)
    if not ref_ext:
        return
    normalized_desenho = ""
    if str(desenho_path or "").strip():
        try:
            normalized_desenho = lugest_storage.import_file_to_storage(
                desenho_path,
                "drawings",
                base_dir=BASE_DIR,
                preferred_name=os.path.basename(str(desenho_path or "").strip()) if str(desenho_path or "").strip() else "",
            )
        except Exception:
            normalized_desenho = str(desenho_path or "").strip()
    operacoes_json = json.dumps(
        {
            "operacoes_lista": list(operacoes_lista or []),
            "operacoes_fluxo": [dict(item or {}) for item in list(operacoes_fluxo or []) if isinstance(item, dict)],
            "operacoes_detalhe": [dict(item or {}) for item in list(operacoes_detalhe or []) if isinstance(item, dict)],
        },
        ensure_ascii=False,
    )
    tempos_json = json.dumps(dict(tempos_operacao or {}), ensure_ascii=False)
    custos_json = json.dumps(dict(custos_operacao or {}), ensure_ascii=False)
    quote_snapshot_json = json.dumps(dict(quote_cost_snapshot or {}), ensure_ascii=False)
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            if not _mysql_runtime_schema_ready("orc_referencias_historico"):
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS orc_referencias_historico (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        ref_externa VARCHAR(100) NOT NULL,
                        ref_interna VARCHAR(50) NULL,
                        descricao TEXT NULL,
                        material VARCHAR(100) NULL,
                        espessura VARCHAR(20) NULL,
                        preco_unit DECIMAL(10,2) NULL,
                        operacao VARCHAR(150) NULL,
                        tempo_peca_min DECIMAL(10,3) NULL,
                        operacoes_json LONGTEXT NULL,
                        tempos_operacao_json LONGTEXT NULL,
                        custos_operacao_json LONGTEXT NULL,
                        quote_cost_snapshot_json LONGTEXT NULL,
                        origem_doc VARCHAR(30) NULL,
                        origem_tipo VARCHAR(80) NULL,
                        estado_origem VARCHAR(80) NULL,
                        approved_at DATETIME NULL,
                        desenho_path VARCHAR(512) NULL,
                        updated_at DATETIME NOT NULL,
                        UNIQUE KEY uq_orc_ref_externa (ref_externa),
                        INDEX idx_orc_ref_interna (ref_interna),
                        INDEX idx_orc_ref_updated_at (updated_at)
                    )
                    """
                )
                _mysql_runtime_schema_mark("orc_referencias_historico")
            cur.execute(
                """
                INSERT INTO orc_referencias_historico (
                    ref_externa, ref_interna, descricao, material, espessura, preco_unit, operacao, tempo_peca_min,
                    operacoes_json, tempos_operacao_json, custos_operacao_json, quote_cost_snapshot_json,
                    origem_doc, origem_tipo, estado_origem, approved_at, desenho_path, updated_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    ref_interna=VALUES(ref_interna),
                    descricao=VALUES(descricao),
                    material=VALUES(material),
                    espessura=VALUES(espessura),
                    preco_unit=VALUES(preco_unit),
                    operacao=VALUES(operacao),
                    tempo_peca_min=VALUES(tempo_peca_min),
                    operacoes_json=VALUES(operacoes_json),
                    tempos_operacao_json=VALUES(tempos_operacao_json),
                    custos_operacao_json=VALUES(custos_operacao_json),
                    quote_cost_snapshot_json=VALUES(quote_cost_snapshot_json),
                    origem_doc=VALUES(origem_doc),
                    origem_tipo=VALUES(origem_tipo),
                    estado_origem=VALUES(estado_origem),
                    approved_at=VALUES(approved_at),
                    desenho_path=VALUES(desenho_path),
                    updated_at=VALUES(updated_at)
                """,
                (
                    ref_ext,
                    _clip(ref_interna, 50),
                    descricao,
                    _clip(material, 100),
                    _clip(espessura, 20),
                    _to_num(preco_unit),
                    _clip(operacao, 150),
                    _to_num(tempo_peca_min),
                    operacoes_json,
                    tempos_json,
                    custos_json,
                    quote_snapshot_json,
                    _clip(origem_doc, 30),
                    _clip(origem_tipo, 80),
                    _clip(estado_origem, 80),
                    _to_mysql_datetime(approved_at),
                    _clip(normalized_desenho, 512),
                    _to_mysql_datetime(now_iso()),
                ),
            )
        conn.commit()
    except Exception:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        if conn:
            conn.close()


def mysql_delete_orc_referencia(ref_externa):
    if not USE_MYSQL_STORAGE or not MYSQL_AVAILABLE:
        return
    ref_ext = _clip(ref_externa, 100)
    if not ref_ext:
        return
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            if not _mysql_runtime_schema_ready("orc_referencias_historico"):
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS orc_referencias_historico (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        ref_externa VARCHAR(100) NOT NULL,
                        ref_interna VARCHAR(50) NULL,
                        descricao TEXT NULL,
                        material VARCHAR(100) NULL,
                        espessura VARCHAR(20) NULL,
                        preco_unit DECIMAL(10,2) NULL,
                        operacao VARCHAR(150) NULL,
                        tempo_peca_min DECIMAL(10,3) NULL,
                        operacoes_json LONGTEXT NULL,
                        tempos_operacao_json LONGTEXT NULL,
                        custos_operacao_json LONGTEXT NULL,
                        quote_cost_snapshot_json LONGTEXT NULL,
                        origem_doc VARCHAR(30) NULL,
                        origem_tipo VARCHAR(80) NULL,
                        estado_origem VARCHAR(80) NULL,
                        approved_at DATETIME NULL,
                        desenho_path VARCHAR(512) NULL,
                        updated_at DATETIME NOT NULL,
                        UNIQUE KEY uq_orc_ref_externa (ref_externa),
                        INDEX idx_orc_ref_interna (ref_interna),
                        INDEX idx_orc_ref_updated_at (updated_at)
                    )
                    """
                )
                _mysql_runtime_schema_mark("orc_referencias_historico")
            cur.execute("DELETE FROM orc_referencias_historico WHERE ref_externa=%s LIMIT 1", (ref_ext,))
        conn.commit()
    except Exception:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        if conn:
            conn.close()


def mysql_upsert_orcamento_com_linhas(data, orc):
    if not USE_MYSQL_STORAGE or not MYSQL_AVAILABLE:
        return
    if not isinstance(orc, dict):
        return
    num = _clip(orc.get("numero"), 30)
    if not num:
        return
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            if not _mysql_runtime_schema_ready("orcamento_upsert_schema"):
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS orcamentos (
                        numero VARCHAR(30) PRIMARY KEY,
                        ano INT NULL,
                        data DATETIME,
                        estado VARCHAR(50),
                        cliente_codigo VARCHAR(20),
                        iva_perc DECIMAL(5,2),
                        desconto_perc DECIMAL(6,2) NULL,
                        desconto_valor DECIMAL(12,2) NULL,
                        subtotal_bruto DECIMAL(12,2) NULL,
                        subtotal DECIMAL(12,2),
                        total DECIMAL(12,2),
                        numero_encomenda VARCHAR(30),
                        nota_cliente TEXT,
                        executado_por VARCHAR(120),
                        nota_transporte TEXT,
                        notas_pdf TEXT,
                        preco_transporte DECIMAL(12,2) NULL,
                        transportadora_id VARCHAR(30) NULL,
                        transportadora_nome VARCHAR(150) NULL,
                        referencia_transporte VARCHAR(80) NULL,
                        zona_transporte VARCHAR(120) NULL,
                        posto_trabalho VARCHAR(80) NULL,
                        meta_json LONGTEXT NULL
                    )
                    """
                )
                _mysql_ensure_column(cur, "orcamentos", "ano", "INT NULL")
                _mysql_ensure_column(cur, "orcamentos", "executado_por", "VARCHAR(120) NULL")
                _mysql_ensure_column(cur, "orcamentos", "nota_transporte", "TEXT NULL")
                _mysql_ensure_column(cur, "orcamentos", "notas_pdf", "TEXT NULL")
                _mysql_ensure_column(cur, "orcamentos", "desconto_perc", "DECIMAL(6,2) NULL")
                _mysql_ensure_column(cur, "orcamentos", "desconto_valor", "DECIMAL(12,2) NULL")
                _mysql_ensure_column(cur, "orcamentos", "subtotal_bruto", "DECIMAL(12,2) NULL")
                _mysql_ensure_column(cur, "orcamentos", "preco_transporte", "DECIMAL(12,2) NULL")
                _mysql_ensure_column(cur, "orcamentos", "transportadora_id", "VARCHAR(30) NULL")
                _mysql_ensure_column(cur, "orcamentos", "transportadora_nome", "VARCHAR(150) NULL")
                _mysql_ensure_column(cur, "orcamentos", "referencia_transporte", "VARCHAR(80) NULL")
                _mysql_ensure_column(cur, "orcamentos", "zona_transporte", "VARCHAR(120) NULL")
                _mysql_ensure_column(cur, "orcamentos", "posto_trabalho", "VARCHAR(80) NULL")
                _mysql_ensure_column(cur, "orcamentos", "meta_json", "LONGTEXT NULL")
                _mysql_ensure_index(cur, "orcamentos", "idx_orcamentos_ano", "ano")
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS orcamento_linhas (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        orcamento_numero VARCHAR(30),
                        ref_interna VARCHAR(50),
                        ref_externa VARCHAR(100),
                        descricao TEXT,
                        material VARCHAR(100),
                        espessura VARCHAR(20),
                        operacao VARCHAR(150),
                        of_codigo VARCHAR(30),
                        qtd DECIMAL(10,2),
                        preco_unit DECIMAL(10,2),
                        tempo_peca_min DECIMAL(10,2),
                        operacoes_json LONGTEXT NULL,
                        tempos_operacao_json LONGTEXT NULL,
                        custos_operacao_json LONGTEXT NULL,
                        quote_cost_snapshot_json LONGTEXT NULL,
                        total DECIMAL(12,2),
                        desenho_path VARCHAR(512),
                        meta_json LONGTEXT NULL,
                        tipo_item VARCHAR(30) NULL,
                        produto_codigo VARCHAR(20) NULL,
                        produto_unid VARCHAR(20) NULL,
                        conjunto_codigo VARCHAR(40) NULL,
                        conjunto_nome VARCHAR(150) NULL,
                        grupo_uuid VARCHAR(60) NULL,
                        qtd_base DECIMAL(10,2) NULL,
                        INDEX idx_orc_linhas_orcamento_numero (orcamento_numero)
                    )
                    """
                )
                _mysql_ensure_column(cur, "orcamento_linhas", "tempo_peca_min", "DECIMAL(10,2) NULL")
                _mysql_ensure_column(cur, "orcamento_linhas", "tipo_item", "VARCHAR(30) NULL")
                _mysql_ensure_column(cur, "orcamento_linhas", "produto_codigo", "VARCHAR(20) NULL")
                _mysql_ensure_column(cur, "orcamento_linhas", "produto_unid", "VARCHAR(20) NULL")
                _mysql_ensure_column(cur, "orcamento_linhas", "conjunto_codigo", "VARCHAR(40) NULL")
                _mysql_ensure_column(cur, "orcamento_linhas", "conjunto_nome", "VARCHAR(150) NULL")
                _mysql_ensure_column(cur, "orcamento_linhas", "grupo_uuid", "VARCHAR(60) NULL")
                _mysql_ensure_column(cur, "orcamento_linhas", "qtd_base", "DECIMAL(10,2) NULL")
                _mysql_ensure_column(cur, "orcamento_linhas", "meta_json", "LONGTEXT NULL")
                _mysql_runtime_schema_mark("orcamento_upsert_schema")

            cliente_cod = _extract_cliente_codigo(orc.get("cliente"), data)
            if cliente_cod:
                cur.execute("SELECT 1 FROM clientes WHERE codigo=%s LIMIT 1", (cliente_cod,))
                if not cur.fetchone():
                    cliente_cod = None

            cur.execute(
                """
                INSERT INTO orcamentos (
                    numero, ano, data, estado, cliente_codigo, iva_perc, subtotal, total, numero_encomenda, nota_cliente,
                    executado_por, nota_transporte, notas_pdf, desconto_perc, desconto_valor, subtotal_bruto,
                    preco_transporte, transportadora_id, transportadora_nome,
                    referencia_transporte, zona_transporte, posto_trabalho, meta_json
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    ano=VALUES(ano),
                    data=VALUES(data),
                    estado=VALUES(estado),
                    cliente_codigo=VALUES(cliente_codigo),
                    iva_perc=VALUES(iva_perc),
                    subtotal=VALUES(subtotal),
                    total=VALUES(total),
                    numero_encomenda=VALUES(numero_encomenda),
                    nota_cliente=VALUES(nota_cliente),
                    executado_por=VALUES(executado_por),
                    nota_transporte=VALUES(nota_transporte),
                    notas_pdf=VALUES(notas_pdf),
                    desconto_perc=VALUES(desconto_perc),
                    desconto_valor=VALUES(desconto_valor),
                    subtotal_bruto=VALUES(subtotal_bruto),
                    preco_transporte=VALUES(preco_transporte),
                    transportadora_id=VALUES(transportadora_id),
                    transportadora_nome=VALUES(transportadora_nome),
                    referencia_transporte=VALUES(referencia_transporte),
                    zona_transporte=VALUES(zona_transporte),
                    posto_trabalho=VALUES(posto_trabalho),
                    meta_json=VALUES(meta_json)
                """,
                (
                    num,
                    _derive_year_from_values(orc.get("data"), orc.get("numero"), default=datetime.now().year),
                    _to_mysql_datetime(orc.get("data") or now_iso()),
                    _clip(orc.get("estado"), 50),
                    _clip(cliente_cod, 20) if cliente_cod else None,
                    _to_num(orc.get("iva_perc")),
                    _to_num(orc.get("subtotal")),
                    _to_num(orc.get("total")),
                    _clip(orc.get("numero_encomenda"), 30),
                    orc.get("nota_cliente"),
                    _clip(orc.get("executado_por"), 120),
                    orc.get("nota_transporte"),
                    orc.get("notas_pdf"),
                    _to_num(orc.get("desconto_perc")),
                    _to_num(orc.get("desconto_valor")),
                    _to_num(orc.get("subtotal_bruto")),
                    _to_num(orc.get("preco_transporte")),
                    _clip(orc.get("transportadora_id"), 30),
                    _clip(orc.get("transportadora_nome"), 150),
                    _clip(orc.get("referencia_transporte"), 80),
                    _clip(orc.get("zona_transporte"), 120),
                    _clip(orc.get("posto_trabalho"), 80),
                    json.dumps(
                        {
                            key: orc.get(key)
                            for key in (
                                "desconto_modo",
                                "desconto_grupos",
                            )
                            if orc.get(key) not in (None, "", [], {})
                        },
                        ensure_ascii=False,
                    ),
                ),
            )

            cur.execute("DELETE FROM orcamento_linhas WHERE orcamento_numero=%s", (num,))
            for l in orc.get("linhas", []):
                if not isinstance(l, dict):
                    continue
                cur.execute(
                    """
                    INSERT INTO orcamento_linhas (
                        orcamento_numero, ref_interna, ref_externa, descricao, material, espessura, operacao,
                        of_codigo, qtd, preco_unit, tempo_peca_min, operacoes_json, tempos_operacao_json,
                        custos_operacao_json, quote_cost_snapshot_json, total, desenho_path, meta_json, tipo_item,
                        produto_codigo, produto_unid, conjunto_codigo, conjunto_nome, grupo_uuid, qtd_base
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        num,
                        _clip(l.get("ref_interna"), 50),
                        _clip(l.get("ref_externa"), 100),
                        l.get("descricao"),
                        _clip(l.get("material"), 100),
                        _clip(l.get("espessura"), 20),
                        _clip(l.get("operacao", l.get("Operacoes", l.get("Operações", ""))), 150),
                        _clip(l.get("of"), 30),
                        _to_num(l.get("qtd")),
                        _to_num(l.get("preco_unit")),
                        _to_num(l.get("tempo_peca_min", l.get("tempo_pecas_min"))),
                        json.dumps(
                            {
                                "operacoes_lista": list(l.get("operacoes_lista", []) or []),
                                "operacoes_fluxo": [dict(item or {}) for item in list(l.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                                "operacoes_detalhe": [dict(item or {}) for item in list(l.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                            },
                            ensure_ascii=False,
                        ),
                        json.dumps(dict(l.get("tempos_operacao", {}) or {}), ensure_ascii=False),
                        json.dumps(dict(l.get("custos_operacao", {}) or {}), ensure_ascii=False),
                        json.dumps(dict(l.get("quote_cost_snapshot", {}) or {}), ensure_ascii=False),
                        _to_num(l.get("total")),
                        _clip(l.get("desenho"), 512),
                        json.dumps(
                            {
                                key: l.get(key)
                                for key in (
                                    "stock_material_id",
                                    "price_per_kg",
                                    "price_base_value",
                                    "price_markup_pct",
                                    "stock_metric_value",
                                    "kg_per_m",
                                    "meters_per_unit",
                                    "weight_total",
                                    "quantity_units",
                                    "calc_mode",
                                    "descricao_base",
                                    "length_mm",
                                    "width_mm",
                                    "thickness_mm",
                                    "density",
                                    "diameter_mm",
                                    "manual_unit_price",
                                    "profile_section",
                                    "profile_size",
                                    "tube_section",
                                    "quality",
                                    "hint",
                                    "price_base_label",
                                    "material_family",
                                    "material_subtype",
                                )
                                if l.get(key) not in (None, "", [], {})
                            },
                            ensure_ascii=False,
                        ),
                        _clip(normalize_orc_line_type(l.get("tipo_item")), 30),
                        _clip(l.get("produto_codigo"), 20),
                        _clip(l.get("produto_unid"), 20),
                        _clip(l.get("conjunto_codigo"), 40),
                        _clip(l.get("conjunto_nome"), 150),
                        _clip(l.get("grupo_uuid"), 60),
                        _to_num(l.get("qtd_base", l.get("qtd"))),
                    ),
                )
        conn.commit()
    except Exception:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        raise
    finally:
        if conn:
            conn.close()


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


TRIAL_CONFIG_FILE = "lugest_trial.json"
TRIAL_CONFIG_KEY = "trial_license"
SUPPLIER_SEQ_CONFIG_FILE = "lugest_supplier_seq.json"
SUPPLIER_SEQ_CONFIG_KEY = "supplier_sequence"
TRANSPORT_SEQ_CONFIG_FILE = "lugest_transport_seq.json"
TRANSPORT_SEQ_CONFIG_KEY = "transport_sequence"
RUNTIME_STATE_CONFIG_FILE = "lugest_runtime_state.json"
RUNTIME_STATE_CONFIG_KEY = "runtime_state"
RUNTIME_STATE_CONFIG_DEFAULTS = {
    "postos_trabalho": [],
    "operador_posto_map": {},
    "tempos_operacao_planeada_min": {},
    "workcenter_catalog": [],
    "plano_bloqueios": [],
}
QUALITY_RUNTIME_CONFIG_DEFAULTS = {
    "quality_nonconformities": [],
    "quality_documents": [],
    "audit_log": [],
}


def _app_config_json_path(filename):
    return os.path.join(BASE_DIR, str(filename or "").strip())


def _app_config_load_json(config_key, filename, conn=None):
    payload = {}
    active_conn = conn
    own_conn = False
    try:
        if USE_MYSQL_STORAGE and MYSQL_AVAILABLE:
            if active_conn is None:
                active_conn = _mysql_connect()
                own_conn = True
            with active_conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS app_config (
                        ckey VARCHAR(80) PRIMARY KEY,
                        cvalue LONGTEXT NULL,
                        updated_at DATETIME NULL
                    )
                    """
                )
                cur.execute("SELECT cvalue FROM app_config WHERE ckey=%s LIMIT 1", (str(config_key or "").strip(),))
                row = cur.fetchone()
            if row:
                raw = row.get("cvalue") if isinstance(row, dict) else row[0]
                if isinstance(raw, (bytes, bytearray)):
                    raw = raw.decode("utf-8", errors="ignore")
                parsed = json.loads(str(raw or "{}"))
                if isinstance(parsed, dict):
                    payload = parsed
    except Exception:
        payload = {}
    finally:
        try:
            if own_conn and active_conn:
                active_conn.close()
        except Exception:
            pass
    if payload:
        return dict(payload)
    try:
        path = _app_config_json_path(filename)
        if os.path.exists(path):
            parsed = json.loads(open(path, "r", encoding="utf-8").read())
            if isinstance(parsed, dict):
                return dict(parsed)
    except Exception:
        pass
    return {}


def _app_config_save_json(config_key, filename, payload, conn=None):
    clean = dict(payload or {})
    try:
        path = _app_config_json_path(filename)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(json.dumps(clean, ensure_ascii=False, indent=2))
    except Exception:
        pass
    active_conn = conn
    own_conn = False
    try:
        if USE_MYSQL_STORAGE and MYSQL_AVAILABLE:
            if active_conn is None:
                active_conn = _mysql_connect()
                own_conn = True
            with active_conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS app_config (
                        ckey VARCHAR(80) PRIMARY KEY,
                        cvalue LONGTEXT NULL,
                        updated_at DATETIME NULL
                    )
                    """
                )
                cur.execute(
                    """
                    INSERT INTO app_config (ckey, cvalue, updated_at)
                    VALUES (%s, %s, NOW())
                    ON DUPLICATE KEY UPDATE cvalue=VALUES(cvalue), updated_at=VALUES(updated_at)
                    """,
                    (str(config_key or "").strip(), json.dumps(clean, ensure_ascii=False)),
                )
            if own_conn:
                active_conn.commit()
    finally:
        try:
            if own_conn and active_conn:
                active_conn.close()
        except Exception:
            pass
    return dict(clean)


def _runtime_state_payload(data):
    payload = {}
    source = dict(data or {})
    for key, default in RUNTIME_STATE_CONFIG_DEFAULTS.items():
        value = copy.deepcopy(source.get(key, default))
        if isinstance(default, list):
            payload[key] = value if isinstance(value, list) else copy.deepcopy(default)
        elif isinstance(default, dict):
            payload[key] = value if isinstance(value, dict) else copy.deepcopy(default)
        else:
            payload[key] = copy.deepcopy(default)
    try:
        return json.loads(json.dumps(payload, ensure_ascii=False))
    except Exception:
        clean = {}
        for key, default in RUNTIME_STATE_CONFIG_DEFAULTS.items():
            clean[key] = copy.deepcopy(default)
        return clean


def _apply_quality_runtime_payload(data, payload=None):
    if not isinstance(data, dict):
        return data
    for key, default in QUALITY_RUNTIME_CONFIG_DEFAULTS.items():
        value = copy.deepcopy(data.get(key, default))
        data[key] = value if isinstance(value, list) else copy.deepcopy(default)
    return data


def _apply_runtime_state_payload(data, payload):
    if not isinstance(data, dict):
        return
    source = dict(payload or {}) if isinstance(payload, dict) else {}
    for key, default in RUNTIME_STATE_CONFIG_DEFAULTS.items():
        value = source.get(key, copy.deepcopy(default))
        if isinstance(default, list):
            data[key] = copy.deepcopy(value) if isinstance(value, list) else copy.deepcopy(default)
        elif isinstance(default, dict):
            data[key] = copy.deepcopy(value) if isinstance(value, dict) else copy.deepcopy(default)
        else:
            data[key] = copy.deepcopy(default)


def _extract_fornecedor_seq(value):
    raw = str(value or "").strip().upper()
    if raw.startswith("FOR-"):
        raw = raw[4:]
    return int(raw) if raw.isdigit() else 0


def _load_fornecedor_sequence_next(data):
    seq = data.setdefault("seq", {})
    current = int(seq.get("fornecedor", 1) or 1)
    existing = max((_extract_fornecedor_seq((row or {}).get("id")) for row in list(data.get("fornecedores", []) or [])), default=0) + 1
    stored = _app_config_load_json(SUPPLIER_SEQ_CONFIG_KEY, SUPPLIER_SEQ_CONFIG_FILE)
    try:
        stored_next = int(dict(stored or {}).get("next", 0) or 0)
    except Exception:
        stored_next = 0
    next_n = max(1, current, existing, stored_next)
    seq["fornecedor"] = next_n
    return next_n


def _store_fornecedor_sequence_next(data, next_n):
    seq = data.setdefault("seq", {})
    value = max(1, int(next_n or 1))
    seq["fornecedor"] = value
    _app_config_save_json(SUPPLIER_SEQ_CONFIG_KEY, SUPPLIER_SEQ_CONFIG_FILE, {"next": value})


def _extract_transport_seq(value):
    raw = str(value or "").strip().upper()
    if raw.startswith("TR-"):
        raw = raw[3:]
    parts = [chunk for chunk in raw.split("-") if chunk]
    if parts and parts[-1].isdigit():
        try:
            return int(parts[-1])
        except Exception:
            return 0
    digits = "".join(ch for ch in raw if ch.isdigit())
    return int(digits[-4:]) if len(digits) >= 4 and digits[-4:].isdigit() else 0


def _load_transport_sequence_next(data):
    seq = data.setdefault("seq", {})
    current = int(seq.get("transporte", 1) or 1)
    existing = max((_extract_transport_seq((row or {}).get("numero")) for row in list(data.get("transportes", []) or [])), default=0) + 1
    stored = _app_config_load_json(TRANSPORT_SEQ_CONFIG_KEY, TRANSPORT_SEQ_CONFIG_FILE)
    try:
        stored_next = int(dict(stored or {}).get("next", 0) or 0)
    except Exception:
        stored_next = 0
    next_n = max(1, current, existing, stored_next)
    seq["transporte"] = next_n
    return next_n


def _store_transport_sequence_next(data, next_n):
    seq = data.setdefault("seq", {})
    value = max(1, int(next_n or 1))
    seq["transporte"] = value
    _app_config_save_json(TRANSPORT_SEQ_CONFIG_KEY, TRANSPORT_SEQ_CONFIG_FILE, {"next": value})


def current_machine_fingerprint():
    raw = "|".join(
        [
            str(platform.system() or "").strip().lower(),
            str(platform.machine() or "").strip().lower(),
            str(platform.node() or os.environ.get("COMPUTERNAME", "") or "").strip().lower(),
            str(uuid.getnode() or "").strip().lower(),
        ]
    )
    digest = hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest().upper()
    return f"{digest[:4]}-{digest[4:8]}-{digest[8:12]}-{digest[12:16]}"


def trial_owner_username():
    return str(os.environ.get("LUGEST_OWNER_USERNAME", "") or "").strip()


def trial_owner_password():
    return str(os.environ.get("LUGEST_OWNER_PASSWORD", "") or "")


def trial_owner_configured():
    return bool(trial_owner_username() and trial_owner_password())


def verify_owner_credentials(username, password):
    owner_username = trial_owner_username()
    owner_password = trial_owner_password()
    if not owner_username or not owner_password:
        return False
    return str(username or "").strip().lower() == owner_username.lower() and verify_password(password, owner_password)


def owner_session_user(username="", password=""):
    owner_username = str(username or "").strip() or trial_owner_username() or "owner"
    return {
        "username": owner_username,
        "password": "",
        "_session_password": str(password or ""),
        "role": "Admin",
        "owner_session": True,
        "active": True,
        "menu_permissions": {},
    }


def _parse_trial_dt(raw_value):
    raw = str(raw_value or "").strip()
    if not raw:
        return None
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except Exception:
        return None


def _trial_config_defaults():
    return {
        "enabled": False,
        "company_name": "",
        "device_fingerprint": current_machine_fingerprint(),
        "started_at": "",
        "duration_days": 60,
        "created_at": "",
        "created_by": "",
        "updated_at": "",
        "updated_by": "",
        "last_success_at": "",
        "last_success_user": "",
        "last_owner_auth_at": "",
        "last_owner_auth_user": "",
        "notes": "",
    }


def load_trial_config():
    cfg = _trial_config_defaults()
    stored = _app_config_load_json(TRIAL_CONFIG_KEY, TRIAL_CONFIG_FILE)
    if isinstance(stored, dict):
        cfg.update(stored)
    try:
        cfg["duration_days"] = max(1, int(cfg.get("duration_days", 60) or 60))
    except Exception:
        cfg["duration_days"] = 60
    cfg["enabled"] = bool(cfg.get("enabled", False))
    cfg["company_name"] = str(cfg.get("company_name", "") or "").strip()
    cfg["device_fingerprint"] = str(cfg.get("device_fingerprint", "") or "").strip() or current_machine_fingerprint()
    for key in (
        "started_at",
        "created_at",
        "created_by",
        "updated_at",
        "updated_by",
        "last_success_at",
        "last_success_user",
        "last_owner_auth_at",
        "last_owner_auth_user",
        "notes",
    ):
        cfg[key] = str(cfg.get(key, "") or "").strip()
    return cfg


def save_trial_config(payload):
    cfg = _trial_config_defaults()
    cfg.update(dict(payload or {}))
    try:
        cfg["duration_days"] = max(1, int(cfg.get("duration_days", 60) or 60))
    except Exception:
        cfg["duration_days"] = 60
    cfg["enabled"] = bool(cfg.get("enabled", False))
    cfg["company_name"] = str(cfg.get("company_name", "") or "").strip()
    cfg["device_fingerprint"] = str(cfg.get("device_fingerprint", "") or "").strip() or current_machine_fingerprint()
    for key in (
        "started_at",
        "created_at",
        "created_by",
        "updated_at",
        "updated_by",
        "last_success_at",
        "last_success_user",
        "last_owner_auth_at",
        "last_owner_auth_user",
        "notes",
    ):
        cfg[key] = str(cfg.get(key, "") or "").strip()
    return _app_config_save_json(TRIAL_CONFIG_KEY, TRIAL_CONFIG_FILE, cfg)


def get_trial_status(now_dt=None):
    cfg = load_trial_config()
    now_obj = now_dt if isinstance(now_dt, datetime) else datetime.now()
    current_fingerprint = current_machine_fingerprint()
    started_dt = _parse_trial_dt(cfg.get("started_at"))
    expires_dt = None
    expired = False
    device_mismatch = False
    blocking = False
    state = "disabled"
    message = "Trial desativado."
    days_remaining = None
    if bool(cfg.get("enabled", False)):
        state = "active"
        if started_dt is None:
            blocking = True
            state = "invalid"
            message = "Trial invalido. Falta a data de inicio."
        else:
            expires_dt = started_dt + timedelta(days=max(1, int(cfg.get("duration_days", 60) or 60)))
            device_mismatch = str(cfg.get("device_fingerprint", "") or "").strip() not in ("", current_fingerprint)
            expired = now_obj > expires_dt
            days_remaining = max(0, (expires_dt.date() - now_obj.date()).days)
            if device_mismatch:
                blocking = True
                state = "device_mismatch"
                message = "Licenca bloqueada: o trial foi criado para outro equipamento."
            elif expired:
                blocking = True
                state = "expired"
                message = "Trial expirado. So o login do proprietario pode autorizar nova utilizacao."
            else:
                message = f"Trial ativo ate {expires_dt.strftime('%d/%m/%Y %H:%M')}."
    return {
        **cfg,
        "state": state,
        "blocking": bool(blocking),
        "expired": bool(expired),
        "device_mismatch": bool(device_mismatch),
        "current_device_fingerprint": current_fingerprint,
        "started_at": started_dt.isoformat(timespec="seconds") if started_dt else "",
        "expires_at": expires_dt.isoformat(timespec="seconds") if expires_dt else "",
        "days_remaining": days_remaining,
        "owner_username": trial_owner_username(),
        "owner_configured": trial_owner_configured(),
        "message": message,
    }


def activate_trial_license(company_name="", duration_days=60, created_by="", notes="", reset_start=True):
    if not trial_owner_configured():
        raise ValueError("Define LUGEST_OWNER_USERNAME e LUGEST_OWNER_PASSWORD no lugest.env antes de ativar um trial.")
    current = load_trial_config()
    start_txt = now_iso() if reset_start or not str(current.get("started_at", "") or "").strip() else str(current.get("started_at", "") or "").strip()
    payload = {
        **current,
        "enabled": True,
        "company_name": str(company_name or current.get("company_name", "") or "").strip(),
        "device_fingerprint": current_machine_fingerprint(),
        "started_at": start_txt,
        "duration_days": max(1, int(duration_days or current.get("duration_days", 60) or 60)),
        "created_at": str(current.get("created_at", "") or "").strip() or now_iso(),
        "created_by": str(created_by or current.get("created_by", "") or "").strip(),
        "updated_at": now_iso(),
        "updated_by": str(created_by or "").strip(),
        "notes": str(notes or current.get("notes", "") or "").strip(),
    }
    save_trial_config(payload)
    return get_trial_status()


def extend_trial_license(extra_days=30, updated_by=""):
    current = load_trial_config()
    if not bool(current.get("enabled", False)):
        raise ValueError("Nao existe um trial ativo para prolongar.")
    try:
        extra_num = max(1, int(extra_days or 0))
    except Exception:
        extra_num = 30
    current["duration_days"] = max(1, int(current.get("duration_days", 60) or 60)) + extra_num
    current["updated_at"] = now_iso()
    current["updated_by"] = str(updated_by or "").strip()
    current["last_owner_auth_at"] = now_iso()
    current["last_owner_auth_user"] = str(updated_by or "").strip()
    save_trial_config(current)
    return get_trial_status()


def disable_trial_license(updated_by=""):
    current = load_trial_config()
    current["enabled"] = False
    current["updated_at"] = now_iso()
    current["updated_by"] = str(updated_by or "").strip()
    save_trial_config(current)
    return get_trial_status()


def touch_trial_success(username="", owner=False):
    current = load_trial_config()
    now_txt = now_iso()
    current["last_success_at"] = now_txt
    current["last_success_user"] = str(username or "").strip()
    if owner:
        current["last_owner_auth_at"] = now_txt
        current["last_owner_auth_user"] = str(username or "").strip()
    save_trial_config(current)
    return get_trial_status()


def ensure_trial_login_session(username="", password="", allow_owner=True):
    if allow_owner and verify_owner_credentials(username, password):
        session_user = owner_session_user(username, password=password)
        touch_trial_success(session_user.get("username", ""), owner=True)
        return session_user
    status = get_trial_status()
    if bool(status.get("blocking", False)):
        raise ValueError(str(status.get("message", "") or "Licenca bloqueada."))
    return None


def ensure_trial_runtime_access():
    status = get_trial_status()
    if bool(status.get("blocking", False)):
        raise RuntimeError(str(status.get("message", "") or "Licenca bloqueada."))
    return status


def normalize_produto_mov_row(row):
    if not isinstance(row, dict):
        return None
    data_txt = str(
        row.get("data")
        or row.get("created_at")
        or row.get("when")
        or now_iso()
    ).strip()
    tipo = str(row.get("tipo", "") or "").strip()
    operador = str(
        row.get("operador")
        or row.get("user")
        or row.get("utilizador")
        or ""
    ).strip()
    codigo = str(row.get("codigo") or row.get("produto") or "").strip()
    descricao = str(row.get("descricao", "") or "").strip()
    qtd = parse_float(row.get("qtd", row.get("quantidade", 0)), 0.0)

    before_raw = row.get("antes")
    after_raw = row.get("depois")
    before = parse_float(before_raw, 0.0) if before_raw is not None else 0.0
    after = parse_float(after_raw, 0.0) if after_raw is not None else 0.0

    obs = str(row.get("obs", row.get("observacoes", "")) or "").strip()
    origem = str(row.get("origem", row.get("source", "")) or "").strip()
    ref_doc = str(row.get("ref_doc", row.get("documento", "")) or "").strip()

    return {
        "data": data_txt,
        "tipo": tipo,
        "operador": operador,
        "codigo": codigo,
        "descricao": descricao,
        "qtd": qtd,
        "antes": before,
        "depois": after,
        "obs": obs,
        "origem": origem,
        "ref_doc": ref_doc,
    }


def add_produto_mov(
    data,
    *,
    tipo="",
    operador="",
    codigo="",
    descricao="",
    qtd=0.0,
    antes=0.0,
    depois=0.0,
    obs="",
    origem="",
    ref_doc="",
    data_mov=None,
):
    mov = normalize_produto_mov_row(
        {
            "data": data_mov or now_iso(),
            "tipo": tipo,
            "operador": operador,
            "codigo": codigo,
            "descricao": descricao,
            "qtd": qtd,
            "antes": antes,
            "depois": depois,
            "obs": obs,
            "origem": origem,
            "ref_doc": ref_doc,
        }
    )
    if mov is None:
        return None
    data.setdefault("produtos_mov", []).append(mov)
    return mov


def _temp_pdf_path(prefix, doc_id=None):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if doc_id:
        return os.path.join(tempfile.gettempdir(), f"{prefix}_{doc_id}_{ts}.pdf")
    return os.path.join(tempfile.gettempdir(), f"{prefix}_{ts}.pdf")


def iso_diff_minutes(start_iso, end_iso=None):
    if not start_iso:
        return None
    try:
        s = datetime.fromisoformat(str(start_iso))
        e = datetime.fromisoformat(str(end_iso)) if end_iso else datetime.now()
        diff = (e - s).total_seconds() / 60.0
        if diff < 0:
            diff = 0
        return round(diff, 2)
    except Exception:
        return None


def load_data():
    if not USE_MYSQL_STORAGE:
        raise RuntimeError("A aplicação está configurada para usar apenas MySQL.")
    try:
        data = _mysql_load_relational_data()
    except Exception as ex:
        raise RuntimeError(format_mysql_runtime_error(ex, action="carregar dados da base MySQL"))
    data = _repair_mojibake_structure(data)
    if not isinstance(data, dict):
        data = _copy_default_data()
    migration_required = False
    raw_users = data.get("users", [])
    if not isinstance(raw_users, list):
        raw_users = []
        migration_required = True
    data["users"] = raw_users
    users_clean = []
    seen_usernames = set()
    for u in data.get("users", []):
        if not isinstance(u, dict):
            continue
        un = _clip(u.get("username"), 50)
        if not un:
            continue
        key = un.lower()
        if key in seen_usernames:
            migration_required = True
            continue
        seen_usernames.add(key)
        stored_password = _clip(u.get("password"), 255) or ""
        normalized_password = _clip(normalize_password_for_storage(un, stored_password, require_strong=False), 255)
        if normalized_password != stored_password:
            migration_required = True
        users_clean.append(
            {
                "username": un,
                "password": normalized_password,
                "role": _normalize_role_name(_clip(u.get("role"), 50)),
            }
        )
    data["users"] = users_clean
    existing = {str(u.get("username", "")).strip().lower() for u in data.get("users", [])}
    if not existing:
        for u in bootstrap_user_payloads():
            un = _clip(u.get("username"), 50)
            if not un:
                continue
            if un.lower() not in existing:
                data["users"].append(
                    {
                        "username": un,
                        "password": _clip(u.get("password"), 255),
                        "role": _normalize_role_name(_clip(u.get("role"), 50)),
                    }
                )
                existing.add(un.lower())
                migration_required = True
    data.setdefault("seq", {"encomenda": 1, "cliente": 1, "ref_interna": {}, "produto": 1, "ne": 1, "fornecedor": 1})
    data["seq"].setdefault("encomenda", 1)
    data["seq"].setdefault("cliente", 1)
    data["seq"].setdefault("ref_interna", {})
    data["seq"].setdefault("produto", 1)
    data["seq"].setdefault("ne", 1)
    data["seq"].setdefault("fornecedor", 1)
    data.setdefault("clientes", [])
    data.setdefault("materiais", [])
    data.setdefault("encomendas", [])
    data.setdefault("qualidade", [])
    data.setdefault("refs", [])
    data.setdefault("materiais_hist", [])
    data.setdefault("espessuras_hist", [])
    data.setdefault("stock_log", [])
    data.setdefault("plano", [])
    data.setdefault("peca_hist", {})
    data.setdefault("rejeitadas_hist", [])
    data.setdefault("operadores", ["Operador 1"])
    data.setdefault("orcamentos", [])
    data.setdefault("orc_seq", 1)
    data.setdefault("orc_refs", {})
    data.setdefault("conjuntos_modelo", [])
    data.setdefault("conjuntos", [])
    data.setdefault("of_seq", 1)
    data.setdefault("opp_seq", 1)
    data.setdefault("orcamentistas", ["Orçamentista 1"])
    # novos modulos
    data["seq"].setdefault("produto", 1)
    data["seq"].setdefault("ne", 1)
    data.setdefault("produtos", [])
    data.setdefault("notas_encomenda", [])
    data.setdefault("expedicoes", [])
    data.setdefault("transportes", [])
    data.setdefault("transportes_tarifarios", [])
    data.setdefault("at_series", [])
    data.setdefault("exp_seq", 1)
    data.setdefault("fornecedores", [])
    data.setdefault("produtos_mov", [])
    data.setdefault("plano_hist", [])
    _apply_runtime_state_payload(data, data)
    _apply_quality_runtime_payload(data)
    normalized_orcamentos = []
    for o in data.get("orcamentos", []):
        if not isinstance(o, dict):
            continue
        o["cliente"] = _normalize_orc_cliente(o.get("cliente"), data)
        o.setdefault("zona_transporte", "")
        if isinstance(o.get("notas_pdf"), list):
            o["notas_pdf"] = "\n".join([str(x).strip() for x in o.get("notas_pdf", []) if str(x).strip()])
        else:
            o["notas_pdf"] = str(o.get("notas_pdf", "") or "")
        o.setdefault("linhas", [])
        o.setdefault("executado_por", "")
        linhas = o.get("linhas", [])
        if not isinstance(linhas, list):
            linhas = []
        fixed_linhas = []
        for l in linhas:
            if not isinstance(l, dict):
                continue
            l.setdefault("ref_interna", l.get("ref", ""))
            l.setdefault("ref_externa", l.get("ref", ""))
            l.setdefault("operacao", l.get("Operacoes", l.get("Operações", "")))
            if not l.get("of"):
                l["of"] = next_of_numero(data)
            fixed_linhas.append(l)
        o["linhas"] = fixed_linhas
        normalized_orcamentos.append(o)
    data["orcamentos"] = normalized_orcamentos
    for c in data["clientes"]:
        c.setdefault("prazo_entrega", "")
        c.setdefault("cond_pagamento", "")
        c.setdefault("obs_tecnicas", "")
    for f in data.get("fornecedores", []):
        f.setdefault("codigo_postal", "")
        f.setdefault("localidade", "")
        f.setdefault("pais", "Portugal")
        f.setdefault("cond_pagamento", "")
        f.setdefault("prazo_entrega_dias", "")
        f.setdefault("website", "")
        f.setdefault("obs", "")
    for m in data["materiais"]:
        m.setdefault("reservado", 0.0)
        m.setdefault("is_sobra", False)
        m.setdefault("Localizacao", "")
        m.setdefault("lote_fornecedor", "")
        m.setdefault("formato", detect_materia_formato(m))
    for e in data["encomendas"]:
        e.setdefault("tempo", e.get("tempo_estimado", 0.0))
        e.setdefault("tempo_estimado", e.get("tempo", 0.0))
        e.setdefault("nota_cliente", "")
        e.setdefault("zona_transporte", "")
        e.setdefault("reservas", [])
        e.setdefault("inicio_encomenda", "")
        e.setdefault("fim_encomenda", "")
        e.setdefault("estado_operador", "")
        e.setdefault("obs_inicio", "")
        e.setdefault("obs_interrupcao", "")
        e.setdefault("tempo_por_espessura", {})
        e.setdefault("materiais", [])
        e.setdefault("espessuras", [])
        for p in e.get("pecas", []):
            p.setdefault("Operacoes", p.get("Operações", ""))
            p.setdefault("inicio_producao", "")
            p.setdefault("fim_producao", "")
            p.setdefault("produzido_qualidade", 0.0)
            if not p.get("of"):
                p["of"] = next_of_numero(data)
            if not p.get("opp"):
                p["opp"] = next_opp_numero(data)
            ensure_peca_operacoes(p)
        if e.get("pecas") and not e.get("materiais"):
            mat_map = {}
            for p in e.get("pecas", []):
                mat = p.get("material", "")
                esp = p.get("espessura", "")
                mat_map.setdefault(mat, {"material": mat, "estado": "Preparacao", "espessuras": {}})
                mat_map[mat]["espessuras"].setdefault(
                    esp, {"espessura": esp, "tempo_min": "", "estado": "Preparacao", "pecas": []}
                )
                mat_map[mat]["espessuras"][esp]["pecas"].append(p)
            e["materiais"] = []
            for m in mat_map.values():
                m["espessuras"] = list(m["espessuras"].values())
                e["materiais"].append(m)
        # garantir Ordem de Fabrico nas pecas em materiais/espessuras
        for idx_p, p in enumerate(encomenda_pecas(e), start=1):
            if not p.get("id"):
                p["id"] = f"{str(e.get('numero','ENC'))}-{idx_p:03d}"
            if not p.get("opp"):
                p["opp"] = next_opp_numero(data)
            ensure_peca_operacoes(p)
            atualizar_estado_peca(p)
        e.setdefault("estado_expedicao", "Não expedida")
    for tr in data.get("transportes", []):
        if not isinstance(tr, dict):
            continue
        tr.setdefault("paragens", [])
        for stop in tr.get("paragens", []):
            if isinstance(stop, dict):
                stop.setdefault("zona_transporte", "")
    for ne in data.get("notas_encomenda", []):
        ne.setdefault("fornecedor", "")
        ne.setdefault("fornecedor_id", "")
        ne.setdefault("contacto", "")
        ne.setdefault("data_entrega", "")
        ne.setdefault("obs", "")
        ne.setdefault("local_descarga", "")
        ne.setdefault("meio_transporte", "")
        ne.setdefault("estado", "Em edição")
        ne.setdefault("linhas", [])
        ne.setdefault("total", 0.0)
        ne.setdefault("entregas", [])
        ne.setdefault("guia_ultima", "")
        ne.setdefault("fatura_ultima", "")
        ne.setdefault("fatura_caminho_ultima", "")
        ne.setdefault("data_doc_ultima", "")
        ne.setdefault("data_ultima_entrega", "")
        ne.setdefault("documentos", [])
        ne.setdefault("oculta", False)
        ne.setdefault("_draft", False)
        ne.setdefault("origem_cotacao", "")
        ne.setdefault("ne_geradas", [])
    normalized_tariffs = []
    for row in data.get("transportes_tarifarios", []):
        if not isinstance(row, dict):
            continue
        normalized_tariffs.append(
            {
                "id": row.get("id", ""),
                "transportadora_id": str(row.get("transportadora_id", "") or "").strip(),
                "transportadora_nome": str(row.get("transportadora_nome", "") or "").strip(),
                "zona": str(row.get("zona", "") or "").strip(),
                "valor_base": _to_num(row.get("valor_base")) or 0.0,
                "valor_por_palete": _to_num(row.get("valor_por_palete")) or 0.0,
                "valor_por_kg": _to_num(row.get("valor_por_kg")) or 0.0,
                "valor_por_m3": _to_num(row.get("valor_por_m3")) or 0.0,
                "custo_minimo": _to_num(row.get("custo_minimo")) or 0.0,
                "ativo": bool(row.get("ativo", True)),
                "observacoes": str(row.get("observacoes", "") or "").strip(),
            }
        )
    data["transportes_tarifarios"] = normalized_tariffs
    normalize_notas_encomenda(data)
    for s in data.get("at_series", []):
        if not isinstance(s, dict):
            continue
        s["doc_type"] = str(s.get("doc_type", "GT") or "GT").strip().upper() or "GT"
        s["serie_id"] = str(s.get("serie_id", "") or "").strip()
        inicio_seq = int(parse_float(s.get("inicio_sequencia", 1), 1) or 1)
        next_seq = int(parse_float(s.get("next_seq", inicio_seq), inicio_seq) or inicio_seq)
        if inicio_seq < 1:
            inicio_seq = 1
        if next_seq < inicio_seq:
            next_seq = inicio_seq
        s["inicio_sequencia"] = inicio_seq
        s["next_seq"] = next_seq
        s.setdefault("data_inicio_prevista", "")
        s.setdefault("validation_code", "")
        s.setdefault("status", "PENDENTE")
        s.setdefault("last_error", "")
        s.setdefault("last_sent_payload_hash", "")
        s.setdefault("updated_at", now_iso())
    for ex in data.get("expedicoes", []):
        if not isinstance(ex, dict):
            continue
        ex.setdefault("numero", "")
        ex.setdefault("tipo", "OFF")
        ex.setdefault("encomenda", "")
        ex.setdefault("cliente", "")
        ex.setdefault("cliente_nome", "")
        ex.setdefault("destinatario", "")
        ex.setdefault("dest_nif", "")
        ex.setdefault("dest_morada", "")
        ex.setdefault("local_carga", "")
        ex.setdefault("local_descarga", "")
        ex.setdefault("matricula", "")
        ex.setdefault("transportador", "")
        ex.setdefault("data_emissao", now_iso())
        ex.setdefault("data_transporte", "")
        ex.setdefault("codigo_at", "")
        ex.setdefault("serie_id", "")
        ex.setdefault("seq_num", 0)
        ex.setdefault("at_validation_code", "")
        ex.setdefault("atcud", "")
        if not ex.get("serie_id"):
            num_match = re.match(r"^GT-(\d{4})-(\d{1,})$", str(ex.get("numero", "") or "").strip())
            if num_match:
                ex["serie_id"] = f"GT{num_match.group(1)}"
        if int(parse_float(ex.get("seq_num", 0), 0) or 0) <= 0:
            num_match = re.match(r"^GT-\d{4}-(\d{1,})$", str(ex.get("numero", "") or "").strip())
            if num_match:
                ex["seq_num"] = int(num_match.group(1))
        emit_cfg = get_guia_emitente_info()
        ex.setdefault("emitente_nome", emit_cfg.get("nome", ""))
        ex.setdefault("emitente_nif", emit_cfg.get("nif", ""))
        ex.setdefault("emitente_morada", emit_cfg.get("morada", ""))
        if not ex.get("at_validation_code") and ex.get("codigo_at"):
            ex["at_validation_code"] = str(ex.get("codigo_at", "") or "")
        if not ex.get("codigo_at") and ex.get("at_validation_code"):
            ex["codigo_at"] = str(ex.get("at_validation_code", "") or "")
        seq_num = int(parse_float(ex.get("seq_num", 0), 0) or 0)
        ex["seq_num"] = seq_num
        if not ex.get("atcud") and ex.get("at_validation_code") and seq_num > 0:
            ex["atcud"] = f"{str(ex.get('at_validation_code', '')).strip()}-{seq_num}"
        ex.setdefault("observacoes", "")
        ex.setdefault("estado", "Emitida")
        ex.setdefault("created_by", "")
        ex.setdefault("anulada", False)
        ex.setdefault("anulada_motivo", "")
        ex.setdefault("linhas", [])
    for ex in data.get("expedicoes", []):
        if not isinstance(ex, dict):
            continue
        sid = str(ex.get("serie_id", "") or "").strip()
        seq_num = int(parse_float(ex.get("seq_num", 0), 0) or 0)
        if not sid or seq_num <= 0:
            continue
        val_code = str(ex.get("at_validation_code", "") or ex.get("codigo_at", "") or "").strip()
        s = ensure_at_series_record(
            data,
            doc_type="GT",
            serie_id=sid,
            issue_date=ex.get("data_emissao") or now_iso(),
            validation_code_hint=val_code,
        )
        next_seq = int(parse_float(s.get("next_seq", 1), 1) or 1)
        if seq_num >= next_seq:
            s["next_seq"] = seq_num + 1
        if val_code and not str(s.get("validation_code", "") or "").strip():
            s["validation_code"] = val_code
            s["status"] = "REGISTADA"
        s["updated_at"] = now_iso()
    for enc in data.get("encomendas", []):
        update_estado_encomenda_por_espessuras(enc)
    if migration_required:
        try:
            save_data(data, force=True)
        except Exception:
            pass
    return data


def _set_last_save_fingerprint(data):
    global _LAST_SAVE_FINGERPRINT
    try:
        payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=str)
        _LAST_SAVE_FINGERPRINT = hashlib.sha1(payload.encode("utf-8")).hexdigest()
    except Exception:
        _LAST_SAVE_FINGERPRINT = None


def _save_data_fingerprint(data):
    try:
        payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=str)
        return hashlib.sha1(payload.encode("utf-8")).hexdigest()
    except Exception:
        return ""


def _snapshot_for_save(data):
    try:
        return copy.deepcopy(data)
    except Exception:
        try:
            return json.loads(json.dumps(data, ensure_ascii=False, default=str))
        except Exception:
            return data


def _async_save_worker_loop():
    global _LAST_SAVE_FINGERPRINT, _LAST_SAVE_TS, _LAST_SAVED_TOKEN
    global _ASYNC_SAVE_PENDING_DATA, _ASYNC_SAVE_PENDING_FP, _ASYNC_SAVE_PENDING_TOKEN
    global _ASYNC_SAVE_IN_PROGRESS, _ASYNC_SAVE_LAST_ERROR
    conn = None
    try:
        while True:
            _ASYNC_SAVE_EVENT.wait(0.5)
            if _ASYNC_SAVE_STOP.is_set() and not _ASYNC_SAVE_EVENT.is_set():
                break
            while True:
                with _ASYNC_SAVE_LOCK:
                    payload = _ASYNC_SAVE_PENDING_DATA
                    fp = _ASYNC_SAVE_PENDING_FP
                    token = _ASYNC_SAVE_PENDING_TOKEN
                    _ASYNC_SAVE_PENDING_DATA = None
                    _ASYNC_SAVE_PENDING_FP = ""
                    _ASYNC_SAVE_PENDING_TOKEN = 0
                    if payload is None:
                        _ASYNC_SAVE_EVENT.clear()
                        break
                    _ASYNC_SAVE_IN_PROGRESS = True
                try:
                    if conn is None:
                        conn = _mysql_connect()
                    snap = _snapshot_for_save(payload)
                    _mysql_save_relational_data(snap if snap is not None else payload, conn=conn)
                    with _ASYNC_SAVE_LOCK:
                        _LAST_SAVE_TS = time.monotonic()
                        if fp:
                            _LAST_SAVE_FINGERPRINT = fp
                        if token:
                            _LAST_SAVED_TOKEN = max(_LAST_SAVED_TOKEN, int(token))
                        _ASYNC_SAVE_LAST_ERROR = ""
                except Exception as ex:
                    with _ASYNC_SAVE_LOCK:
                        _ASYNC_SAVE_LAST_ERROR = str(ex)
                    if conn is not None:
                        try:
                            conn.close()
                        except Exception:
                            pass
                        conn = None
                finally:
                    with _ASYNC_SAVE_LOCK:
                        _ASYNC_SAVE_IN_PROGRESS = False
            if _ASYNC_SAVE_STOP.is_set():
                with _ASYNC_SAVE_LOCK:
                    if _ASYNC_SAVE_PENDING_DATA is None and not _ASYNC_SAVE_IN_PROGRESS:
                        break
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def _start_async_save_worker():
    global _ASYNC_SAVE_THREAD
    if not _ASYNC_SAVE_ENABLED:
        return
    if _ASYNC_SAVE_THREAD is not None and _ASYNC_SAVE_THREAD.is_alive():
        return
    _ASYNC_SAVE_STOP.clear()
    _ASYNC_SAVE_THREAD = threading.Thread(target=_async_save_worker_loop, name="lugest-save-worker", daemon=True)
    _ASYNC_SAVE_THREAD.start()


def _queue_async_save(data, fp="", token=0):
    global _ASYNC_SAVE_PENDING_DATA, _ASYNC_SAVE_PENDING_FP, _ASYNC_SAVE_PENDING_TOKEN
    with _ASYNC_SAVE_LOCK:
        _ASYNC_SAVE_PENDING_DATA = data
        _ASYNC_SAVE_PENDING_FP = str(fp or "")
        _ASYNC_SAVE_PENDING_TOKEN = int(token or 0)
    _ASYNC_SAVE_EVENT.set()
    _start_async_save_worker()


def _drain_async_saves(timeout_sec=12.0):
    if not _ASYNC_SAVE_ENABLED:
        return True
    deadline = time.time() + max(0.1, float(timeout_sec or 0))
    while time.time() < deadline:
        with _ASYNC_SAVE_LOCK:
            done = (_ASYNC_SAVE_PENDING_DATA is None) and (not _ASYNC_SAVE_IN_PROGRESS)
        if done:
            return True
        time.sleep(0.05)
    return False


def _consume_async_save_error():
    global _ASYNC_SAVE_LAST_ERROR
    with _ASYNC_SAVE_LOCK:
        err = _ASYNC_SAVE_LAST_ERROR
        _ASYNC_SAVE_LAST_ERROR = ""
    return err


def _save_data_now(data, fp="", token=0, blocking=False):
    global _LAST_SAVE_FINGERPRINT, _LAST_SAVE_TS, _LAST_SAVED_TOKEN
    if _ASYNC_SAVE_ENABLED and not blocking:
        _queue_async_save(data, fp=fp, token=token)
        return
    _mysql_save_relational_data(data)
    _LAST_SAVE_TS = time.monotonic()
    if fp:
        _LAST_SAVE_FINGERPRINT = fp
    if token:
        _LAST_SAVED_TOKEN = max(_LAST_SAVED_TOKEN, int(token))


def flush_pending_save(force=False):
    pending = _PENDING_SAVE_DATA
    if pending is None:
        return False
    if not force and _SAVE_MIN_INTERVAL_SEC > 0 and (time.monotonic() - _LAST_SAVE_TS) < _SAVE_MIN_INTERVAL_SEC:
        return False
    save_data(pending, force=True)
    return True


def save_data(data, force=False):
    global _LAST_SAVE_FINGERPRINT, _PENDING_SAVE_DATA, _SAVE_CHANGE_TOKEN, _LAST_SAVED_TOKEN
    if not USE_MYSQL_STORAGE:
        raise RuntimeError("A aplicação está configurada para usar apenas MySQL.")
    normalize_notas_encomenda(data)
    token = 0
    fp = _save_data_fingerprint(data)
    if not force:
        _PENDING_SAVE_DATA = data
        _SAVE_CHANGE_TOKEN += 1
        token = _SAVE_CHANGE_TOKEN
        if fp and fp == _LAST_SAVE_FINGERPRINT:
            _PENDING_SAVE_DATA = None
            _LAST_SAVED_TOKEN = max(_LAST_SAVED_TOKEN, token)
            return
        if _ASYNC_SAVE_ENABLED:
            if _SAVE_MIN_INTERVAL_SEC > 0 and (time.monotonic() - _LAST_SAVE_TS) < _SAVE_MIN_INTERVAL_SEC:
                return
            data = _PENDING_SAVE_DATA
            _PENDING_SAVE_DATA = None
            try:
                _save_data_now(data, fp=fp, token=token, blocking=False)
            except Exception as ex:
                try:
                    messagebox.showerror("Erro MySQL", f"Falha ao guardar na base de dados:\n{ex}")
                except Exception:
                    pass
                raise
            return
        force = True
    if _PENDING_SAVE_DATA is not None:
        data = _PENDING_SAVE_DATA
        _PENDING_SAVE_DATA = None
        token = _SAVE_CHANGE_TOKEN
        fp = _save_data_fingerprint(data)
        if token and _LAST_SAVED_TOKEN == token:
            return
    if fp and fp == _LAST_SAVE_FINGERPRINT:
        return
    try:
        _save_data_now(data, fp=fp, token=token, blocking=True)
    except Exception as ex:
        try:
            messagebox.showerror("Erro MySQL", f"Falha ao guardar na base de dados:\n{ex}")
        except Exception:
            pass
        raise


def upsert_local_user_account(username, password, role="Admin", *, reset_existing=False):
    username_txt = _clip(username, 50)
    role_txt = _normalize_role_name(_clip(role, 50)) or "Admin"
    if not username_txt:
        raise ValueError("Indica um username para o utilizador local.")
    validate_local_password(username_txt, password)
    stored_password = _clip(normalize_password_for_storage(username_txt, password, require_strong=True), 255)
    data = load_data()
    users = list(data.get("users", []) or [])
    target = None
    created = True
    for row in users:
        if str(row.get("username", "") or "").strip().lower() == username_txt.lower():
            target = row
            created = False
            break
    if target is not None and not reset_existing:
        raise ValueError("Ja existe um utilizador com esse username. Usa --reset-admin para repor a password.")
    if target is None:
        target = {"username": username_txt, "password": stored_password, "role": role_txt}
        users.append(target)
    else:
        target.update({"username": username_txt, "password": stored_password, "role": role_txt})
    data["users"] = users
    normalize_notas_encomenda(data)
    _save_data_now(data, fp=_save_data_fingerprint(data), token=0, blocking=True)
    return {
        "username": username_txt,
        "role": role_txt,
        "created": bool(created),
        "reset": bool(not created),
    }


def handle_admin_setup_cli(argv=None):
    args = list(argv if argv is not None else sys.argv)
    if "--setup-admin" not in args:
        return None
    parser = argparse.ArgumentParser(prog=(os.path.basename(args[0]) if args else "main.exe"))
    parser.add_argument("--setup-admin", action="store_true")
    parser.add_argument("--admin-username", default="")
    parser.add_argument("--admin-password", default="")
    parser.add_argument("--admin-role", default="Admin")
    parser.add_argument("--reset-admin", action="store_true")
    cli = parser.parse_args(args[1:])
    username = str(cli.admin_username or "").strip()
    password = str(cli.admin_password or "")
    if not username:
        raise SystemExit("Falta --admin-username.")
    if not password:
        raise SystemExit("Falta --admin-password.")
    _assert_mysql_runtime_ready()
    result = upsert_local_user_account(
        username,
        password,
        role=str(cli.admin_role or "Admin"),
        reset_existing=bool(cli.reset_admin),
    )
    action_txt = "atualizada" if bool(result.get("reset", False)) else "criada"
    print(f"admin-setup-ok: conta local {action_txt} para '{result['username']}' ({result['role']})")
    return 0


def _mysql_next_counter(counter_key, initial_next=1):
    if not USE_MYSQL_STORAGE or not MYSQL_AVAILABLE:
        return None
    key = str(counter_key or "").strip()
    if not key:
        return None
    start = max(1, int(parse_float(initial_next, 1) or 1))
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS app_counters (
                    ckey VARCHAR(80) PRIMARY KEY,
                    next_value BIGINT NOT NULL,
                    updated_at DATETIME NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                """
            )
            conn.commit()
            cur.execute(
                """
                INSERT INTO app_counters (ckey, next_value, updated_at)
                VALUES (%s, %s, NOW())
                ON DUPLICATE KEY UPDATE next_value=next_value
                """,
                (key, start),
            )
            cur.execute("SELECT next_value FROM app_counters WHERE ckey=%s FOR UPDATE", (key,))
            row = cur.fetchone() or {}
            n = max(start, int(parse_float(row.get("next_value", start), start) or start))
            cur.execute("UPDATE app_counters SET next_value=%s, updated_at=NOW() WHERE ckey=%s", (n + 1, key))
        conn.commit()
        return n
    except Exception:
        try:
            if conn is not None:
                conn.rollback()
        except Exception:
            pass
        return None
    finally:
        try:
            if conn is not None:
                conn.close()
        except Exception:
            pass


def next_encomenda_numero(data):
    seq = data.setdefault("seq", {})
    max_n = 0
    for e in data.get("encomendas", []):
        num = e.get("numero", "")
        if num.startswith("BARCELBAL") and num[9:].isdigit():
            max_n = max(max_n, int(num[9:]))
    local_next = max(max_n + 1, int(seq.get("encomenda", 1) or 1))
    n = _mysql_next_counter("encomenda", local_next) or local_next
    seq["encomenda"] = n + 1
    return f"BARCELBAL{n:04d}"


def next_orc_numero(data):
    year = datetime.now().year
    max_n = _next_seq_from_pattern(
        [o.get("numero") for o in data.get("orcamentos", []) if isinstance(o, dict)],
        rf"^ORC-{year}-(\d{{4,}})$",
        1,
    )
    local_next = max(max_n, int(data.get("orc_seq", 1) or 1))
    n = _mysql_next_counter(f"orc:{year}", local_next) or local_next
    data["orc_seq"] = n + 1
    return f"ORC-{year}-{n:04d}"


def next_of_numero(data):
    year = datetime.now().year
    values = [p.get("of") for e in data.get("encomendas", []) for p in encomenda_pecas(e)]
    for o in data.get("orcamentos", []):
        if isinstance(o, dict):
            values.extend((line or {}).get("of") for line in list(o.get("linhas", []) or []) if isinstance(line, dict))
    max_n = _next_seq_from_pattern(values, rf"^OF-{year}-(\d{{4,}})$", 1)
    local_next = max(max_n, int(data.get("of_seq", 1) or 1))
    n = _mysql_next_counter(f"of:{year}", local_next) or local_next
    data["of_seq"] = n + 1
    return f"OF-{year}-{n:04d}"


def next_produto_numero(data):
    seq = data.setdefault("seq", {})
    # gera com base no maior ja existente para evitar saltos mesmo com dados antigos
    max_n = 0
    for p in data.get("produtos", []):
        cod = p.get("codigo", "")
        if cod.startswith("PRD-") and cod[4:].isdigit():
            max_n = max(max_n, int(cod[4:]))
    seq_n = int(seq.get("produto", 1))
    local_next = max(max_n + 1, seq_n)
    n = _mysql_next_counter("produto", local_next) or local_next
    seq["produto"] = n + 1
    return f"PRD-{n:04d}"


def peek_next_produto_numero(data):
    max_n = 0
    for p in data.get("produtos", []):
        cod = p.get("codigo", "")
        if cod.startswith("PRD-") and cod[4:].isdigit():
            max_n = max(max_n, int(cod[4:]))
    return f"PRD-{(max_n + 1):04d}"


def ensure_produto_seq(data, codigo):
    cod = str(codigo or "").strip()
    if not (cod.startswith("PRD-") and cod[4:].isdigit()):
        return
    n = int(cod[4:])
    seq = data.setdefault("seq", {})
    atual = int(seq.get("produto", 1))
    seq["produto"] = max(atual, n + 1)


def next_ne_numero(data):
    seq = data.setdefault("seq", {})
    year = datetime.now().year
    max_n = 0
    for ne in data.get("notas_encomenda", []):
        num = str(ne.get("numero", ""))
        prefix = f"NE-{year}-"
        if num.startswith(prefix):
            tail = num[len(prefix):]
            if tail.isdigit():
                max_n = max(max_n, int(tail))
    local_next = max(max_n + 1, int(seq.get("ne", 1) or 1))
    n = _mysql_next_counter(f"ne:{year}", local_next) or local_next
    seq["ne"] = n + 1
    return f"NE-{year}-{n:04d}"


def peek_next_ne_numero(data):
    year = datetime.now().year
    max_n = 0
    for ne in data.get("notas_encomenda", []):
        num = str(ne.get("numero", ""))
        prefix = f"NE-{year}-"
        if num.startswith(prefix):
            tail = num[len(prefix):]
            if tail.isdigit():
                max_n = max(max_n, int(tail))
    n = max_n + 1
    return f"NE-{year}-{n:04d}"


def next_expedicao_numero(data):
    year = datetime.now().year
    max_n = 0
    for ex in data.get("expedicoes", []):
        if not isinstance(ex, dict):
            continue
        num = str(ex.get("numero", ""))
        prefix = f"GT-{year}-"
        if num.startswith(prefix):
            tail = num[len(prefix):]
            if tail.isdigit():
                max_n = max(max_n, int(tail))
    local_next = max(max_n + 1, int(data.get("exp_seq", 1) or 1))
    n = _mysql_next_counter(f"expedicao:{year}", local_next) or local_next
    data["exp_seq"] = n + 1
    return f"GT-{year}-{n:04d}"


def _doc_year_from_value(value):
    txt = str(value or "").replace("T", " ").strip()
    m = re.match(r"^(\d{4})", txt)
    if m:
        try:
            y = int(m.group(1))
            if 1900 <= y <= 9999:
                return y
        except Exception:
            pass
    return datetime.now().year


def _exp_default_serie_id(doc_type="GT", issue_date=None):
    dtp = str(doc_type or "GT").strip().upper() or "GT"
    year = _doc_year_from_value(issue_date)
    return f"{dtp}{year}"


def _find_at_series(data, *, doc_type, serie_id):
    dtp = str(doc_type or "").strip().upper()
    sid = str(serie_id or "").strip()
    if not dtp or not sid:
        return None
    for s in data.get("at_series", []):
        if not isinstance(s, dict):
            continue
        if str(s.get("doc_type", "") or "").strip().upper() == dtp and str(s.get("serie_id", "") or "").strip() == sid:
            return s
    return None


def ensure_at_series_record(
    data,
    *,
    doc_type="GT",
    serie_id=None,
    issue_date=None,
    validation_code_hint="",
):
    dtp = str(doc_type or "GT").strip().upper() or "GT"
    sid = str(serie_id or "").strip() or _exp_default_serie_id(dtp, issue_date or now_iso())
    data.setdefault("at_series", [])
    s = _find_at_series(data, doc_type=dtp, serie_id=sid)
    if s is None:
        start_seq = 1
        s = {
            "doc_type": dtp,
            "serie_id": sid,
            "inicio_sequencia": start_seq,
            "next_seq": start_seq,
            "data_inicio_prevista": str(issue_date or now_iso())[:10],
            "validation_code": "",
            "status": "PENDENTE",
            "last_error": "",
            "last_sent_payload_hash": "",
            "updated_at": now_iso(),
        }
        data["at_series"].append(s)
    s["doc_type"] = dtp
    s["serie_id"] = sid
    inicio_seq = int(parse_float(s.get("inicio_sequencia", 1), 1) or 1)
    next_seq = int(parse_float(s.get("next_seq", inicio_seq), inicio_seq) or inicio_seq)
    if inicio_seq < 1:
        inicio_seq = 1
    if next_seq < inicio_seq:
        next_seq = inicio_seq
    s["inicio_sequencia"] = inicio_seq
    s["next_seq"] = next_seq
    if not s.get("data_inicio_prevista"):
        s["data_inicio_prevista"] = str(issue_date or now_iso())[:10]
    hint = str(validation_code_hint or "").strip()
    if hint and not str(s.get("validation_code", "") or "").strip():
        s["validation_code"] = hint
        s["status"] = "REGISTADA"
    s.setdefault("status", "PENDENTE")
    s.setdefault("last_error", "")
    s.setdefault("last_sent_payload_hash", "")
    s["updated_at"] = now_iso()
    return s


def next_expedicao_identifiers(
    data,
    *,
    issue_date=None,
    doc_type="GT",
    serie_id=None,
    validation_code_hint="",
):
    s = ensure_at_series_record(
        data,
        doc_type=doc_type,
        serie_id=serie_id,
        issue_date=issue_date or now_iso(),
        validation_code_hint=validation_code_hint,
    )
    validation_code = str(s.get("validation_code", "") or "").strip()
    if not validation_code:
        sid = str(s.get("serie_id", "") or "").strip() or _exp_default_serie_id(doc_type, issue_date or now_iso())
        return None, f"A serie {sid} nao tem codigo de validacao AT. Registe a serie na AT ou indique o codigo para inicializar."

    sid = str(s.get("serie_id", "") or "").strip()
    year = _doc_year_from_value(issue_date or now_iso())
    seq = int(parse_float(s.get("next_seq", 1), 1) or 1)
    start_seq = int(parse_float(s.get("inicio_sequencia", 1), 1) or 1)
    if seq < start_seq:
        seq = start_seq
    if seq < 1:
        seq = 1

    used_doc_nums = set()
    used_seq = set()
    for ex in data.get("expedicoes", []):
        if not isinstance(ex, dict):
            continue
        n = str(ex.get("numero", "") or "").strip()
        if n:
            used_doc_nums.add(n)
        ex_sid = str(ex.get("serie_id", "") or "").strip()
        ex_seq = int(parse_float(ex.get("seq_num", 0), 0) or 0)
        if ex_sid and ex_seq > 0:
            used_seq.add((ex_sid, ex_seq))

    while True:
        numero = f"GT-{year}-{seq:04d}"
        if (sid, seq) not in used_seq and numero not in used_doc_nums:
            break
        seq += 1

    atcud = f"{validation_code}-{seq}"
    s["next_seq"] = seq + 1
    s["status"] = "REGISTADA"
    s["updated_at"] = now_iso()
    return {
        "numero": numero,
        "serie_id": sid,
        "seq_num": seq,
        "validation_code": validation_code,
        "atcud": atcud,
    }, ""


def next_fornecedor_numero(data):
    local_next = _load_fornecedor_sequence_next(data)
    n = _mysql_next_counter("fornecedor", local_next) or local_next
    _store_fornecedor_sequence_next(data, n + 1)
    return f"FOR-{n:04d}"


def peek_next_fornecedor_numero(data):
    n = _load_fornecedor_sequence_next(data)
    return f"FOR-{n:04d}"


def reserve_fornecedor_numero(data, supplier_id):
    current = _load_fornecedor_sequence_next(data)
    supplier_n = _extract_fornecedor_seq(supplier_id)
    if supplier_n <= 0:
        return
    _store_fornecedor_sequence_next(data, max(current, supplier_n + 1))


def next_transporte_numero(data):
    year = datetime.now().year
    local_next = _load_transport_sequence_next(data)
    n = _mysql_next_counter(f"transporte:{year}", local_next) or local_next
    _store_transport_sequence_next(data, n + 1)
    return f"TR-{year}-{n:04d}"


def peek_next_transporte_numero(data):
    n = _load_transport_sequence_next(data)
    return f"TR-{datetime.now().year}-{n:04d}"


def reserve_transporte_numero(data, transport_number):
    current = _load_transport_sequence_next(data)
    transport_n = _extract_transport_seq(transport_number)
    if transport_n <= 0:
        return
    _store_transport_sequence_next(data, max(current, transport_n + 1))


def next_opp_numero(data):
    year = datetime.now().year
    values = []
    for e in data.get("encomendas", []):
        for p in encomenda_pecas(e):
            values.append(p.get("opp"))
    for o in data.get("orcamentos", []):
        if isinstance(o, dict):
            values.extend((line or {}).get("opp") for line in list(o.get("linhas", []) or []) if isinstance(line, dict))
    local_next = max(_next_seq_from_pattern(values, rf"^OPP-{year}-(\d{{4,}})$", 1), int(data.get("opp_seq", 1) or 1))
    n = _mysql_next_counter(f"opp:{year}", local_next) or local_next
    data["opp_seq"] = n + 1
    return f"OPP-{year}-{n:04d}"


def next_cliente_codigo(data):
    usados = set()
    for c in data.get("clientes", []):
        codigo = str(c.get("codigo", "") or "").strip().upper()
        if codigo.startswith("CL") and codigo[2:].isdigit():
            usados.add(int(codigo[2:]))
    n = 1
    while n in usados:
        n += 1
    data.setdefault("seq", {})
    data["seq"]["cliente"] = n + 1
    return f"CL{n:04d}"


def next_ref_interna(data, cliente_codigo):
    seqs = data["seq"].setdefault("ref_interna", {})
    current = int(seqs.get(cliente_codigo, 0))
    if current <= 0:
        current = _highest_ref_interna_seq(data, cliente_codigo)
    current += 1
    seqs[cliente_codigo] = current
    return f"{cliente_codigo}-{current:04d}REV00"


def find_cliente(data, codigo):
    for c in data["clientes"]:
        if c["codigo"] == codigo:
            return c
    return None


def next_ref_interna_global(data, cliente_codigo):
    current = _highest_ref_interna_seq(data, cliente_codigo) + 1
    data.setdefault("seq", {}).setdefault("ref_interna", {})[cliente_codigo] = current
    return f"{cliente_codigo}-{current:04d}REV00"


def next_ref_interna_unique(data, cliente_codigo, existing_refs=None):
    existing_refs = set(existing_refs or [])
    current = _highest_ref_interna_seq(data, cliente_codigo)
    while True:
        current += 1
        cand = f"{cliente_codigo}-{current:04d}REV00"
        if cand not in existing_refs:
            data.setdefault("seq", {}).setdefault("ref_interna", {})[cliente_codigo] = current
            return cand


def _extract_ref_interna_seq(ref, cliente_codigo=""):
    raw = str(ref or "").strip().upper()
    cli = str(cliente_codigo or "").strip().upper()
    if not raw:
        return 0
    if cli and not raw.startswith(f"{cli}-"):
        return 0
    import re
    match = re.search(r"-(\d{4,5})(?:REV\d{2})?$", raw)
    if not match:
        return 0
    try:
        return int(match.group(1))
    except Exception:
        return 0


def _highest_ref_interna_seq(data, cliente_codigo):
    current = 0
    cli = str(cliente_codigo or "").strip()
    for e in data.get("encomendas", []):
        if str(e.get("cliente", "") or "").strip() != cli:
            continue
        for p in encomenda_pecas(e):
            current = max(current, _extract_ref_interna_seq(p.get("ref_interna", ""), cli))
    for o in data.get("orcamentos", []):
        for l in o.get("linhas", []):
            current = max(current, _extract_ref_interna_seq(l.get("ref_interna", ""), cli))
    for _ref_ext, payload in (data.get("orc_refs", {}) or {}).items():
        current = max(current, _extract_ref_interna_seq((payload or {}).get("ref_interna", ""), cli))
    return current


def list_unique(data, key):
    values = set()
    for m in data["materiais"]:
        v = m.get(key)
        if v is not None:
            values.add(str(v))
    return sorted(values)


def _resolve_branding_path(p):
    txt = str(p or "").strip()
    if not txt:
        return ""
    if os.path.isabs(txt):
        return txt
    candidates = [os.path.join(BASE_DIR, txt)]
    base_parent = os.path.dirname(os.path.abspath(BASE_DIR))
    if base_parent and os.path.normcase(os.path.abspath(base_parent)) != os.path.normcase(os.path.abspath(BASE_DIR)):
        candidates.append(os.path.join(base_parent, txt))
    candidates.append(os.path.join(os.getcwd(), txt))
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return txt


def _mysql_read_branding_payload():
    use_mysql = bool(globals().get("USE_MYSQL", False))
    mysql_available = bool(globals().get("MYSQL_AVAILABLE", False))
    if not use_mysql or not mysql_available:
        return {}
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS app_config (
                    ckey VARCHAR(80) PRIMARY KEY,
                    cvalue LONGTEXT NULL,
                    updated_at DATETIME NULL
                )
                """
            )
            cur.execute("SELECT cvalue FROM app_config WHERE ckey=%s LIMIT 1", ("branding_config",))
            row = cur.fetchone()
        if not row:
            return {}
        raw = row.get("cvalue") if isinstance(row, dict) else row[0]
        if raw is None:
            return {}
        if isinstance(raw, (bytes, bytearray)):
            raw = raw.decode("utf-8", errors="ignore")
        parsed = json.loads(str(raw))
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass


def get_branding_config():
    global _BRANDING_CACHE
    if _BRANDING_CACHE is not None:
        return _BRANDING_CACHE

    cfg = _mysql_read_branding_payload()
    cfg_candidates = [
        os.path.join(BASE_DIR, BRANDING_FILE),
        os.path.join(os.path.dirname(os.path.abspath(BASE_DIR)), BRANDING_FILE),
        BRANDING_FILE,
    ]
    if not isinstance(cfg, dict) or not cfg:
        cfg = {}
        for cp in cfg_candidates:
            try:
                if os.path.exists(cp):
                    with open(cp, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    if isinstance(data, dict):
                        cfg = data
                        break
            except Exception:
                pass
    if isinstance(cfg, dict):
        cfg = _repair_mojibake_structure(cfg)
    else:
        cfg = {}

    logos = []
    seen = set()

    def add_logo(v):
        p = _resolve_branding_path(v)
        if p and p not in seen:
            seen.add(p)
            logos.append(p)

    add_logo(cfg.get("logo", ""))
    for v in cfg.get("logo_candidates", []) if isinstance(cfg.get("logo_candidates", []), list) else []:
        add_logo(v)
    for v in ORC_LOGO_CANDIDATES:
        add_logo(v)

    rodape = cfg.get("empresa_info_rodape", ORC_EMPRESA_INFO_RODAPE)
    if isinstance(rodape, str):
        rodape = [rodape]
    if not isinstance(rodape, list):
        rodape = list(ORC_EMPRESA_INFO_RODAPE)
    rodape = [str(x) for x in rodape if str(x).strip()]
    if not rodape:
        rodape = list(ORC_EMPRESA_INFO_RODAPE)

    emit_cfg = cfg.get("guia_emitente", {})
    if not isinstance(emit_cfg, dict):
        emit_cfg = {}
    nif_guess = ""
    for ln in rodape:
        m = re.search(r"\bNIF[:\s]*([0-9]{9})\b", str(ln), flags=re.IGNORECASE)
        if m:
            nif_guess = m.group(1)
            break
    emitente = {
        "nome": str(emit_cfg.get("nome", "") or (rodape[0] if rodape else "")).strip(),
        "nif": str(emit_cfg.get("nif", "") or nif_guess).strip(),
        "morada": str(emit_cfg.get("morada", "") or (rodape[1] if len(rodape) > 1 else "")).strip(),
        "local_carga": str(emit_cfg.get("local_carga", "") or (rodape[1] if len(rodape) > 1 else "")).strip(),
    }
    guia_extra = cfg.get("guia_info_extra", [])
    if isinstance(guia_extra, str):
        guia_extra = [guia_extra]
    if not isinstance(guia_extra, list):
        guia_extra = []
    guia_extra = [str(x).strip() for x in guia_extra if str(x).strip()]

    primary_color = _normalize_hex_color(cfg.get("primary_color"), CTK_PRIMARY_RED) or CTK_PRIMARY_RED
    try:
        logo_scale_pct = max(50, min(250, int(float(cfg.get("logo_scale_pct", 100) or 100))))
    except Exception:
        logo_scale_pct = 100

    _BRANDING_CACHE = {
        "logo_candidates": logos,
        "empresa_info_rodape": rodape,
        "guia_emitente": emitente,
        "guia_info_extra": guia_extra,
        "primary_color": primary_color,
        "logo_scale_pct": logo_scale_pct,
    }
    return _BRANDING_CACHE


def get_orc_logo_path():
    for p in get_branding_config().get("logo_candidates", []):
        if p and os.path.exists(p):
            return p
    return ""


def draw_pdf_logo_box(c, page_h, x, y_top, box_size=34, padding=3, draw_border=True, box_h=None):
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader

    box_w = float(box_size)
    box_h = float(box_h if box_h is not None else box_size)
    y = page_h - y_top - box_h
    c.saveState()
    try:
        if draw_border:
            c.setLineWidth(0.8)
            c.setStrokeColor(colors.HexColor("#9fb0c3"))
            c.rect(x, y, box_w, box_h, stroke=1, fill=0)

        logo = get_orc_logo_path()
        if not logo or not os.path.exists(logo):
            return

        img_reader = None
        iw = 0
        ih = 0

        # Remove transparent/white margins to keep the logo visually larger in the fixed box.
        try:
            from PIL import Image, ImageChops
            img = Image.open(logo).convert("RGBA")
            alpha = img.split()[-1]
            bbox = alpha.getbbox()
            if not bbox:
                rgb = img.convert("RGB")
                bg = Image.new("RGB", rgb.size, (255, 255, 255))
                diff = ImageChops.difference(rgb, bg)
                bbox = diff.getbbox()
            if bbox:
                img = img.crop(bbox)
            iw, ih = img.size
            img_reader = ImageReader(img)
        except Exception:
            try:
                img_reader = ImageReader(logo)
                iw, ih = img_reader.getSize()
            except Exception:
                img_reader = None
                iw = ih = 0

        if not img_reader or iw <= 0 or ih <= 0:
            return

        pad = max(0.0, float(padding)) / 3.2
        max_w = max(1.0, box_w - (2.0 * pad))
        max_h = max(1.0, box_h - (2.0 * pad))
        scale = min(max_w / float(iw), max_h / float(ih))
        try:
            logo_scale = max(0.5, min(2.5, float(get_branding_config().get("logo_scale_pct", 100) or 100) / 100.0))
        except Exception:
            logo_scale = 1.0
        scale *= logo_scale
        draw_w = max(1.0, float(iw) * scale)
        draw_h = max(1.0, float(ih) * scale)
        draw_x = x + (box_w - draw_w) / 2.0
        draw_y = y + (box_h - draw_h) / 2.0
        clip_path = c.beginPath()
        clip_path.rect(x, y, box_w, box_h)
        c.clipPath(clip_path, stroke=0, fill=0)
        c.drawImage(img_reader, draw_x, draw_y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
    finally:
        c.restoreState()


def draw_pdf_logo_plate(
    c,
    page_h,
    x,
    y_top,
    box_w=118,
    box_h=54,
    *,
    padding=4,
    radius=12,
    fill_color="#FFFFFF",
    stroke_color="#D7DEE8",
    line_width=0.9,
):
    from reportlab.lib import colors

    box_w = float(box_w)
    box_h = float(box_h)
    y = page_h - y_top - box_h
    c.saveState()
    try:
        c.setFillColor(colors.HexColor(str(fill_color or "#FFFFFF")))
        c.setStrokeColor(colors.HexColor(str(stroke_color or "#D7DEE8")))
        c.setLineWidth(float(line_width or 0.9))
        c.roundRect(x, y, box_w, box_h, float(radius or 12), stroke=1, fill=1)
    finally:
        c.restoreState()
    draw_pdf_logo_box(c, page_h, x, y_top, box_size=box_w, box_h=box_h, padding=padding, draw_border=False)


def draw_pdf_header_panel(
    c,
    page_h,
    x,
    y_top,
    width,
    height,
    *,
    radius=14,
    fill_color="#FFFFFF",
    stroke_color="#D7DEE8",
    line_width=1.0,
    accent_color="#E8EEF5",
    accent_height=4,
):
    from reportlab.lib import colors

    width = float(width)
    height = float(height)
    y = page_h - y_top - height
    c.saveState()
    try:
        c.setFillColor(colors.HexColor(str(fill_color or "#FFFFFF")))
        c.setStrokeColor(colors.HexColor(str(stroke_color or "#D7DEE8")))
        c.setLineWidth(float(line_width or 1.0))
        c.roundRect(x, y, width, height, float(radius or 14), stroke=1, fill=1)
        if accent_height and accent_color:
            accent_h = max(1.0, min(float(accent_height), height * 0.16))
            c.setFillColor(colors.HexColor(str(accent_color)))
            c.roundRect(x + 1.2, y + height - accent_h - 1.2, max(8.0, width - 2.4), accent_h, max(2.0, float(radius or 14) * 0.45), stroke=0, fill=1)
    finally:
        c.restoreState()


def apply_window_icon(win):
    if not win:
        return
    ico_candidates = [
        os.path.join(BASE_DIR, "app.ico"),
        os.path.join(BASE_DIR, "logo.ico"),
        "app.ico",
        "logo.ico",
    ]
    for ico in ico_candidates:
        try:
            if ico and os.path.exists(ico):
                win.iconbitmap(ico)
                return
        except Exception:
            pass

    logo = get_orc_logo_path()
    if not logo or not os.path.exists(logo):
        return
    try:
        img = PhotoImage(file=logo)
        win.iconphoto(True, img)
        win._lugest_icon_ref = img
        return
    except Exception:
        pass
    try:
        from PIL import Image, ImageTk
        img = Image.open(logo)
        if max(img.size) > 64:
            scale = 64.0 / float(max(img.size))
            img = img.resize((max(1, int(img.width * scale)), max(1, int(img.height * scale))))
        tk_img = ImageTk.PhotoImage(img)
        win.iconphoto(True, tk_img)
        win._lugest_icon_ref = tk_img
    except Exception:
        pass


def get_empresa_rodape_lines():
    return list(get_branding_config().get("empresa_info_rodape", ORC_EMPRESA_INFO_RODAPE))


def get_guia_emitente_info():
    cfg = get_branding_config()
    base = cfg.get("guia_emitente", {})
    if not isinstance(base, dict):
        base = {}
    return {
        "nome": str(base.get("nome", "") or "").strip(),
        "nif": str(base.get("nif", "") or "").strip(),
        "morada": str(base.get("morada", "") or "").strip(),
        "local_carga": str(base.get("local_carga", "") or "").strip(),
    }


def get_guia_extra_info_lines():
    cfg = get_branding_config()
    lines = cfg.get("guia_info_extra", [])
    if isinstance(lines, str):
        lines = [lines]
    if not isinstance(lines, list):
        return []
    return [str(x).strip() for x in lines if str(x).strip()]


def push_unique(lst, value):
    v = str(value).strip()
    if not v:
        return
    if v not in lst:
        lst.append(v)


def log_stock(data, action, details, operador=""):
    details = format_stock_detail(data, details)
    data.setdefault("stock_log", []).append({
        "data": now_iso(),
        "acao": action,
        "operador": str(operador or "").strip(),
        "detalhes": details,
    })


def format_stock_detail(data, details):
    ids = re.findall(r"(MAT\\d{5})", details)
    if not ids:
        return details
    out = details
    for mat_id in ids:
        for m in data.get("materiais", []):
            if m.get("id") == mat_id:
                dim = f"{m.get('comprimento','')}x{m.get('largura','')}"
                lote = m.get("lote_fornecedor", "")
                rep = f"{m.get('material','')} {m.get('espessura','')} {dim} Lote:{lote}".strip()
                out = out.replace(mat_id, rep)
                break
    return out


def update_refs(data, ref_interna, ref_externa):
    if ref_interna and ref_interna not in data["refs"]:
        data["refs"].append(ref_interna)
    if ref_externa and ref_externa not in data["refs"]:
        data["refs"].append(ref_externa)


def calcular_reservas_encomenda(encomenda):
    reservas = {}
    for p in encomenda_pecas(encomenda):
        key = (p["material"], p["espessura"])
        reservas[key] = reservas.get(key, 0) + p["quantidade_pedida"]
    out = []
    for (material, espessura), qtd in reservas.items():
        out.append({"material": material, "espessura": espessura, "quantidade": qtd})
    return out


def aplicar_reserva_em_stock(data, reservas, fator):
    for r in reservas:
        if r.get("material_id"):
            for m in data["materiais"]:
                if m["id"] == r["material_id"]:
                    m["reservado"] = max(0, m.get("reservado", 0) + fator * r["quantidade"])
                    m["atualizado_em"] = now_iso()
                    break
            continue
        for m in data["materiais"]:
            if m["material"] == r["material"] and m["espessura"] == r["espessura"]:
                m["reservado"] = max(0, m.get("reservado", 0) + fator * r["quantidade"])
                m["atualizado_em"] = now_iso()
                break


def recalc_reservas(data, enc):
    if enc.get("reservas"):
        aplicar_reserva_em_stock(data, enc["reservas"], -1)
    enc["reservas"] = calcular_reservas_encomenda(enc)
    aplicar_reserva_em_stock(data, enc["reservas"], 1)


def atualizar_estados_encomenda(encomenda):
    update_estado_encomenda_por_espessuras(encomenda)


def atualizar_estado_peca(peca):
    fluxo = ensure_peca_operacoes(peca)
    estado_atual = norm_text(peca.get("estado", ""))
    avaria_ativa = bool(peca.get("avaria_ativa"))
    pausa_registada = bool(str(peca.get("interrupcao_peca_motivo", "") or "").strip()) or bool(str(peca.get("interrupcao_peca_ts", "") or "").strip())
    total = (
        parse_float(peca.get("produzido_ok", 0), 0.0)
        + parse_float(peca.get("produzido_nok", 0), 0.0)
        + parse_float(peca.get("produzido_qualidade", 0), 0.0)
    )
    qtd_planeada = parse_float(peca.get("quantidade_pedida", 0), 0.0)
    has_ops_progress = any(operacao_qtd_total(op, fallback_done=qtd_planeada) > 0 or "concl" in norm_text(op.get("estado", "")) for op in fluxo)
    has_ops_running = any("produ" in norm_text(op.get("estado", "")) for op in fluxo)
    ever_started = (total > 0) or bool(peca.get("inicio_producao")) or has_ops_progress or has_ops_running
    has_inicio_aberto = bool(peca.get("inicio_producao")) and not bool(peca.get("fim_producao"))
    embalagem_ok = is_peca_embalada(peca)
    ops_ok = peca_operacoes_completas(peca)
    # Uma peca so fica concluida quando TODAS as operacoes do fluxo estiverem concluidas
    # e a quantidade final estiver registada.
    if ops_ok and embalagem_ok and (qtd_planeada <= 0 or total >= qtd_planeada):
        peca["estado"] = "Concluida"
        return
    if avaria_ativa and not has_ops_running:
        peca["estado"] = "Avaria"
    elif pausa_registada and ("interromp" in estado_atual or "paus" in estado_atual) and not has_ops_running:
        peca["estado"] = "Em pausa"
    elif has_ops_running or has_inicio_aberto:
        peca["estado"] = "Em producao"
    elif ever_started:
        # Já iniciou em algum momento, mas sem operação em curso neste instante.
        peca["estado"] = "Incompleta"
    else:
        peca["estado"] = "Preparacao"


def fmt_num(value):
    try:
        v = float(value)
    except Exception:
        return str(value)
    if v.is_integer():
        return str(int(v))
    return f"{v:.2f}".rstrip("0").rstrip(".")


def parse_float(val, default=0.0):
    try:
        return float(str(val).replace(",", "."))
    except Exception:
        return default


def norm_text(value):
    txt = unicodedata.normalize("NFKD", str(value or ""))
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    return txt.lower().strip()


def _nota_encomenda_last_delivery_date(ne):
    dates = []

    def _push(raw):
        txt = str(raw or "").strip()
        if txt:
            dates.append(txt[:10])

    if not isinstance(ne, dict):
        return ""
    _push(ne.get("data_ultima_entrega"))
    _push(ne.get("data_entregue"))
    for ent in ne.get("entregas", []):
        if not isinstance(ent, dict):
            continue
        _push(ent.get("data_entrega"))
        _push(ent.get("data_documento"))
    for line in ne.get("linhas", []):
        if not isinstance(line, dict):
            continue
        _push(line.get("data_entrega_real"))
        _push(line.get("data_doc_entrega"))
        for ent_l in line.get("entregas_linha", []):
            if not isinstance(ent_l, dict):
                continue
            _push(ent_l.get("data_entrega"))
            _push(ent_l.get("data_documento"))
    return max(dates) if dates else ""


def normalize_nota_encomenda_estado(ne):
    if not isinstance(ne, dict):
        return False
    changed = False
    estado_atual = str(ne.get("estado", "") or "").strip()
    estado_norm = norm_text(estado_atual)

    if "cancel" in estado_norm:
        return False

    if "convert" in estado_norm or bool(ne.get("oculta")):
        if estado_atual != "Convertida":
            ne["estado"] = "Convertida"
            changed = True
        return changed

    if bool(ne.get("_draft")):
        desired = estado_atual or "Em edicao"
        if desired != estado_atual:
            ne["estado"] = desired
            changed = True
        return changed

    linhas = [line for line in ne.get("linhas", []) if isinstance(line, dict)]
    if not linhas:
        desired = estado_atual or "Em edicao"
        if desired != estado_atual:
            ne["estado"] = desired
            changed = True
        return changed

    any_started = False
    all_done = True
    meaningful_lines = 0

    for line in linhas:
        qtd_total = max(0.0, parse_float(line.get("qtd", 0), 0))
        qtd_ent_raw = parse_float(
            line.get("qtd_entregue", qtd_total if line.get("entregue") else 0),
            0,
        )
        qtd_ent = max(0.0, qtd_ent_raw)
        flag_done = bool(line.get("entregue"))
        flag_stock = bool(line.get("_stock_in")) or bool(line.get("stock_in"))

        if qtd_total > 0:
            meaningful_lines += 1
            qtd_ent = min(qtd_total, qtd_ent)
            if flag_done and qtd_ent < (qtd_total - 1e-9):
                qtd_ent = qtd_total
            if abs(qtd_ent - qtd_ent_raw) > 1e-9:
                line["qtd_entregue"] = qtd_ent
                changed = True

            line_done = flag_done or qtd_ent >= (qtd_total - 1e-9)
            if bool(line.get("entregue")) != line_done:
                line["entregue"] = line_done
                changed = True

            any_started = any_started or flag_done or flag_stock or qtd_ent > 0
            all_done = all_done and line_done
        else:
            any_started = any_started or flag_done or flag_stock or qtd_ent > 0

    if meaningful_lines and all_done:
        desired = "Entregue"
    elif any_started:
        desired = "Parcialmente Entregue"
    elif "edi" in estado_norm and "apro" not in estado_norm:
        desired = estado_atual or "Em edicao"
    else:
        desired = "Aprovada"

    if desired != estado_atual:
        ne["estado"] = desired
        changed = True

    if desired == "Entregue":
        last_dt = _nota_encomenda_last_delivery_date(ne)
        if last_dt and not str(ne.get("data_ultima_entrega", "") or "").strip():
            ne["data_ultima_entrega"] = last_dt
            changed = True
        if last_dt and not str(ne.get("data_entregue", "") or "").strip():
            ne["data_entregue"] = last_dt
            changed = True

    return changed


def normalize_notas_encomenda(data):
    if not isinstance(data, dict):
        return False
    changed = False
    for ne in data.get("notas_encomenda", []):
        if normalize_nota_encomenda_estado(ne):
            changed = True
    return changed


def pdf_normalize_text(value):
    txt = _repair_mojibake_text(value)
    txt = txt.replace("№", "N.").replace("\u00a0", " ")
    return "".join(PDF_TEXT_REPLACEMENTS.get(ch, ch) for ch in txt)


def normalize_operacao_nome(value):
    txt = str(value or "").strip()
    n = norm_text(txt)
    if not n:
        return ""
    if "embal" in n:
        return "Embalamento"
    if "laser" in n:
        return "Corte Laser"
    if "maquin" in n:
        return "Maquinacao"
    if "quin" in n:
        return "Quinagem"
    if "serralh" in n:
        return "Serralharia"
    if "rosc" in n:
        return "Roscagem"
    if "sold" in n:
        return "Soldadura"
    if "laca" in n:
        return "Lacagem"
    if "pint" in n:
        return "Pintura"
    if "montag" in n:
        return "Montagem"
    return txt


def normalize_planeamento_operacao(value):
    txt = str(value or "").strip()
    if not txt:
        return ""
    nome = normalize_operacao_nome(txt) or txt
    n = norm_text(nome)
    if not n:
        return ""
    if "laser" in n:
        return "Corte Laser"
    if "maquin" in n:
        return "Maquinacao"
    if "quin" in n:
        return "Quinagem"
    if "serralh" in n or "sold" in n:
        return "Serralharia"
    if "rosc" in n:
        return "Roscagem"
    if "laca" in n or "pint" in n:
        return "Lacagem"
    if "montag" in n:
        return "Montagem"
    return nome


def parse_planeamento_operacoes(value):
    ordered = []
    for nome in parse_operacoes_lista(value):
        planeamento_nome = normalize_planeamento_operacao(nome)
        if not planeamento_nome:
            continue
        if planeamento_nome not in PLANEAMENTO_OPERACOES_DISPONIVEIS:
            continue
        if planeamento_nome not in ordered:
            ordered.append(planeamento_nome)
    return ordered


def parse_operacoes_lista(value):
    ops = []
    if isinstance(value, list):
        for item in value:
            if isinstance(item, dict):
                nome = normalize_operacao_nome(item.get("nome") or item.get("operacao"))
            else:
                nome = normalize_operacao_nome(item)
            if nome and nome not in ops:
                ops.append(nome)
    elif isinstance(value, dict):
        nome = normalize_operacao_nome(value.get("nome") or value.get("operacao"))
        if nome:
            ops.append(nome)
    else:
        txt = str(value or "").strip()
        if txt:
            for token in re.split(r"[+,;|/\n]+", txt):
                nome = normalize_operacao_nome(token)
                if nome and nome not in ops:
                    ops.append(nome)
    if OFF_OPERACAO_OBRIGATORIA not in ops:
        ops.append(OFF_OPERACAO_OBRIGATORIA)
    ordered = []
    for nome in OFF_OPERACOES_DISPONIVEIS:
        if nome in ops:
            ordered.append(nome)
    for nome in ops:
        if nome not in ordered:
            ordered.append(nome)
    return ordered


def normalize_orc_line_type(value):
    raw = str(value or "").strip()
    key = norm_text(raw)
    if not key:
        return ORC_LINE_TYPE_PIECE
    if key in ("pecafabricada", "peca", "fabricada", "fabricacao", "fabricado"):
        return ORC_LINE_TYPE_PIECE
    if key in ("produtostock", "produto", "stock", "componente", "componentestock"):
        return ORC_LINE_TYPE_PRODUCT
    if key in ("servicomontagem", "montagem", "servico", "servicoassemblagem", "assemblagem"):
        return ORC_LINE_TYPE_SERVICE
    if "produto" in key or "stock" in key or "component" in key:
        return ORC_LINE_TYPE_PRODUCT
    if "montag" in key or "servic" in key:
        return ORC_LINE_TYPE_SERVICE
    return ORC_LINE_TYPE_PIECE


def orc_line_type_label(value):
    tipo = normalize_orc_line_type(value)
    if tipo == ORC_LINE_TYPE_PRODUCT:
        return "Produto stock"
    if tipo == ORC_LINE_TYPE_SERVICE:
        return "Servico montagem"
    return "Peca fabricada"


def orc_line_is_piece(row):
    if isinstance(row, dict):
        return normalize_orc_line_type(row.get("tipo_item")) == ORC_LINE_TYPE_PIECE
    return normalize_orc_line_type(row) == ORC_LINE_TYPE_PIECE


def orc_line_is_product(row):
    if isinstance(row, dict):
        return normalize_orc_line_type(row.get("tipo_item")) == ORC_LINE_TYPE_PRODUCT
    return normalize_orc_line_type(row) == ORC_LINE_TYPE_PRODUCT


def orc_line_is_service(row):
    if isinstance(row, dict):
        return normalize_orc_line_type(row.get("tipo_item")) == ORC_LINE_TYPE_SERVICE
    return normalize_orc_line_type(row) == ORC_LINE_TYPE_SERVICE


def build_operacoes_fluxo(ops, existing=None):
    existing_map = {}
    for item in (existing or []):
        if not isinstance(item, dict):
            continue
        nome = normalize_operacao_nome(item.get("nome") or item.get("operacao"))
        if nome:
            existing_map[nome] = item
    fluxo = []
    for nome in parse_operacoes_lista(ops):
        old = existing_map.get(nome, {})
        estado_raw = str(old.get("estado", "") or "").strip()
        estado_norm = norm_text(estado_raw)
        if "concl" in estado_norm:
            estado = "Concluida"
        elif "incom" in estado_norm:
            estado = "Incompleta"
        elif "produ" in estado_norm or "curso" in estado_norm:
            estado = "Em producao"
        elif "prep" in estado_norm:
            estado = "Preparacao"
        elif "pend" in estado_norm:
            estado = "Pendente"
        else:
            estado = estado_raw or "Pendente"
        fluxo.append(
            {
                "nome": nome,
                "estado": estado,
                "inicio": str(old.get("inicio", "") or ""),
                "fim": str(old.get("fim", "") or ""),
                "user": str(old.get("user", "") or ""),
                "qtd_ok": _to_num(old.get("qtd_ok")) or 0.0,
                "qtd_nok": _to_num(old.get("qtd_nok")) or 0.0,
                "qtd_qual": _to_num(old.get("qtd_qual")) or 0.0,
            }
        )
    return fluxo


def peca_operacoes_fluxo(peca):
    fluxo = peca.get("operacoes_fluxo")
    if isinstance(fluxo, list) and fluxo:
        return fluxo
    ops_txt = (
        peca.get("Operacoes", "")
        or peca.get("Operações", "")
        or peca.get("OperaÃ§Ãµes", "")
        or peca.get("operacoes", "")
        or ""
    )
    return build_operacoes_fluxo(ops_txt)


def ensure_peca_operacoes(peca):
    ops_txt = (
        peca.get("Operacoes", "")
        or peca.get("Operações", "")
        or peca.get("OperaÃ§Ãµes", "")
        or peca.get("operacoes", "")
        or ""
    )
    fluxo = build_operacoes_fluxo(ops_txt, peca.get("operacoes_fluxo"))
    peca["operacoes_fluxo"] = fluxo
    peca["Operacoes"] = " + ".join([x.get("nome", "") for x in fluxo if x.get("nome")])
    peca.setdefault("hist", [])
    peca.setdefault("qtd_expedida", 0.0)
    peca.setdefault("expedicoes", [])
    return fluxo


def operacao_qtd_total(operacao, fallback_done=0.0):
    op = operacao if isinstance(operacao, dict) else {}
    total = (
        parse_float(op.get("qtd_ok", 0), 0.0)
        + parse_float(op.get("qtd_nok", 0), 0.0)
        + parse_float(op.get("qtd_qual", 0), 0.0)
    )
    if total > 0:
        return round(total, 4)
    if "concl" in norm_text(op.get("estado", "")) and fallback_done > 0:
        return round(float(fallback_done), 4)
    return 0.0


def operacao_qtd_ok(operacao, fallback_ok=0.0):
    op = operacao if isinstance(operacao, dict) else {}
    qtd_ok = parse_float(op.get("qtd_ok", 0), 0.0)
    if qtd_ok > 0:
        return round(qtd_ok, 4)
    if "concl" in norm_text(op.get("estado", "")) and fallback_ok > 0:
        return round(float(fallback_ok), 4)
    return 0.0


def operacao_input_qtd(peca, operacao_nome):
    fluxo = ensure_peca_operacoes(peca)
    target = normalize_operacao_nome(operacao_nome)
    qtd_planeada = parse_float(peca.get("quantidade_pedida", 0), 0.0)
    for index, op in enumerate(fluxo):
        nome = normalize_operacao_nome(op.get("nome"))
        if nome != target:
            continue
        if index <= 0:
            return round(qtd_planeada, 4)
        prev = fluxo[index - 1]
        prev_total = operacao_qtd_total(prev, fallback_done=qtd_planeada)
        return round(max(0.0, prev_total), 4)
    return round(qtd_planeada, 4)


def operacao_esta_concluida(peca, operacao):
    op = operacao if isinstance(operacao, dict) else {}
    nome = normalize_operacao_nome(op.get("nome"))
    if not nome:
        return False
    capacidade = operacao_input_qtd(peca, nome)
    if capacidade <= 0:
        return False
    return operacao_qtd_total(op, fallback_done=capacidade) >= capacidade


def is_peca_embalada(peca):
    for op in ensure_peca_operacoes(peca):
        if op.get("nome") == OFF_OPERACAO_OBRIGATORIA:
            return "concl" in norm_text(op.get("estado", ""))
    return False


def peca_operacoes_completas(peca):
    fluxo = ensure_peca_operacoes(peca)
    if not fluxo:
        return False
    for op in fluxo:
        nome = normalize_operacao_nome(op.get("nome"))
        if not nome:
            continue
        if not operacao_esta_concluida(peca, op):
            return False
    return True


def peca_operacao_expedicao_nome(peca):
    fluxo = ensure_peca_operacoes(peca)
    for op in fluxo:
        nome = normalize_operacao_nome(op.get("nome"))
        if nome == OFF_OPERACAO_OBRIGATORIA:
            return nome
    if fluxo:
        return normalize_operacao_nome(fluxo[-1].get("nome"))
    return OFF_OPERACAO_OBRIGATORIA


def peca_qtd_pronta_expedicao(peca):
    fluxo = ensure_peca_operacoes(peca)
    if not fluxo:
        return 0.0
    target_name = peca_operacao_expedicao_nome(peca)
    target_row = None
    for op in fluxo:
        nome = normalize_operacao_nome(op.get("nome"))
        if nome == target_name:
            target_row = op
            break
    if target_row is None:
        target_row = fluxo[-1]
        target_name = normalize_operacao_nome(target_row.get("nome"))
    fallback_ok = parse_float(peca.get("produzido_ok", 0), 0.0)
    pronta = operacao_qtd_ok(target_row, fallback_ok=fallback_ok)
    if pronta < 0:
        pronta = 0.0
    return round(pronta, 4)


def peca_qtd_disponivel_expedicao(peca):
    qtd_ok = peca_qtd_pronta_expedicao(peca)
    qtd_exp = parse_float(peca.get("qtd_expedida", 0), 0.0)
    disp = qtd_ok - qtd_exp
    if disp < 0:
        disp = 0.0
    return round(disp, 4)


def concluir_operacoes_peca(peca, operacoes, user=""):
    fluxo = ensure_peca_operacoes(peca)
    if not operacoes:
        return
    ts = now_iso()
    alvo = {normalize_operacao_nome(x) for x in operacoes}
    for op in fluxo:
        nome = normalize_operacao_nome(op.get("nome"))
        if not nome or nome not in alvo:
            continue
        if operacao_esta_concluida(peca, op):
            continue
        if not op.get("inicio"):
            op["inicio"] = ts
        op["fim"] = ts
        op["user"] = (user or "").strip()
        op["estado"] = "Concluida" if operacao_esta_concluida(peca, op) else "Incompleta"
    peca["operacoes_fluxo"] = fluxo
    peca["Operacoes"] = " + ".join([x.get("nome", "") for x in fluxo if x.get("nome")])


def peca_operacoes_pendentes(peca):
    pending = []
    for op in ensure_peca_operacoes(peca):
        nome = normalize_operacao_nome(op.get("nome"))
        if not nome:
            continue
        if operacao_esta_concluida(peca, op):
            continue
        pending.append(nome)
    return pending


def peca_operacoes_concluidas(peca):
    done = []
    for op in ensure_peca_operacoes(peca):
        nome = normalize_operacao_nome(op.get("nome"))
        if not nome:
            continue
        if operacao_esta_concluida(peca, op):
            done.append(nome)
    return done


def origem_is_materia(origem):
    t = norm_text(origem)
    return ("materia" in t) or t.startswith("mat")


def is_chapa_categoria(cat):
    return "chapa" in (cat or "").lower()


def is_linear_categoria(cat):
    c = (cat or "").lower()
    return ("tubo" in c) or ("perfil" in c) or ("viga" in c)


def is_metal_categoria(cat):
    c = (cat or "").lower()
    return is_chapa_categoria(c) or is_linear_categoria(c) or ("ipe" in c) or ("upn" in c)


def produto_modo_preco(categoria, tipo=""):
    c = norm_text(categoria)
    t = norm_text(tipo)
    if "chapa" in t or "chapa" in c:
        return "peso"
    if "tubo" in t:
        return "metros"
    if any(k in t for k in ("perfil", "viga", "ipe", "upn")):
        return "peso"
    if any(k in c for k in ("viga", "ipe", "upn")):
        return "peso"
    if "tubo" in c and "perfil" in c:
        # Categoria mista: sem tipo explicito assume tubo (m).
        return "metros"
    if "tubo" in c:
        return "metros"
    if "perfil" in c:
        return "peso"
    return "compra"


def detect_materia_formato(m):
    fmt = str((m or {}).get("formato", "")).strip()
    if fmt:
        return fmt
    mat_txt = norm_text((m or {}).get("material", ""))
    secao_txt = norm_text((m or {}).get("secao_tipo", (m or {}).get("tipo_secao", "")))
    if "nervurado" in mat_txt or secao_txt in {"nervurado", "redondo_nervurado"}:
        return "Varão nervurado"
    if "cantoneira" in mat_txt or secao_txt in {"abas_iguais", "abas_desiguais"}:
        return "Cantoneira"
    if "tubo" in mat_txt:
        return "Tubo"
    if "barra" in mat_txt or secao_txt == "chata":
        return "Barra"
    if any(k in mat_txt for k in ("perfil", "viga", "ipe", "upn")):
        return "Perfil"
    metros = parse_float((m or {}).get("metros", 0), 0)
    comp = parse_float((m or {}).get("comprimento", 0), 0)
    larg = parse_float((m or {}).get("largura", 0), 0)
    if metros > 0 and (comp <= 0 or larg <= 0):
        return "Tubo"
    return "Chapa"


def materia_preco_unitario(m):
    compra = parse_float((m or {}).get("p_compra", 0), 0)
    formato = detect_materia_formato(m)
    if formato == "Tubo":
        return parse_float((m or {}).get("metros", 0), 0) * compra
    if formato in ("Chapa", "Perfil", "Cantoneira", "Barra", "Varão nervurado"):
        return parse_float((m or {}).get("peso_unid", 0), 0) * compra
    return compra


def produto_preco_unitario(prod):
    cat = prod.get("categoria", "")
    tipo = prod.get("tipo", "")
    p = parse_float(prod.get("p_compra", 0), 0)
    modo = produto_modo_preco(cat, tipo)
    if modo == "peso":
        return parse_float(prod.get("peso_unid", 0), 0) * p
    if modo == "metros":
        return parse_float(prod.get("metros_unidade", prod.get("metros", 0)), 0) * p
    return p


def load_logo(max_width):
    # Resolve logo from branding candidates first (works with .jpg/.png and BASE_DIR).
    logo_path = get_orc_logo_path()
    if not logo_path:
        fallback = _resolve_branding_path(LOGO_PATH)
        if fallback and os.path.exists(fallback):
            logo_path = fallback
    if not logo_path or not os.path.exists(logo_path):
        return None

    # Prefer PIL so .jpg and .png are both supported.
    try:
        from PIL import Image, ImageTk
        img = Image.open(logo_path).convert("RGBA")
        if max_width and img.width > max_width:
            ratio = max_width / float(max(1, img.width))
            new_h = max(1, int(img.height * ratio))
            img = img.resize((max_width, new_h))
        return ImageTk.PhotoImage(img)
    except Exception:
        # Fallback for environments without PIL (typically supports .png).
        try:
            img = PhotoImage(file=logo_path)
            w = img.width()
            if max_width and w > max_width:
                factor = max(1, w // max_width)
                img = img.subsample(factor, factor)
            return img
        except Exception:
            return None


def week_start(date_obj):
    return date_obj - timedelta(days=date_obj.weekday())


def time_to_minutes(hhmm):
    h, m = hhmm.split(":")
    return int(h) * 60 + int(m)


def minutes_to_time(minutes):
    h = minutes // 60
    m = minutes % 60
    return f"{h:02d}:{m:02d}"


def encomenda_espessuras(enc):
    esp = []
    for m in enc.get("materiais", []):
        for e in m.get("espessuras", []):
            esp.append(e.get("espessura", ""))
    return esp


def encomenda_pecas(enc):
    pecas = []
    for m in enc.get("materiais", []):
        for e in m.get("espessuras", []):
            pecas.extend(e.get("pecas", []))
    return pecas


def encomenda_montagem_itens(enc):
    rows = []
    for item in list((enc or {}).get("montagem_itens", []) or []):
        if isinstance(item, dict):
            rows.append(item)
    return rows


def encomenda_montagem_estado(enc):
    stock_items = [
        row
        for row in encomenda_montagem_itens(enc)
        if normalize_orc_line_type(row.get("tipo_item")) == ORC_LINE_TYPE_PRODUCT
        or (
            normalize_orc_line_type(row.get("tipo_item")) == ORC_LINE_TYPE_PIECE
            and (
                str(row.get("stock_item_kind", "") or "").strip() == "raw_material"
                or str(row.get("stock_material_id", "") or "").strip()
            )
        )
    ]
    if not stock_items:
        return "Nao aplicavel"
    pending = []
    for row in stock_items:
        plan = parse_float(row.get("qtd_planeada", row.get("qtd", 0)), 0.0)
        done = parse_float(row.get("qtd_consumida", 0), 0.0)
        if done + 1e-9 < plan:
            pending.append(row)
    if pending:
        return "Pendente"
    return "Consumida"


def encomenda_montagem_tempo_min(enc):
    items = encomenda_montagem_itens(enc)
    explicit_total = 0.0
    explicit_found = False
    for row in items:
        qtd = parse_float(row.get("qtd_planeada", row.get("qtd", 0)), 0.0)
        if row.get("tempo_total_min") not in (None, ""):
            explicit_total += parse_float(row.get("tempo_total_min"), 0.0)
            explicit_found = True
            continue
        if row.get("tempo_unit_min") not in (None, ""):
            explicit_total += parse_float(row.get("tempo_unit_min"), 0.0) * max(qtd, 0.0)
            explicit_found = True
            continue
        if row.get("tempo_peca_min") not in (None, ""):
            explicit_total += parse_float(row.get("tempo_peca_min"), 0.0) * max(qtd, 0.0)
            explicit_found = True
    if explicit_found and explicit_total > 0:
        return round(explicit_total, 2)

    total_estimado = parse_float((enc or {}).get("tempo_estimado", (enc or {}).get("tempo", 0)), 0.0)
    total_producao = 0.0
    for mat in list((enc or {}).get("materiais", []) or []):
        for esp in list((mat or {}).get("espessuras", []) or []):
            total_producao += parse_float((esp or {}).get("tempo_min", 0), 0.0)
    restante = max(0.0, total_estimado - total_producao)
    return round(restante, 2)


def encomenda_montagem_resumo(enc):
    items = encomenda_montagem_itens(enc)
    conjuntos = []
    seen = set()
    for row in items:
        nome = str((row or {}).get("conjunto_nome", "") or (row or {}).get("conjunto_codigo", "") or "").strip()
        if nome and nome not in seen:
            seen.add(nome)
            conjuntos.append(nome)
    parts = []
    if conjuntos:
        parts.append(", ".join(conjuntos[:2]) + ("..." if len(conjuntos) > 2 else ""))
    produtos = sum(1 for row in items if normalize_orc_line_type((row or {}).get("tipo_item")) == ORC_LINE_TYPE_PRODUCT)
    materias = sum(
        1
        for row in items
        if normalize_orc_line_type((row or {}).get("tipo_item")) == ORC_LINE_TYPE_PIECE
        and (
            str((row or {}).get("stock_item_kind", "") or "").strip() == "raw_material"
            or str((row or {}).get("stock_material_id", "") or "").strip()
        )
    )
    servicos = sum(1 for row in items if normalize_orc_line_type((row or {}).get("tipo_item")) == ORC_LINE_TYPE_SERVICE)
    if produtos:
        parts.append(f"{produtos} comp.")
    if materias:
        parts.append(f"{materias} MP")
    if servicos:
        parts.append(f"{servicos} serv.")
    return " | ".join(parts) if parts else "Montagem final"


def update_estado_encomenda_por_espessuras(enc):
    mats = enc.get("materiais", [])
    montagem_estado = encomenda_montagem_estado(enc)
    if not mats:
        if montagem_estado == "Pendente":
            enc["estado"] = "Montagem"
        elif montagem_estado == "Consumida":
            enc["estado"] = "Concluida"
        else:
            enc["estado"] = "Preparacao"
        return
    # Recalcula o estado de cada espessura a partir das peças para evitar
    # estados "Concluida" desatualizados quando existem OPP interrompidas/pausadas.
    for m in mats:
        for e in m.get("espessuras", []) or []:
            pecas = e.get("pecas", []) or []
            if not pecas:
                e["estado"] = "Preparacao"
                continue
            p_norm = [norm_text(p.get("estado", "")) for p in pecas]
            has_avaria = any("avari" in s for s in p_norm)
            has_running = any(("produ" in s) and ("pausad" not in s) and ("interromp" not in s) for s in p_norm)
            all_concl = bool(p_norm) and all("concl" in s for s in p_norm)
            all_prepar = bool(p_norm) and all("prepar" in s for s in p_norm)
            has_started = any(
                ("produ" in s) or ("concl" in s) or ("incomplet" in s) or ("interromp" in s) or ("paus" in s) or ("avari" in s)
                for s in p_norm
            ) or bool(e.get("inicio_producao"))
            if all_concl:
                e["estado"] = "Concluida"
            elif has_avaria:
                e["estado"] = "Avaria"
            elif has_running:
                e["estado"] = "Em producao"
            elif has_started and (not all_prepar):
                e["estado"] = "Em pausa"
            else:
                e["estado"] = "Preparacao"
    for m in mats:
        esp_list = m.get("espessuras", []) or []
        pecas_m = [p for e in esp_list for p in (e.get("pecas", []) or [])]
        if not pecas_m:
            m["estado"] = "Preparacao"
            continue
        p_norm_m = [norm_text(p.get("estado", "")) for p in pecas_m]
        m_has_avaria = any("avari" in s for s in p_norm_m)
        m_has_produ = any(("produ" in s) and ("paus" not in s) for s in p_norm_m)
        m_all_concl = bool(p_norm_m) and all("concl" in s for s in p_norm_m)
        m_has_pause = any(("interromp" in s) or ("incomplet" in s) or ("paus" in s) or ("avari" in s) for s in p_norm_m)
        if m_has_avaria:
            m["estado"] = "Avaria"
        elif m_has_produ:
            m["estado"] = "Em producao"
        elif m_all_concl:
            m["estado"] = "Concluida"
        elif m_has_pause:
            m["estado"] = "Em pausa"
        else:
            m["estado"] = "Preparacao"

    pecas_all = encomenda_pecas(enc)
    p_norm_all = [norm_text(p.get("estado", "")) for p in pecas_all]
    has_avaria = any("avari" in s for s in p_norm_all)
    has_produ = any(("produ" in s) and ("paus" not in s) for s in p_norm_all)
    all_concl = bool(p_norm_all) and all("concl" in s for s in p_norm_all)
    has_pause = any(("interromp" in s) or ("incomplet" in s) or ("paus" in s) or ("avari" in s) for s in p_norm_all)
    if has_avaria:
        enc["estado"] = "Avaria"
    elif has_produ:
        enc["estado"] = "Em producao"
    elif all_concl:
        enc["estado"] = "Montagem" if montagem_estado == "Pendente" else "Concluida"
    elif has_pause:
        enc["estado"] = "Em pausa"
    else:
        enc["estado"] = "Preparacao"

    # Metricas globais de tempo para analise
    tempos_esp = []
    tempos_peca = []
    inicios = []
    fins = []
    for m in mats:
        for e in m.get("espessuras", []):
            tesp = parse_float(e.get("tempo_producao_min"), 0.0)
            if tesp > 0:
                tempos_esp.append(tesp)
            if e.get("inicio_producao"):
                inicios.append(str(e.get("inicio_producao")))
            if e.get("fim_producao"):
                fins.append(str(e.get("fim_producao")))
            for p in e.get("pecas", []):
                tpeca = parse_float(p.get("tempo_producao_min"), 0.0)
                if tpeca > 0:
                    tempos_peca.append(tpeca)
                if p.get("inicio_producao"):
                    inicios.append(str(p.get("inicio_producao")))
                if p.get("fim_producao"):
                    fins.append(str(p.get("fim_producao")))
    if tempos_esp:
        enc["tempo_espessuras_min"] = round(sum(tempos_esp), 2)
    if tempos_peca:
        enc["tempo_pecas_min"] = round(sum(tempos_peca), 2)
    if inicios and not enc.get("inicio_producao"):
        enc["inicio_producao"] = min(inicios)
    if enc.get("estado") == "Concluida" and fins:
        enc["fim_producao"] = max(fins)
        dur = iso_diff_minutes(enc.get("inicio_producao"), enc.get("fim_producao"))
        if dur is not None:
            enc["tempo_producao_min"] = dur
    try:
        update_estado_expedicao_encomenda(enc)
    except Exception:
        pass


def update_estado_expedicao_encomenda(enc):
    total_pronta = 0.0
    total_exp = 0.0
    has_open_pieces = False
    for p in encomenda_pecas(enc):
        ensure_peca_operacoes(p)
        total_pronta += max(0.0, peca_qtd_pronta_expedicao(p))
        total_exp += parse_float(p.get("qtd_expedida", 0), 0.0)
        estado_norm = norm_text(p.get("estado", ""))
        if "concl" not in estado_norm and "cancel" not in estado_norm:
            has_open_pieces = True
    enc["qtd_pronta_expedicao"] = round(total_pronta, 4)
    enc["qtd_expedida"] = round(total_exp, 4)
    if total_exp <= 0:
        enc["estado_expedicao"] = "Não expedida"
    elif total_exp + 1e-9 < total_pronta:
        enc["estado_expedicao"] = "Parcialmente expedida"
    elif has_open_pieces:
        enc["estado_expedicao"] = "Parcialmente expedida"
    else:
        enc["estado_expedicao"] = "Totalmente expedida"


class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.user = None
        self.loaded_data = None
        use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_LOGIN", "1") != "0"
        if use_custom:
            self._build_custom_login()
        else:
            self._build_classic_login()

    def _trial_hint_text(self):
        try:
            status = get_trial_status()
        except Exception:
            return ""
        company = str(status.get("company_name", "") or "").strip()
        owner = str(status.get("owner_username", "") or "").strip()
        days_remaining = status.get("days_remaining", None)
        if bool(status.get("enabled", False)):
            if bool(status.get("blocking", False)):
                return f"Trial bloqueado{' | proprietario' if owner else ''}"
            parts = ["Trial ativo"]
            if company:
                parts.append(company)
            if days_remaining is not None:
                parts.append(f"{int(days_remaining)} dia(s)")
            return " | ".join(parts)
        return "Modo normal"

    def _build_classic_login(self):
        self.win = Toplevel(self.root)
        self.win.title("luGEST - Login")
        self.win.geometry("560x560")
        self.win.resizable(False, False)
        self.win.configure(bg="#ffffff")
        apply_window_icon(self.win)
        self.win.grab_set()

        card = Frame(self.win, bg="white", highlightbackground="#e7cfd3", highlightthickness=1)
        card.place(relx=0.5, rely=0.5, anchor="center", width=450, height=455)
        header = Frame(card, bg="white")
        header.pack(fill="x", pady=(26, 12))
        self.logo_img = None
        try:
            logo_path = get_orc_logo_path()
            if logo_path:
                try:
                    from PIL import Image, ImageTk
                    img = Image.open(logo_path)
                    img = img.resize((170, 74))
                    self.logo_img = ImageTk.PhotoImage(img)
                    Label(header, image=self.logo_img, bg="white").pack()
                except Exception:
                    pass
        except Exception:
            pass
        Label(header, text="luGEST", font=_ui_font(28, "bold"), bg="white", fg="#7a0f1a").pack()
        Label(header, text="Sistema de Gestão de Produção", font=_ui_font(12), bg="white", fg="#4c5d73").pack(pady=(4, 0))
        trial_hint = self._trial_hint_text()
        if trial_hint:
            Label(header, text=trial_hint, font=_ui_font(10, "bold"), bg="white", fg="#7a0f1a").pack(pady=(8, 0))
        Frame(card, bg="#e6ecf5", height=1).pack(fill="x", padx=24, pady=(8, 14))

        Label(card, text="Utilizador", font=_ui_font(11, "bold"), bg="white", fg="#1f3552").pack(anchor="w", padx=32)
        self.username = StringVar()
        Entry(card, textvariable=self.username, font=_ui_font(12), relief="solid", bd=1).pack(fill="x", padx=32, pady=(6, 14), ipady=6)

        Label(card, text="Password", font=_ui_font(11, "bold"), bg="white", fg="#1f3552").pack(anchor="w", padx=32)
        self.password = StringVar()
        pw_row = Frame(card, bg="white")
        pw_row.pack(fill="x", padx=32, pady=(6, 10))
        self._pw_visible = False
        self.pw_entry = Entry(pw_row, textvariable=self.password, show="*", font=_ui_font(12), relief="solid", bd=1)
        self.pw_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.pw_toggle = Button(
            pw_row,
            text="Mostrar",
            width=8,
            command=self.toggle_password,
            relief="flat",
            bd=0,
            bg="#e8f0fb",
            fg="#7a0f1a",
            activebackground="#fde2e4",
            activeforeground="#7a0f1a",
        )
        self.pw_toggle.pack(side="left", padx=(8, 0))

        btns = Frame(card, bg="white")
        btns.pack(pady=(18, 10))
        Button(
            btns,
            text="Entrar",
            command=self.on_login,
            width=14,
            relief="flat",
            bd=0,
            bg=CTK_PRIMARY_RED,
            fg="white",
            activebackground=CTK_PRIMARY_RED_HOVER,
            activeforeground="white",
            font=_ui_font(11, "bold"),
        ).pack(side="left", padx=6, ipady=4)
        Button(
            btns,
            text="Sair",
            command=self.on_exit,
            width=14,
            relief="flat",
            bd=0,
            bg="#b42318",
            fg="white",
            activebackground="#8f1d14",
            activeforeground="white",
            font=_ui_font(11, "bold"),
        ).pack(side="left", padx=6, ipady=4)

        self.win.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.win.bind("<Return>", lambda _: self.on_login())
        self.win.update_idletasks()
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        ww = self.win.winfo_width()
        wh = self.win.winfo_height()
        x = max(0, (sw - ww) // 2)
        y = max(0, (sh - wh) // 2 - 20)
        self.win.geometry(f"{ww}x{wh}+{x}+{y}")

    def _build_custom_login(self):
        ctk.set_appearance_mode("light")
        self.win = ctk.CTkToplevel(self.root)
        self.win.title("luGEST - Login")
        self.win.geometry("560x560")
        self.win.resizable(False, False)
        self.win.configure(fg_color="#ffffff")
        apply_window_icon(self.win)
        self.win.grab_set()

        card = ctk.CTkFrame(
            self.win,
            corner_radius=14,
            width=460,
            height=470,
            fg_color="#ffffff",
            border_width=1,
            border_color="#e7cfd3",
        )
        card.place(relx=0.5, rely=0.5, anchor="center")

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", pady=(28, 12))
        self.logo_img = None
        try:
            logo_path = get_orc_logo_path()
            if logo_path:
                from PIL import Image, ImageTk
                img = Image.open(logo_path).resize((180, 78))
                self.logo_img = ImageTk.PhotoImage(img)
                ctk.CTkLabel(header, image=self.logo_img, text="").pack()
        except Exception:
            pass

        ctk.CTkLabel(header, text="luGEST", font=_ui_font(30, "bold"), text_color="#7a0f1a").pack()
        ctk.CTkLabel(header, text="Sistema de Gestão de Produção", font=_ui_font(12), text_color="#4c5d73").pack(pady=(4, 0))
        trial_hint = self._trial_hint_text()
        if trial_hint:
            ctk.CTkLabel(header, text=trial_hint, font=_ui_font(10, "bold"), text_color="#7a0f1a", wraplength=360, justify="center").pack(pady=(8, 0))
        ctk.CTkFrame(card, fg_color="#e6ecf5", height=1).pack(fill="x", padx=24, pady=(8, 14))

        entry_cfg = {
            "height": 40,
            "corner_radius": 10,
            "fg_color": "#f8fbff",
            "border_color": "#d8b9bf",
            "text_color": "#1f2937",
            "font": _ui_font(12),
        }
        ctk.CTkLabel(card, text="Utilizador", font=_ui_font(11, "bold"), text_color="#1f3552").pack(anchor="w", padx=32)
        self.username = StringVar()
        ctk.CTkEntry(card, textvariable=self.username, **entry_cfg).pack(fill="x", padx=32, pady=(6, 14))

        ctk.CTkLabel(card, text="Password", font=_ui_font(11, "bold"), text_color="#1f3552").pack(anchor="w", padx=32)
        self.password = StringVar()
        pw_row = ctk.CTkFrame(card, fg_color="transparent")
        pw_row.pack(fill="x", padx=32, pady=(6, 10))
        self._pw_visible = False
        self.pw_entry = ctk.CTkEntry(pw_row, textvariable=self.password, show="*", width=240, **entry_cfg)
        self.pw_entry.pack(side="left", fill="x", expand=True)
        self.pw_toggle = ctk.CTkButton(
            pw_row,
            text="Mostrar",
            width=86,
            height=40,
            corner_radius=10,
            command=self.toggle_password,
            fg_color="#e8f0fb",
            hover_color="#fde2e4",
            text_color="#7a0f1a",
        )
        self.pw_toggle.pack(side="left", padx=(8, 0))

        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.pack(pady=(18, 10))
        ctk.CTkButton(
            btns,
            text="Entrar",
            command=self.on_login,
            width=150,
            height=42,
            corner_radius=11,
            font=_ui_font(13, "bold"),
            fg_color=CTK_PRIMARY_RED,
            hover_color=CTK_PRIMARY_RED_HOVER,
            text_color="#ffffff",
            border_width=0,
        ).pack(side="left", padx=6)
        ctk.CTkButton(
            btns,
            text="Sair",
            command=self.on_exit,
            width=150,
            height=42,
            corner_radius=11,
            font=_ui_font(13, "bold"),
            fg_color="#b42318",
            hover_color="#8f1d14",
            text_color="#ffffff",
            border_width=0,
        ).pack(side="left", padx=6)

        self.win.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.win.bind("<Return>", lambda _: self.on_login())
        self.win.update_idletasks()
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        ww = self.win.winfo_width()
        wh = self.win.winfo_height()
        x = max(0, (sw - ww) // 2)
        y = max(0, (sh - wh) // 2 - 20)
        self.win.geometry(f"{ww}x{wh}+{x}+{y}")

    def toggle_password(self):
        self._pw_visible = not self._pw_visible
        self.pw_entry.configure(show="" if self._pw_visible else "*")
        self.pw_toggle.configure(text="Ocultar" if self._pw_visible else "Mostrar")

    def on_login(self):
        if not isinstance(self.loaded_data, dict):
            try:
                self.loaded_data = load_data()
            except Exception as exc:
                messagebox.showerror("Ligação MySQL", str(exc))
                return
        data = self.loaded_data
        u = self.username.get().strip()
        p = self.password.get().strip()
        try:
            owner_session = ensure_trial_login_session(u, p, allow_owner=True)
        except Exception as exc:
            messagebox.showerror("Licenciamento", str(exc))
            return
        if isinstance(owner_session, dict):
            self.user = owner_session
            self.win.destroy()
            return
        user = authenticate_local_user(data, u, p)
        if isinstance(user, dict):
            try:
                ensure_trial_login_session(u, p, allow_owner=False)
                touch_trial_success(u, owner=False)
            except Exception as exc:
                messagebox.showerror("Licenciamento", str(exc))
                return
            self.user = build_authenticated_user_session(user, p)
            self.win.destroy()
            return
        messagebox.showerror("Erro", "Credenciais invalidas")

    def on_exit(self):
        self.root.destroy()


class App:
    def __init__(self, root, user, prefetched_data=None):
        global _RUNTIME_DATA_REF
        self.root = root
        self.user = user
        apply_primary_theme_color(get_branding_config().get("primary_color", DEFAULT_PRIMARY_RED))
        self.CTK_PRIMARY_RED = CTK_PRIMARY_RED
        self.CTK_PRIMARY_RED_HOVER = CTK_PRIMARY_RED_HOVER
        self.THEME_HEADER_BG = THEME_HEADER_BG
        self.THEME_HEADER_ACTIVE = THEME_HEADER_ACTIVE
        self.THEME_SELECT_BG = THEME_SELECT_BG
        self.THEME_SELECT_FG = THEME_SELECT_FG
        _set_ctk_button_defaults_red()
        self.data = prefetched_data if isinstance(prefetched_data, dict) else load_data()
        _RUNTIME_DATA_REF = self.data
        _set_last_save_fingerprint(self.data)
        ui_build_blocks.configure(globals())
        clientes_actions.configure(globals())
        materia_actions.configure(globals())
        encomendas_actions.configure(globals())
        qualidade_actions.configure(globals())
        orc_actions.configure(globals())
        operador_ordens_actions.configure(globals())
        ne_expedicao_actions.configure(globals())
        plan_actions.configure(globals())
        produtos_actions.configure(globals())
        app_misc_actions.configure(globals())
        apply_window_icon(self.root)
        self.root.title(f"LU-GEST - {user['role']}")
        self.root.geometry("1100x700")
        self.root.configure(bg="#ffffff")

        self.setup_styles()

        self.nb = ttk.Notebook(root, style="Custom.TNotebook")
        self.nb.pack(fill="both", expand=True, padx=10, pady=(10, 0))
        self.menu_only_mode = True

        self.tab_menu = ttk.Frame(self.nb, style="Menu.TFrame")
        self.tab_clientes = ttk.Frame(self.nb)
        self.tab_materia = ttk.Frame(self.nb)
        self.tab_encomendas = ttk.Frame(self.nb)
        self.tab_plano = ttk.Frame(self.nb)
        self.tab_qualidade = ttk.Frame(self.nb)
        self.tab_orc = ttk.Frame(self.nb)
        self.tab_ordens = ttk.Frame(self.nb)
        self.tab_export = ttk.Frame(self.nb)
        self.tab_expedicao = ttk.Frame(self.nb)
        self.tab_produtos = ttk.Frame(self.nb)
        self.tab_ne = ttk.Frame(self.nb)
        self.tab_operador = ttk.Frame(self.nb)
        self.op_menu = None
        self.op_manage_win = None
        self.menu_back_bar = None

        self.nb.add(self.tab_menu, text="Menu")
        self.nb.add(self.tab_clientes, text="Clientes")
        self.nb.add(self.tab_materia, text="Matéria-Prima")
        self.nb.add(self.tab_encomendas, text="Encomendas")
        self.nb.add(self.tab_plano, text="Plano de Produção")
        self.nb.add(self.tab_qualidade, text="Qualidade")
        self.nb.add(self.tab_orc, text="Orçamentação")
        self.nb.add(self.tab_ordens, text="Ordens de Fabrico")
        self.nb.add(self.tab_export, text="Exportações")
        self.nb.add(self.tab_expedicao, text="Expedição")
        self.nb.add(self.tab_produtos, text="Produtos")
        self.nb.add(self.tab_ne, text="Notas Encomenda")
        self.nb.add(self.tab_operador, text="Operador")

        self._tab_builders = {
            str(self.tab_clientes): self.build_clientes,
            str(self.tab_materia): self.build_materia,
            str(self.tab_encomendas): self.build_encomendas,
            str(self.tab_plano): self.build_plano,
            str(self.tab_qualidade): self.build_qualidade,
            str(self.tab_orc): self.build_orc,
            str(self.tab_ordens): self.build_ordens,
            str(self.tab_export): self.build_export,
            str(self.tab_expedicao): self.build_expedicao,
            str(self.tab_produtos): self.build_produtos,
            str(self.tab_ne): self.build_ne,
            str(self.tab_operador): self.build_operador,
        }
        self._built_tabs = set()
        self._tab_warmup_queue = []
        self._tab_warmup_job = None
        self._dirty_tabs = set()
        self._tree_fill_tokens = {}

        self.selected_encomenda_numero = None
        self.selected_material = None
        self.selected_espessura = None
        self.selected_orc_numero = None
        self._ui_debounce_jobs = {}
        self._orc_refreshing = False
        self.build_menu_dashboard()
        self._built_tabs.add(str(self.tab_menu))

        self.nb.bind("<<NotebookTabChanged>>", self.on_notebook_tab_changed, add="+")
        self.apply_permissions()
        self.root.after(80, self.refresh_menu_dashboard)
        self.bind_shortcuts()
        self.selected_encomenda_numero = None
        if self.user.get("role") != "Operador":
            try:
                self.nb.select(self.tab_menu)
            except Exception:
                pass
        self.ensure_tab_built_by_widget(self.nb.select())
        self.update_menu_back_button()
        self._runtime_year_token = str(datetime.now().year)
        try:
            self.root.after(2000, self._check_year_rollover)
        except Exception:
            pass
        try:
            self.root.after(900, self.schedule_tab_warmup)
        except Exception:
            pass
        try:
            self.root.state("zoomed")
        except Exception:
            try:
                sw = self.root.winfo_screenwidth()
                sh = self.root.winfo_screenheight()
                self.root.geometry(f"{sw}x{sh}+0+0")
            except Exception:
                pass

    def escolher_operacoes_fluxo(self, current_text="", parent=None):
        return app_misc_actions.escolher_operacoes_fluxo(self, current_text, parent=parent)


    def escolher_operacoes_concluir(self, peca, parent=None):
        return app_misc_actions.escolher_operacoes_concluir(self, peca, parent=parent)


    def build_menu_dashboard(self):
        return menu_rooting.build_menu_dashboard(
            self,
            custom_tk_available=CUSTOM_TK_AVAILABLE,
            load_logo_fn=load_logo,
            get_orc_logo_path_fn=get_orc_logo_path,
        )

    def create_menu_button(self, parent, text, command, danger=False, compact=True):
        return menu_rooting.create_menu_button(self, parent, text, command, danger=danger, compact=compact)

    def refresh_menu_dashboard(self):
        return menu_rooting.refresh_menu_dashboard(self)

    def refresh_menu_dashboard_stats(self):
        return menu_rooting.refresh_menu_dashboard_stats(self, norm_text_fn=norm_text, parse_float_fn=parse_float)

    def refresh_menu_dashboard_chart(self):
        return menu_rooting.refresh_menu_dashboard_chart(self, norm_text_fn=norm_text, parse_float_fn=parse_float)

    def navigate_to_tab(self, tab):
        return menu_rooting.navigate_to_tab(self, tab)

    def go_to_main_menu(self):
        return menu_rooting.go_to_main_menu(self)

    def on_notebook_tab_changed(self, _event=None):
        return menu_rooting.on_notebook_tab_changed(self, _event)

    def ensure_tab_built_by_widget(self, tab_widget, background=False):
        key = str(tab_widget or "")
        if not key or key in self._built_tabs:
            return
        builder = self._tab_builders.get(key)
        if builder is None:
            return
        if not background:
            try:
                self.root.configure(cursor="watch")
                self.root.update_idletasks()
            except Exception:
                pass
        try:
            builder()
            self._built_tabs.add(key)
        finally:
            if not background:
                try:
                    self.root.configure(cursor="")
                    self.root.update_idletasks()
                except Exception:
                    pass

    def schedule_tab_warmup(self):
        warm_tabs = [
            self.tab_operador,
            self.tab_materia,
            self.tab_encomendas,
            self.tab_plano,
            self.tab_ne,
            self.tab_ordens,
        ]
        self._tab_warmup_queue = [str(tab) for tab in warm_tabs if str(tab) not in self._built_tabs]
        if self._tab_warmup_job is None:
            try:
                self._tab_warmup_job = self.root.after(60, self._warm_next_tab)
            except Exception:
                self._warm_next_tab()

    def _warm_next_tab(self):
        self._tab_warmup_job = None
        if not self._tab_warmup_queue:
            return
        key = self._tab_warmup_queue.pop(0)
        try:
            self.ensure_tab_built_by_widget(key, background=True)
        except Exception:
            pass
        if self._tab_warmup_queue:
            try:
                self._tab_warmup_job = self.root.after(120, self._warm_next_tab)
            except Exception:
                self._warm_next_tab()

    def mark_tab_dirty(self, *tab_names):
        for name in tab_names:
            txt = str(name or "").strip().lower()
            if txt:
                self._dirty_tabs.add(txt)

    def refresh_selected_tab_if_dirty(self, tab_widget=None):
        key = str(tab_widget or self.nb.select() or "")
        mapping = {
            str(self.tab_materia): ("materia", self.refresh_materia),
            str(self.tab_encomendas): ("encomendas", self.refresh_encomendas),
            str(self.tab_plano): ("plano", self.refresh_plano),
            str(self.tab_expedicao): ("expedicao", self.refresh_expedicao),
            str(self.tab_operador): ("operador", self.refresh_operador),
            str(self.tab_ordens): ("ordens", self.refresh_ordens),
            str(self.tab_ne): ("ne", self.refresh_ne),
        }
        target = mapping.get(key)
        if not target:
            return
        name, callback = target
        if name not in self._dirty_tabs:
            return
        self._dirty_tabs.discard(name)
        try:
            callback()
        except Exception:
            self._dirty_tabs.add(name)

    def fill_treeview_in_batches(self, tree, rows, token_key, chunk_size=140, on_done=None):
        if tree is None:
            try:
                if callable(on_done):
                    on_done()
            except Exception:
                pass
            return
        try:
            children = tree.get_children()
            if children:
                tree.delete(*children)
        except Exception:
            pass
        try:
            token = int(self._tree_fill_tokens.get(token_key, 0)) + 1
        except Exception:
            token = 1
        self._tree_fill_tokens[token_key] = token
        total = len(rows or [])

        def _finish():
            try:
                if callable(on_done):
                    on_done()
            except Exception:
                pass

        if total <= 0:
            _finish()
            return

        def _run(start=0):
            if self._tree_fill_tokens.get(token_key) != token:
                return
            try:
                if not tree.winfo_exists():
                    return
            except Exception:
                return
            end = min(total, start + max(25, int(chunk_size or 140)))
            for values, tags in rows[start:end]:
                if self._tree_fill_tokens.get(token_key) != token:
                    return
                try:
                    tree.insert("", END, values=values, tags=tuple(tags or ()))
                except Exception:
                    return
            if end < total:
                try:
                    self.root.after(1, lambda: _run(end))
                except Exception:
                    _run(end)
            else:
                _finish()

        try:
            self.root.after_idle(_run)
        except Exception:
            _run()

    def update_menu_back_button(self):
        return menu_rooting.update_menu_back_button(self)

    def apply_permissions(self):
        return app_misc_actions.apply_permissions(self)


    def apply_menu_only_mode(self):
        return menu_rooting.apply_menu_only_mode(self)

    def show_operador_menu(self):
        return menu_rooting.show_operador_menu(self)

    def setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#ffffff")
        style.configure("TLabelframe", background="#ffffff", borderwidth=1, relief="solid")
        style.configure("TButton", padding=(12, 6), font=_ui_font(11))
        style.configure("TButton", foreground="#0f172a")
        style.map("TButton", foreground=[("disabled", "#64748b"), ("active", "#0f172a"), ("!disabled", "#0f172a")])
        style.configure("Op.TButton", padding=(14, 8), font=_ui_font(11, "bold"), foreground="#0f172a", background="#e5e7eb")
        style.map("Op.TButton", background=[("active", "#d7dde4"), ("!disabled", "#e5e7eb")])
        style.configure("TLabel", font=_ui_font(11))
        style.configure("TEntry", font=_ui_font(11))
        style.configure("TCombobox", font=_ui_font(11))
        style.configure(
            "Vertical.TScrollbar",
            gripcount=0,
            background="#cbd5e1",
            darkcolor="#cbd5e1",
            lightcolor="#cbd5e1",
            troughcolor="#f8fafc",
            bordercolor="#e2e8f0",
            arrowcolor="#64748b",
            relief="flat",
        )
        style.map(
            "Vertical.TScrollbar",
            background=[("active", "#94a3b8"), ("pressed", "#64748b")],
            arrowcolor=[("active", "#334155"), ("pressed", "#1e293b")],
        )
        style.configure(
            "Horizontal.TScrollbar",
            gripcount=0,
            background="#cbd5e1",
            darkcolor="#cbd5e1",
            lightcolor="#cbd5e1",
            troughcolor="#f8fafc",
            bordercolor="#e2e8f0",
            arrowcolor="#64748b",
            relief="flat",
        )
        style.map(
            "Horizontal.TScrollbar",
            background=[("active", "#94a3b8"), ("pressed", "#64748b")],
            arrowcolor=[("active", "#334155"), ("pressed", "#1e293b")],
        )
        style.configure("Menu.TFrame", background="#ffffff")
        style.configure("Menu.TLabel", background="#ffffff", font=_ui_font(11))
        style.configure("Menu.Title.TLabel", background="#ffffff", font=_ui_font(24, "bold"), foreground="#7a0f1a")
        style.configure("Menu.Sub.TLabel", background="#ffffff", font=_ui_font(12), foreground="#4b5563")
        style.configure("Treeview", rowheight=28, font=_ui_font(11))
        style.configure("Treeview.Heading", font=_ui_font(11, "bold"))
        style.configure("TLabelframe.Label", font=_ui_font(11, "bold"), background="#ffffff")
        style.configure("Custom.TNotebook", background="#ffffff", borderwidth=0, tabmargins=(4, 4, 4, 0))
        style.configure(
            "Custom.TNotebook.Tab",
            padding=(14, 6),
            font=_ui_font(11, "bold"),
            background="#e5e7eb",
            foreground="#0f172a",
        )
        style.map(
            "Custom.TNotebook.Tab",
            background=[("selected", "#d7dde4"), ("active", "#e2e8f0")],
            foreground=[("selected", "#0f172a")],
        )
        style.configure("MenuOnly.TNotebook", background="#ffffff", borderwidth=0, tabmargins=(0, 0, 0, 0))
        try:
            style.layout("MenuOnly.TNotebook.Tab", [])
        except Exception:
            pass
        style.configure("MenuOnly.TNotebook.Tab", padding=(0, 0), font=_ui_font(1))
        self.btn_cfg = {
            "font": _ui_font(11, "bold"),
            "bg": "#e5e7eb",
            "fg": "#0f172a",
            "activebackground": "#d7dde4",
            "activeforeground": "#0f172a",
            "relief": "raised",
            "bd": 1,
        }

    def bind_shortcuts(self):
        self.root.bind("<F2>", self.on_shortcut_new)
        self.root.bind("<F3>", self.on_shortcut_edit)
        self.root.bind("<Delete>", self.on_shortcut_delete)
        self.root.bind("<F5>", self.on_shortcut_refresh)

    def debounce_call(self, key, callback, delay_ms=180):
        jobs = getattr(self, "_ui_debounce_jobs", None)
        if jobs is None:
            jobs = {}
            self._ui_debounce_jobs = jobs
        prev = jobs.get(key)
        if prev is not None:
            try:
                self.root.after_cancel(prev)
            except Exception:
                pass
        def _run():
            jobs.pop(key, None)
            try:
                callback()
            except Exception:
                pass
        try:
            jobs[key] = self.root.after(max(1, int(delay_ms)), _run)
        except Exception:
            _run()

    def bind_clear_on_empty(self, tree, clear_fn):
        def handler(event):
            if not tree.identify_row(event.y):
                tree.selection_remove(tree.selection())
                clear_fn()
        tree.bind("<Button-1>", handler, add="+")

    def on_shortcut_new(self, _):
        return plan_actions.on_shortcut_new(self, _)


    def on_shortcut_edit(self, _):
        tab = self.nb.select()
        if tab == str(self.tab_clientes):
            self.edit_cliente()
        elif tab == str(self.tab_materia):
            self.edit_material()

    def on_shortcut_delete(self, _):
        tab = self.nb.select()
        if tab == str(self.tab_clientes):
            self.remove_cliente()
        elif tab == str(self.tab_materia):
            self.remove_material()

    def on_shortcut_refresh(self, _):
        self.refresh()

    def _check_year_rollover(self):
        now_year = str(datetime.now().year)
        last_year = str(getattr(self, "_runtime_year_token", now_year))
        if now_year != last_year:
            self._runtime_year_token = now_year
            try:
                self.refresh_encomendas_year_options(keep_selection=False)
                self._on_encomendas_year_change(now_year)
            except Exception:
                pass
            try:
                self.refresh_orc_year_options(keep_selection=False)
                self._on_orc_year_change(now_year)
            except Exception:
                pass
            try:
                self.refresh_operador_year_options(keep_selection=False)
                self._on_operador_year_change(now_year)
            except Exception:
                pass
            try:
                self.refresh_ordens_year_options(keep_selection=False)
                self._on_ordens_year_change(now_year)
            except Exception:
                pass
        try:
            self.root.after(60000, self._check_year_rollover)
        except Exception:
            pass

    def sort_treeview(self, tree, col, reverse):
        return app_misc_actions.sort_treeview(self, tree, col, reverse)


    def refresh(self, full=False, persist=False):
        return app_misc_actions.refresh(self, full, persist)


    def pick_date(self, target_var, parent=None):
        return plan_actions.pick_date(self, target_var, parent)


    def build_clientes(self):
        return clientes_rooting.build_clientes(
            self,
            custom_tk_available=CUSTOM_TK_AVAILABLE,
            cond_pagamento_opcoes=COND_PAGAMENTO_OPCOES,
        )

    def on_select_cliente(self, _):
        return clientes_actions.on_select_cliente(self, _)


    def _set_cliente_form_values(self, values):
        return clientes_actions._set_cliente_form_values(self, values)


    def _on_select_cliente_custom(self, codigo):
        return clientes_actions._on_select_cliente_custom(self, codigo)


    def add_cliente(self):
        return clientes_actions.add_cliente(self)


    def edit_cliente(self):
        return clientes_actions.edit_cliente(self)


    def remove_cliente(self):
        return clientes_actions.remove_cliente(self)


    def refresh_clientes(self):
        return clientes_actions.refresh_clientes(self)


    def build_materia(self):
        return ui_build_blocks.build_materia(self)


    def on_select_stock_material(self, _):
        return materia_actions.on_select_stock_material(self, _)


    def add_material(self):
        return materia_actions.add_material(self)


    def edit_material(self):
        return materia_actions.edit_material(self)


    def corrigir_stock(self):
        return materia_actions.corrigir_stock(self)


    def remove_material(self):
        return materia_actions.remove_material(self)


    def baixa_material(self):
        return materia_actions.baixa_material(self)


    def refresh_materia(self):
        return materia_actions.refresh_materia(self)


    def show_stock_log(self):
        return materia_actions.show_stock_log(self)


    def render_stock_log_pdf(self, path):
        return materia_actions.render_stock_log_pdf(self, path)


    def build_encomendas(self):
        return ui_build_blocks.build_encomendas(self)


    def get_clientes_codes(self):
        return [c["codigo"] for c in self.data["clientes"]]

    def get_clientes_display(self):
        return clientes_actions.get_clientes_display(self)


    def add_encomenda(self):
        return encomendas_actions.add_encomenda(self)


    def edit_encomenda(self):
        return encomendas_actions.edit_encomenda(self)


    def remove_encomenda(self):
        return encomendas_actions.remove_encomenda(self)


    def edit_tempo_espessura(self):
        return encomendas_actions.edit_tempo_espessura(self)


    def on_cativar_toggle(self):
        if not self.tbl_encomendas.selection():
            if self.e_cativar.get():
                messagebox.showinfo("Info", "Crie ou selecione uma encomenda para cativar")
                self.e_cativar.set(False)
            return
        if self.e_cativar.get():
            opened = self.cativar_stock()
            if not opened:
                self.e_cativar.set(False)
        self.save_cativar_encomenda()

    def on_select_encomenda(self, _=None):
        return encomendas_actions.on_select_encomenda(self, _)


    def clear_encomenda_selection(self):
        return encomendas_actions.clear_encomenda_selection(self)


    def clear_material_selection(self):
        self.selected_material = None
        self.selected_espessura = None
        for t in [self.tbl_espessuras, self.tbl_pecas]:
            for i in t.get_children():
                t.delete(i)

    def clear_espessura_selection(self):
        return encomendas_actions.clear_espessura_selection(self)


    def clear_peca_selection(self):
        return encomendas_actions.clear_peca_selection(self)


    def _update_nota_cliente_visual(self):
        return encomendas_actions._update_nota_cliente_visual(self)


    def save_nota_cliente_encomenda(self, _e=None):
        return encomendas_actions.save_nota_cliente_encomenda(self, _e)


    def save_data_entrega_encomenda(self, _e=None):
        return encomendas_actions.save_data_entrega_encomenda(self, _e)

    def save_cativar_encomenda(self, _e=None):
        return encomendas_actions.save_cativar_encomenda(self, _e)



    def get_encomenda_by_numero(self, numero):
        return encomendas_actions.get_encomenda_by_numero(self, numero)


    def add_material_encomenda(self):
        return encomendas_actions.add_material_encomenda(self)


    def remove_material_encomenda(self):
        return encomendas_actions.remove_material_encomenda(self)


    def add_espessura(self, enc_override=None, material_override=None):
        return encomendas_actions.add_espessura(
            self,
            enc_override=enc_override,
            material_override=material_override,
        )


    def remove_espessura(self):
        return encomendas_actions.remove_espessura(self)


    def remove_peca(self):
        return encomendas_actions.remove_peca(self)


    def open_selected_peca_desenho(self):
        return encomendas_actions.open_selected_peca_desenho(self)

    def open_peca_desenho_by_refs(self, numero=None, ref_interna="", ref_externa="", enc=None, silent=False):
        return encomendas_actions.open_peca_desenho_by_refs(
            self,
            numero=numero,
            ref_interna=ref_interna,
            ref_externa=ref_externa,
            enc=enc,
            silent=silent,
        )

    def open_encomenda_by_numero(self, numero, open_editor=False):
        return encomendas_actions.open_encomenda_by_numero(self, numero, open_editor=open_editor)

    def open_encomenda_info_by_numero(self, numero, highlight_ref=None):
        return encomendas_actions.open_encomenda_info_by_numero(self, numero, highlight_ref=highlight_ref)


    def on_select_material(self, _):
        return encomendas_actions.on_select_material(self, _)


    def on_select_espessura(self, _):
        return encomendas_actions.on_select_espessura(self, _)


    def refresh_materiais(self, enc):
        return app_misc_actions.refresh_materiais(self, enc)


    def refresh_espessuras(self, enc, material=None):
        return encomendas_actions.refresh_espessuras(self, enc, material)


    def add_peca(self, esp_override=None):
        return encomendas_actions.add_peca(self, esp_override=esp_override)


    def open_ref_search(self, target_var, values, on_pick=None):
        return encomendas_actions.open_ref_search(self, target_var, values, on_pick)


    def registar_producao(self, enc=None, preselect_ref=None, default_ok=None):
        return encomendas_actions.registar_producao(self, enc, preselect_ref, default_ok)


    def reabrir_encomenda(self):
        return encomendas_actions.reabrir_encomenda(self)


    def libertar_reserva(self):
        return app_misc_actions.libertar_reserva(self)


    def cativar_stock(self, **kwargs):
        return encomendas_actions.cativar_stock(self, **kwargs)

    def cativar_stock_selecao(self):
        return encomendas_actions.cativar_stock_selecao(self)

    def descativar_stock_selecao(self, **kwargs):
        return encomendas_actions.descativar_stock_selecao(self, **kwargs)


    def get_selected_encomenda(self):
        return encomendas_actions.get_selected_encomenda(self)


    def refresh_encomendas(self):
        return encomendas_actions.refresh_encomendas(self)


    def refresh_encomendas_year_options(self, keep_selection=True):
        return encomendas_actions.refresh_encomendas_year_options(self, keep_selection)


    def _on_encomendas_year_change(self, value=None):
        return encomendas_actions._on_encomendas_year_change(self, value)


    def refresh_pecas(self, enc, espessura=None):
        return encomendas_actions.refresh_pecas(self, enc, espessura)


    def reselect_encomenda_material_espessura(self):
        return encomendas_actions.reselect_encomenda_material_espessura(self)


    def update_cativadas_display(self, enc):
        return encomendas_actions.update_cativadas_display(self, enc)


    def preview_encomenda(self):
        return encomendas_actions.preview_encomenda(self)


    def preview_stock_a4(self):
        return materia_actions.preview_stock_a4(self)


    def build_plano(self):
        return ui_build_blocks.build_plano(self)


    def sort_plano(self):
        self.data["encomendas"].sort(key=lambda e: e.get("data_entrega", ""))
        self.refresh_plano()

    def _sync_plano_hist(self):
        return plan_actions._sync_plano_hist(self)


    def refresh_plano(self):
        return plan_actions.refresh_plano(self)


    def get_plano_grid_metrics(self):
        return app_misc_actions.get_plano_grid_metrics(self)


    def get_plano_pausa_almoco(self):
        return (time_to_minutes("12:30"), time_to_minutes("14:00"))

    def plano_intervalo_bloqueado(self, start_min, end_min):
        p_ini, p_fim = self.get_plano_pausa_almoco()
        return not (end_min <= p_ini or start_min >= p_fim)

    def prev_week(self):
        return plan_actions.prev_week(self)


    def next_week(self):
        return plan_actions.next_week(self)


    def on_plano_drag_start(self, _):
        return app_misc_actions.on_plano_drag_start(self, _)


    def on_plano_click(self, event):
        return plan_actions.on_plano_click(self, event)


    def on_plano_drag_motion(self, event):
        return app_misc_actions.on_plano_drag_motion(self, event)


    def on_plano_drop(self, event):
        return plan_actions.on_plano_drop(self, event)


    def is_plano_duplicado(self, numero, material, espessura):
        for p in self.data.get("plano", []):
            if p.get("encomenda") == numero and p.get("material") == material and p.get("espessura") == espessura:
                return True
        return False

    def preview_plano_a4(self):
        return plan_actions.preview_plano_a4(self)


    def auto_planear(self):
        return plan_actions.auto_planear(self)


    def select_plano_order(self, items):
        return plan_actions.select_plano_order(self, items)


    def desplanear_tudo(self):
        return plan_actions.desplanear_tudo(self)


    def on_plano_block(self, item):
        return plan_actions.on_plano_block(self, item)


    def get_chapa_reservada(self, numero, material=None, espessura=None):
        return app_misc_actions.get_chapa_reservada(self, numero, material, espessura)


    def build_qualidade(self):
        return ui_build_blocks.build_qualidade(self)


    def build_orc(self):
        return ui_build_blocks.build_orc(self)


    def get_orc_by_numero(self, numero):
        return orc_actions.get_orc_by_numero(self, numero)


    def refresh_orc_list(self):
        return orc_actions.refresh_orc_list(self)


    def refresh_orc_year_options(self, keep_selection=True):
        return orc_actions.refresh_orc_year_options(self, keep_selection)


    def _on_orc_filter_click(self, event):
        return orc_actions._on_orc_filter_click(self, event)


    def _on_orc_year_change(self, value=None):
        return orc_actions._on_orc_year_change(self, value)


    def clear_orc_details(self):
        return orc_actions.clear_orc_details(self)


    def on_orc_select(self, _=None):
        return orc_actions.on_orc_select(self, _)


    def refresh_orc_linhas(self, orc):
        return orc_actions.refresh_orc_linhas(self, orc)


    def add_orcamento(self):
        return orc_actions.add_orcamento(self)


    def remove_orcamento(self):
        return orc_actions.remove_orcamento(self)


    def set_orc_estado(self, estado):
        return orc_actions.set_orc_estado(self, estado)


    def save_orc_fields(self, refresh_list=True):
        return orc_actions.save_orc_fields(self, refresh_list)


    def fill_orc_from_cliente(self):
        return orc_actions.fill_orc_from_cliente(self)


    def add_orc_linha(self):
        return orc_actions.add_orc_linha(self)


    def edit_orc_linha(self):
        return orc_actions.edit_orc_linha(self)


    def remove_orc_linha(self):
        return orc_actions.remove_orc_linha(self)


    def open_orc_linha_desenho(self):
        return orc_actions.open_orc_linha_desenho(self)


    def open_orc_linha(self, edit_index=None):
        return orc_actions.open_orc_linha(self, edit_index)


    def recalc_orc(self, orc):
        return orc_actions.recalc_orc(self, orc)


    def _orc_get_notes_text(self):
        return orc_actions._orc_get_notes_text(self)


    def _orc_set_notes_text(self, text):
        return orc_actions._orc_set_notes_text(self, text)


    def _orc_append_pdf_note(self, line):
        return orc_actions._orc_append_pdf_note(self, line)


    def _extract_orc_operacoes(self, orc):
        return orc_actions._extract_orc_operacoes(self, orc)


    def _build_orc_notes_lines(self, operacoes):
        return orc_actions._build_orc_notes_lines(self, operacoes)


    def orc_fill_notes_by_ops(self):
        return orc_actions.orc_fill_notes_by_ops(self)


    def convert_orc_to_encomenda(self):
        return orc_actions.convert_orc_to_encomenda(self)


    def prompt_orc_nota_cliente(self):
        return orc_actions.prompt_orc_nota_cliente(self)


    def render_orc_pdf(self, path, orc):
        return orc_actions.render_orc_pdf(self, path, orc)

    def preview_orcamento(self):
        return orc_actions.preview_orcamento(self)


    def save_orc_pdf(self):
        return orc_actions.save_orc_pdf(self)


    def print_orc_pdf(self):
        return orc_actions.print_orc_pdf(self)


    def open_orc_pdf_with(self):
        return orc_actions.open_orc_pdf_with(self)


    def build_operador(self):
        return ui_build_blocks.build_operador(self)


    def _op_select_enc_custom(self, numero):
        self.op_sel_enc_num = numero
        self._open_operador_encomenda_detail(numero)

    def _open_operador_encomenda_detail(self, numero):
        return operador_ordens_actions._open_operador_encomenda_detail(self, numero)


    def _op_esp_label(self, mat, esp):
        return f"{mat} | {fmt_num(esp)} mm"

    def _op_select_esp_combo(self):
        return app_misc_actions._op_select_esp_combo(self)


    def _on_op_status_click(self, value=None):
        if getattr(self, "_suppress_operador_filter_cb", False):
            return
        try:
            if value is not None:
                self.op_status_filter.set(value)
        except Exception:
            pass
        self.refresh_operador()

    def _op_select_peca_custom(self, p):
        if getattr(self, "op_use_full_custom", False):
            pid = str((p or {}).get("id", "") or "").strip()
            if not hasattr(self, "op_sel_pecas_ids") or not isinstance(getattr(self, "op_sel_pecas_ids", None), set):
                self.op_sel_pecas_ids = set()
            if pid:
                if pid in self.op_sel_pecas_ids:
                    self.op_sel_pecas_ids.discard(pid)
                else:
                    self.op_sel_pecas_ids.add(pid)
            self.op_sel_peca = p
        else:
            self.op_sel_peca = p
        try:
            enc = self.get_operador_encomenda(show_error=False)
            if enc and getattr(self, "op_use_full_custom", False):
                self.preencher_pecas_operador(enc, getattr(self, "op_material", None), getattr(self, "op_espessura", None))
            else:
                self.on_operador_select_espessura(None)
        except Exception:
            pass

    def _op_select_esp_custom(self, mat, esp):
        return app_misc_actions._op_select_esp_custom(self, mat, esp)


    def build_ordens(self):
        return ui_build_blocks.build_ordens(self)


    def _ordens_select_opp_custom(self, opp):
        return operador_ordens_actions._ordens_select_opp_custom(self, opp)


    def _ordens_get_selected_opp(self):
        return operador_ordens_actions._ordens_get_selected_opp(self)


    def refresh_operador(self):
        return operador_ordens_actions.refresh_operador(self)

    def refresh_operador_year_options(self, keep_selection=True):
        return operador_ordens_actions.refresh_operador_year_options(self, keep_selection)

    def _on_operador_year_change(self, value=None):
        return operador_ordens_actions._on_operador_year_change(self, value)


    def op_blink_schedule(self):
        return app_misc_actions.op_blink_schedule(self)


    def refresh_ordens(self):
        return operador_ordens_actions.refresh_ordens(self)

    def refresh_ordens_year_options(self, keep_selection=True):
        return operador_ordens_actions.refresh_ordens_year_options(self, keep_selection)

    def _on_ordens_year_change(self, value=None):
        return operador_ordens_actions._on_ordens_year_change(self, value)


    def preview_opp_selected(self):
        return app_misc_actions.preview_opp_selected(self)


    def calc_esp_totals(self, esp_obj):
        planeado = 0.0
        produzido = 0.0
        for p in esp_obj.get("pecas", []):
            planeado += float(p.get("quantidade_pedida", 0))
            produzido += float(p.get("produzido_ok", 0)) + float(p.get("produzido_nok", 0)) + float(p.get("produzido_qualidade", 0))
        return planeado, produzido

    def atualizar_estado_espessura(self, esp_obj):
        planeado, produzido = self.calc_esp_totals(esp_obj)
        pecas = esp_obj.get("pecas", [])
        if pecas and all(p.get("estado") == "Concluida" for p in pecas):
            esp_obj["estado"] = "Concluida"
        elif any("avari" in norm_text(p.get("estado", "")) for p in pecas):
            esp_obj["estado"] = "Avaria"
        elif any(p.get("estado") == "Em producao" for p in pecas) or (
            esp_obj.get("inicio_producao") and not esp_obj.get("fim_producao")
        ):
            esp_obj["estado"] = "Em producao"
        elif pecas:
            p_norm = [norm_text(p.get("estado", "")) for p in pecas]
            has_started = any(("produ" in s) or ("concl" in s) or ("incomplet" in s) or ("interromp" in s) or ("paus" in s) for s in p_norm) or bool(esp_obj.get("inicio_producao"))
            has_running = any(("produ" in s) and ("paus" not in s) and ("interromp" not in s) for s in p_norm)
            all_concl = p_norm and all("concl" in s for s in p_norm)
            all_prepar = p_norm and all("prepar" in s for s in p_norm)
            if has_started and (not has_running) and (not all_concl) and (not all_prepar):
                esp_obj["estado"] = "Em pausa"
            else:
                esp_obj["estado"] = "Preparacao"
        else:
            esp_obj["estado"] = "Preparacao"
        return planeado, produzido

    def on_operador_select(self, _):
        return operador_ordens_actions.on_operador_select(self, _)


    def clear_operador_selection(self):
        return operador_ordens_actions.clear_operador_selection(self)


    def clear_operador_esp(self):
        return operador_ordens_actions.clear_operador_esp(self)


    def on_operador_select_espessura(self, _):
        return operador_ordens_actions.on_operador_select_espessura(self, _)


    def preencher_pecas_operador(self, enc, mat, esp):
        return operador_ordens_actions.preencher_pecas_operador(self, enc, mat, esp)


    def operador_ajustar_quantidade(self, p, delta):
        return operador_ordens_actions.operador_ajustar_quantidade(self, p, delta)


    def get_operador_esp_obj(self, enc, mat, esp):
        return operador_ordens_actions.get_operador_esp_obj(self, enc, mat, esp)


    def finalizar_espessura(self, enc, mat, esp, auto=False):
        return operador_ordens_actions.finalizar_espessura(self, enc, mat, esp, auto)


    def confirmar_baixa_stock(self, mat, esp, total):
        return app_misc_actions.confirmar_baixa_stock(self, mat, esp, total)


    def operador_concluir_peca(self):
        return operador_ordens_actions.operador_concluir_peca(self)


    def operador_inserir_qtd(self):
        return operador_ordens_actions.operador_inserir_qtd(self)


    def operador_subtrair_qtd(self):
        return operador_ordens_actions.operador_subtrair_qtd(self)


    def operador_reabrir_peca(self):
        return operador_ordens_actions.operador_reabrir_peca(self)


    def gerir_operadores(self):
        return operador_ordens_actions.gerir_operadores(self)


    def gerir_orcamentistas(self):
        return orc_actions.gerir_orcamentistas(self)

    def refresh_runtime_impulse_data(self, cleanup_orphans=True):
        return mysql_refresh_runtime_impulse_data(self.data, cleanup_orphans=cleanup_orphans)


    def preview_operador_pdf(self):
        return operador_ordens_actions.preview_operador_pdf(self)


    def render_operador_producao_pdf(self, path, enc, op_name=""):
        return operador_ordens_actions.render_operador_producao_pdf(self, path, enc, op_name)


    def _preview_piece_pdf(self, enc, p):
        return operador_ordens_actions._preview_piece_pdf(self, enc, p)


    def preview_peca_pdf(self):
        enc = self.get_operador_encomenda()
        if not enc:
            return
        p = self.get_operador_peca(enc)
        if not p:
            return
        self._preview_piece_pdf(enc, p)

    def print_of_selected_encomenda(self):
        return operador_ordens_actions.print_of_selected_encomenda(self)


    def render_of_pdf(self, path, enc):
        return operador_ordens_actions.render_of_pdf(self, path, enc)


    def get_operador_encomenda(self, show_error=True):
        return operador_ordens_actions.get_operador_encomenda(self, show_error)


    def restore_operador_selection(self):
        return operador_ordens_actions.restore_operador_selection(self)


    def get_operador_peca(self, enc):
        return operador_ordens_actions.get_operador_peca(self, enc)


    def operador_inicio_peca(self):
        return operador_ordens_actions.operador_inicio_peca(self)


    def operador_inicio_todas_pecas_espessura(self):
        return operador_ordens_actions.operador_inicio_todas_pecas_espessura(self)


    def operador_retomar_peca(self):
        return operador_ordens_actions.operador_retomar_peca(self)


    def operador_fim_peca(self):
        return operador_ordens_actions.operador_fim_peca(self)


    def operador_fim_todas_pecas_espessura(self):
        return operador_ordens_actions.operador_fim_todas_pecas_espessura(self)


    def operador_interromper_peca(self):
        return operador_ordens_actions.operador_interromper_peca(self)


    def operador_registar_avaria(self):
        return operador_ordens_actions.operador_registar_avaria(self)

    def operador_fechar_avaria(self):
        return operador_ordens_actions.operador_fechar_avaria(self)


    def operador_reabrir_peca_total(self):
        return operador_ordens_actions.operador_reabrir_peca_total(self)


    def show_rejeitadas_hist(self):
        return operador_ordens_actions.show_rejeitadas_hist(self)


    def _get_operador_mat_esp(self):
        return operador_ordens_actions._get_operador_mat_esp(self)


    def operador_iniciar_espessura(self):
        return operador_ordens_actions.operador_iniciar_espessura(self)


    def operador_finalizar_espessura(self):
        return operador_ordens_actions.operador_finalizar_espessura(self)


    def operador_iniciar_encomenda(self):
        return operador_ordens_actions.operador_iniciar_encomenda(self)


    def operador_interromper(self):
        return operador_ordens_actions.operador_interromper(self)


    def operador_finalizar_encomenda(self):
        return operador_ordens_actions.operador_finalizar_encomenda(self)


    def refresh_qualidade(self):
        return qualidade_actions.refresh_qualidade(self)


    def build_export(self):
        return ui_build_blocks.build_export(self)


    def build_expedicao(self):
        return ui_build_blocks.build_expedicao(self)


    def clear_expedicao_selection(self):
        return ne_expedicao_actions.clear_expedicao_selection(self)


    def get_exp_selected_encomenda(self):
        if hasattr(self, "exp_tbl_enc"):
            sel = self.exp_tbl_enc.selection()
            if sel:
                num = str(self.exp_tbl_enc.item(sel[0], "values")[0]).strip()
                if num:
                    self.exp_sel_enc_num = num
        if not self.exp_sel_enc_num:
            return None
        return self.get_encomenda_by_numero(self.exp_sel_enc_num)

    def on_exp_select_encomenda(self, _=None):
        enc = self.get_exp_selected_encomenda()
        if not enc:
            return
        self.exp_draft_linhas = []
        self.refresh_expedicao_pecas()
        self.refresh_expedicao_draft()

    def refresh_expedicao(self):
        return ne_expedicao_actions.refresh_expedicao(self)


    def refresh_expedicao_pending(self):
        return ne_expedicao_actions.refresh_expedicao_pending(self)


    def refresh_expedicao_pecas(self):
        return ne_expedicao_actions.refresh_expedicao_pecas(self)


    def refresh_expedicao_draft(self):
        return ne_expedicao_actions.refresh_expedicao_draft(self)


    def refresh_expedicao_hist(self):
        return ne_expedicao_actions.refresh_expedicao_hist(self)


    def get_selected_expedicao(self, show_error=True):
        return ne_expedicao_actions.get_selected_expedicao(self, show_error)


    def on_exp_hist_click_open_pdf(self, event=None):
        return app_misc_actions.on_exp_hist_click_open_pdf(self, event)


    def _exp_parse_datetime(self, value, default_iso=""):
        return app_misc_actions._exp_parse_datetime(self, value, default_iso)


    def _exp_datetime_ui(self, iso_txt):
        raw = str(iso_txt or "").strip()
        if not raw:
            return datetime.now().strftime("%Y-%m-%d %H:%M")
        try:
            return datetime.fromisoformat(raw.replace("Z", "+00:00")).strftime("%Y-%m-%d %H:%M")
        except Exception:
            if "T" in raw:
                return raw.replace("T", " ")[:16]
            return raw[:16]

    def prompt_dados_guia(self, title, initial, linhas):
        return ne_expedicao_actions.prompt_dados_guia(self, title, initial, linhas)


    def render_expedicao_pdf(self, path, ex):
        return ne_expedicao_actions.render_expedicao_pdf(self, path, ex)


    def preview_expedicao_pdf_by_num(self, numero):
        return ne_expedicao_actions.preview_expedicao_pdf_by_num(self, numero)


    def preview_expedicao_pdf(self):
        return ne_expedicao_actions.preview_expedicao_pdf(self)


    def save_expedicao_pdf(self):
        return ne_expedicao_actions.save_expedicao_pdf(self)


    def print_expedicao_pdf(self):
        return ne_expedicao_actions.print_expedicao_pdf(self)


    def editar_expedicao(self):
        return ne_expedicao_actions.editar_expedicao(self)


    def add_exp_linha(self):
        return ne_expedicao_actions.add_exp_linha(self)


    def remove_exp_linha(self):
        sel = self.exp_tbl_linhas.selection() if hasattr(self, "exp_tbl_linhas") else ()
        if not sel:
            return
        idx = self.exp_tbl_linhas.index(sel[0])
        if idx < 0 or idx >= len(self.exp_draft_linhas):
            return
        self.exp_draft_linhas.pop(idx)
        self.refresh_expedicao_pecas()
        self.refresh_expedicao_draft()

    def emitir_expedicao_off(self):
        return ne_expedicao_actions.emitir_expedicao_off(self)


    def anular_expedicao(self):
        return ne_expedicao_actions.anular_expedicao(self)


    def criar_guia_manual(self):
        return ne_expedicao_actions.criar_guia_manual(self)


    def edit_empresa_info(self):
        return app_misc_actions.edit_empresa_info(self)

    def manage_user_accounts(self):
        return app_misc_actions.manage_user_accounts(self)


    def choose_primary_color(self):
        return app_misc_actions.choose_primary_color(self)


    def reset_primary_color(self):
        return app_misc_actions.reset_primary_color(self)


    def apply_primary_color_runtime(self, old_primary="", old_hover=""):
        return app_misc_actions.apply_primary_color_runtime(self, old_primary, old_hover)


    def refresh_module_contexts(self):
        main_globals = globals()
        ui_build_blocks.configure(main_globals)
        clientes_actions.configure(main_globals)
        materia_actions.configure(main_globals)
        encomendas_actions.configure(main_globals)
        qualidade_actions.configure(main_globals)
        orc_actions.configure(main_globals)
        operador_ordens_actions.configure(main_globals)
        ne_expedicao_actions.configure(main_globals)
        plan_actions.configure(main_globals)
        produtos_actions.configure(main_globals)
        app_misc_actions.configure(main_globals)


    def export_csv(self, tipo):
        return app_misc_actions.export_csv(self, tipo)

    def open_sheet_calculator(self):
        return app_misc_actions.open_sheet_calculator(self)


    # ------------------ PRODUTOS ------------------
    def build_produtos(self):
        return ui_build_blocks.build_produtos(self)


    def produto_dict_from_form(self):
        return produtos_actions.produto_dict_from_form(self)


    def novo_produto(self):
        return produtos_actions.novo_produto(self)


    def product_apply_category_mode(self):
        return app_misc_actions.product_apply_category_mode(self)


    def editar_produto_dialog(self):
        return produtos_actions.editar_produto_dialog(self)


    def dialog_novo_produto(self, prod_init=None):
        return produtos_actions.dialog_novo_produto(self, prod_init)


    def refresh_produtos(self):
        return produtos_actions.refresh_produtos(self)


    def _set_prod_form_from_obj(self, p):
        return app_misc_actions._set_prod_form_from_obj(self, p)


    def _select_produto_custom(self, codigo):
        return produtos_actions._select_produto_custom(self, codigo)


    def on_produto_select(self, _event=None):
        return produtos_actions.on_produto_select(self, _event)


    def _get_selected_prod_codigo(self):
        if self.prod_use_full_custom:
            if self.prod_sel_codigo:
                return self.prod_sel_codigo
            c = self.prod_codigo.get().strip()
            return c if c else ""
        sel = self.tbl_produtos.selection()
        if not sel:
            return ""
        return self.tbl_produtos.item(sel[0], "values")[0]

    def guardar_produto(self):
        return produtos_actions.guardar_produto(self)


    def remover_produto(self):
        return produtos_actions.remover_produto(self)


    def render_produtos_stock_pdf(self, path):
        return produtos_actions.render_produtos_stock_pdf(self, path)


    def preview_produtos_stock_pdf(self):
        return produtos_actions.preview_produtos_stock_pdf(self)


    def save_produtos_stock_pdf(self):
        return produtos_actions.save_produtos_stock_pdf(self)


    def print_produtos_stock_pdf(self):
        return produtos_actions.print_produtos_stock_pdf(self)


    def get_selected_produto_obj(self):
        return produtos_actions.get_selected_produto_obj(self)


    def produto_dar_baixa_dialog(self):
        return produtos_actions.produto_dar_baixa_dialog(self)


    def _get_produtos_mov_operador(self, operador=""):
        return operador_ordens_actions._get_produtos_mov_operador(self, operador)


    def render_produtos_mov_operador_pdf(self, path, operador="Todos"):
        return operador_ordens_actions.render_produtos_mov_operador_pdf(self, path, operador)


    def produto_mov_operador_dialog(self, default_operador=""):
        return operador_ordens_actions.produto_mov_operador_dialog(self, default_operador)


    # ------------------ NOTA ENCOMENDA ------------------
    def build_ne(self):
        return ui_build_blocks.build_ne(self)


    def nova_ne(self, create_draft=True):
        return ne_expedicao_actions.nova_ne(self, create_draft)


    def refresh_ne_fornecedores(self):
        return ne_expedicao_actions.refresh_ne_fornecedores(self)


    def on_ne_fornecedor_change(self, _e=None):
        return ne_expedicao_actions.on_ne_fornecedor_change(self, _e)


    def manage_fornecedores(self):
        return ne_expedicao_actions.manage_fornecedores(self)


    def ne_collect_lines(self):
        return ne_expedicao_actions.ne_collect_lines(self)


    def _update_produto_preco_from_unit(self, produto_codigo, preco_unit):
        return produtos_actions._update_produto_preco_from_unit(self, produto_codigo, preco_unit)


    def _update_materia_preco_from_unit(self, materia_id, preco_unit):
        return ne_expedicao_actions._update_materia_preco_from_unit(self, materia_id, preco_unit)


    def _sync_ne_linhas_with_produtos(self, ne):
        return produtos_actions._sync_ne_linhas_with_produtos(self, ne)


    def _sync_ne_linhas_with_materia(self, ne):
        return ne_expedicao_actions._sync_ne_linhas_with_materia(self, ne)


    def sync_all_ne_from_materia(self):
        return app_misc_actions.sync_all_ne_from_materia(self)


    def _ne_line_parse_values(self, vals):
        return ne_expedicao_actions._ne_line_parse_values(self, vals)


    def refresh_ne_total(self):
        return ne_expedicao_actions.refresh_ne_total(self)


    def _on_ne_estado_filter_click(self, value=None):
        return ne_expedicao_actions._on_ne_estado_filter_click(self, value)


    def ne_refresh_line_tags(self):
        return ne_expedicao_actions.ne_refresh_line_tags(self)


    def ne_add_linha(self):
        return ne_expedicao_actions.ne_add_linha(self)


    def _dialog_ne_origem_linha(self):
        return ne_expedicao_actions._dialog_ne_origem_linha(self)


    def _dialog_escolher_materia_prima_ne(self):
        return ne_expedicao_actions._dialog_escolher_materia_prima_ne(self)

    def _dialog_escolher_materia_para_produto(self):
        return produtos_actions._dialog_escolher_materia_para_produto(self)


    def on_ne_lin_select(self, _e=None):
        pass

    def ne_edit_linha(self, _e=None):
        return ne_expedicao_actions.ne_edit_linha(self, _e)


    def ne_del_linha(self):
        return ne_expedicao_actions.ne_del_linha(self)


    def guardar_ne(self):
        return ne_expedicao_actions.guardar_ne(self)


    def remover_ne(self):
        return ne_expedicao_actions.remover_ne(self)


    def refresh_ne(self):
        return ne_expedicao_actions.refresh_ne(self)


    def on_ne_select(self, _e=None):
        return ne_expedicao_actions.on_ne_select(self, _e)


    def get_ne_selected(self):
        return ne_expedicao_actions.get_ne_selected(self)


    def aprovar_ne(self):
        return ne_expedicao_actions.aprovar_ne(self)


    def gerar_nes_por_fornecedor(self):
        return ne_expedicao_actions.gerar_nes_por_fornecedor(self)


    def _next_materia_id(self):
        return ne_expedicao_actions._next_materia_id(self)


    def _receber_linha_materia_ne(self, l, ne_num, qtd_recebida=None, **kwargs):
        return ne_expedicao_actions._receber_linha_materia_ne(self, l, ne_num, qtd_recebida, **kwargs)


    def _dialog_associar_fatura_ne(self, ne):
        return ne_expedicao_actions._dialog_associar_fatura_ne(self, ne)


    def associar_fatura_ne(self):
        return ne_expedicao_actions.associar_fatura_ne(self)


    def show_ne_documentos(self):
        return ne_expedicao_actions.show_ne_documentos(self)


    def _dialog_confirmar_entrega_ne(self, ne):
        return ne_expedicao_actions._dialog_confirmar_entrega_ne(self, ne)


    def confirmar_entrega_ne(self):
        return ne_expedicao_actions.confirmar_entrega_ne(self)


    def ne_blink_schedule(self):
        return ne_expedicao_actions.ne_blink_schedule(self)


    def render_ne_pdf(self, path, ne):
        return ne_expedicao_actions.render_ne_pdf(self, path, ne)


    def preview_ne(self):
        return ne_expedicao_actions.preview_ne(self)


    def save_ne_pdf(self):
        return ne_expedicao_actions.save_ne_pdf(self)


    def open_ne_pdf_with(self):
        return ne_expedicao_actions.open_ne_pdf_with(self)


    def _open_pdf_default(self, path):
        return ne_expedicao_actions._open_pdf_default(self, path)


    def render_ne_cotacao_pdf(self, path, ne):
        return ne_expedicao_actions.render_ne_cotacao_pdf(self, path, ne)


    def preview_ne_cotacao(self):
        return ne_expedicao_actions.preview_ne_cotacao(self)


    def save_ne_cotacao_pdf(self):
        return ne_expedicao_actions.save_ne_cotacao_pdf(self)


    def open_ne_cotacao_pdf_with(self):
        return ne_expedicao_actions.open_ne_cotacao_pdf_with(self)


    # ------------------ UTILIDADES PRODUTOS/NE ------------------
    def _dialog_escolher_produto(self, for_ne=False):
        return produtos_actions._dialog_escolher_produto(self, for_ne)



class SimpleInput(Toplevel):
    def __init__(self, root, title, prompt):
        super().__init__(root)
        self.title(title)
        self.value = None
        ttk.Label(self, text=prompt).pack(padx=10, pady=10)
        self.var = StringVar()
        ttk.Entry(self, textvariable=self.var).pack(padx=10)
        ttk.Button(self, text="OK", command=self.on_ok).pack(pady=10)
        self.grab_set()

    def on_ok(self):
        self.value = self.var.get().strip()
        self.destroy()


def simple_input(root, title, prompt):
    win = SimpleInput(root, title, prompt)
    root.wait_window(win)
    return win.value


def _resolve_login_video_path():
    fname = "grok-video-a10bdedc-d7d4-499e-b10a-6f43bf5cd2b2.mp4"
    fname2 = "grok-video-a10bdedc-d7d4-499e-b10a-6f43bf5cd2b2 (2).mp4"
    fname4 = "grok-video-a10bdedc-d7d4-499e-b10a-6f43bf5cd2b2 (4).mp4"
    fname_new = "grok-video-d1b43a87-e57a-48d4-a765-c668cac4ea16 (1).mp4"
    fname_new2 = "grok-video-d1b43a87-e57a-48d4-a765-c668cac4ea16 (2).mp4"
    env_path = str(os.environ.get("LUGEST_LOGIN_VIDEO", "") or "").strip()
    if env_path and os.path.exists(env_path):
        return env_path
    candidates = [
        os.path.join(os.path.expanduser("~"), "Desktop", fname_new2),
        os.path.join(os.path.expanduser("~"), "Desktop", fname_new),
        os.path.join(os.path.expanduser("~"), "Desktop", fname4),
        os.path.join(os.path.expanduser("~"), "Desktop", fname2),
        os.path.join(BASE_DIR, "Logos", fname_new2),
        os.path.join(BASE_DIR, "Logos", fname_new),
        os.path.join(BASE_DIR, "Logos", fname4),
        os.path.join(BASE_DIR, "Logos", fname),
        os.path.join(BASE_DIR, "Logos", fname2),
        os.path.join(BASE_DIR, "app", "Logos", fname_new2),
        os.path.join(BASE_DIR, "app", "Logos", fname_new),
        os.path.join(BASE_DIR, "app", "Logos", fname4),
        os.path.join(BASE_DIR, "app", "Logos", fname),
        os.path.join(BASE_DIR, "app", "Logos", fname2),
        os.path.join(os.path.expanduser("~"), "Desktop", "LUGEST_Instalacao_NovoPC V2", "app", "Logos", fname_new2),
        os.path.join(os.path.expanduser("~"), "Desktop", "LUGEST_Instalacao_NovoPC V2", "app", "Logos", fname_new),
        os.path.join(os.path.expanduser("~"), "Desktop", "LUGEST_Instalacao_NovoPC V2", "app", "Logos", fname4),
        os.path.join(os.path.expanduser("~"), "Desktop", "LUGEST_Instalacao_NovoPC V2", "app", "Logos", fname),
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return ""


def _play_login_video_splash(root):
    if str(os.environ.get("LUGEST_LOGIN_VIDEO_ENABLED", "1") or "1").strip().lower() in ("0", "false", "no", "off"):
        return
    video_path = _resolve_login_video_path()
    if not video_path:
        return
    try:
        import cv2  # type: ignore
        from PIL import Image, ImageTk  # type: ignore
    except Exception:
        return
    cap = None
    win = None
    try:
        cap = cv2.VideoCapture(video_path)
        if not cap or not cap.isOpened():
            return
        fps = float(cap.get(cv2.CAP_PROP_FPS) or 0.0)
        if fps <= 1.0:
            fps = 24.0
        delay = max(8, int(1000.0 / fps))
        src_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 1280)
        src_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 720)
        sw = int(root.winfo_screenwidth() or 1280)
        sh = int(root.winfo_screenheight() or 720)
        try:
            dpi = float(root.winfo_fpixels("1i") or 96.0)
        except Exception:
            dpi = 96.0
        # Tamanho maximo do splash (default: 60mm, pode ajustar por env)
        mm_limit = float(os.environ.get("LUGEST_LOGIN_VIDEO_MM", "60") or "60")
        mm_limit = max(35.0, min(90.0, mm_limit))
        max_logo_px = int((mm_limit / 25.4) * dpi)
        max_logo_px = max(80, max_logo_px)
        max_w = min(max_logo_px, int(sw * 0.35))
        max_h = min(max_logo_px, int(sh * 0.35))
        dst_w = max_w
        dst_h = max_h

        stop_flag = {"stop": False}
        win = Toplevel(root)
        win.overrideredirect(True)
        win.configure(bg="black")
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        lbl = Label(win, bg="black", bd=0, highlightthickness=0)
        lbl.pack(fill="both", expand=True)
        win.bind("<Button-1>", lambda _e=None: stop_flag.__setitem__("stop", True))
        win.bind("<Escape>", lambda _e=None: stop_flag.__setitem__("stop", True))
        win.bind("<space>", lambda _e=None: stop_flag.__setitem__("stop", True))

        # Sem recorte: usa frame completo com as bordas originais do video.
        ok0, frame0 = cap.read()
        if ok0 and frame0 is not None:
            scale = min(max_w / max(1, src_w), max_h / max(1, src_h))
            dst_w = max(64, int(src_w * scale))
            dst_h = max(64, int(src_h * scale))
            x = max(0, (sw - dst_w) // 2)
            y = max(0, (sh - dst_h) // 2)
            win.geometry(f"{dst_w}x{dst_h}+{x}+{y}")
            cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
        else:
            scale = min(max_w / max(1, src_w), max_h / max(1, src_h))
            dst_w = max(64, int(src_w * scale))
            dst_h = max(64, int(src_h * scale))
            x = max(0, (sw - dst_w) // 2)
            y = max(0, (sh - dst_h) // 2)
            win.geometry(f"{dst_w}x{dst_h}+{x}+{y}")

        max_sec = float(os.environ.get("LUGEST_LOGIN_VIDEO_MAX_SEC", "45") or "45")
        t0 = time.time()
        while True:
            if stop_flag["stop"]:
                break
            if (time.time() - t0) > max_sec:
                break
            frame_t0 = time.time()
            ok, frame = cap.read()
            if not ok:
                break
            if frame is None:
                continue
            frame = cv2.resize(frame, (dst_w, dst_h), interpolation=cv2.INTER_AREA)
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame)
            tk_img = ImageTk.PhotoImage(img)
            lbl.configure(image=tk_img)
            lbl.image = tk_img
            try:
                win.update_idletasks()
                win.update()
            except Exception:
                break
            elapsed_ms = int((time.time() - frame_t0) * 1000.0)
            sleep_ms = max(1, delay - elapsed_ms)
            time.sleep(sleep_ms / 1000.0)
    except Exception:
        pass
    finally:
        try:
            if cap is not None:
                cap.release()
        except Exception:
            pass
        try:
            if win is not None and win.winfo_exists():
                win.destroy()
        except Exception:
            pass


def _resolve_login_image_path():
    env_path = str(os.environ.get("LUGEST_LOGIN_IMAGE", "") or "").strip()
    if env_path and os.path.exists(env_path):
        return env_path
    names = [
        "grok-video-1d91139b-3f64-4c98-bfdb-5901c9fc5177.mp4",
        "grok-video-ac2bf3af-8bcf-49e6-ae21-385edfb44115.mp4",
        "logo-removebg-preview.png",
        "lugest_splash.png",
        "lugest_splash.jpg",
        "lugest_splash.jpeg",
        "lugest_login.png",
        "lugest_login.jpg",
    ]
    candidates = []
    for n in names:
        candidates.extend(
            [
                os.path.join(BASE_DIR, "Logos", n),
                os.path.join(BASE_DIR, "app", "Logos", n),
                os.path.join(os.path.expanduser("~"), "Desktop", n),
                os.path.join(os.path.expanduser("~"), "Desktop", "LUGEST_Instalacao_NovoPC V2", "app", "Logos", n),
            ]
        )
    for p in candidates:
        if p and os.path.exists(p):
            return p
    # fallback: logo principal atual
    return get_orc_logo_path() or ""


def _show_login_image_splash(root):
    if str(os.environ.get("LUGEST_LOGIN_IMAGE_ENABLED", "1") or "1").strip().lower() in ("0", "false", "no", "off"):
        return
    img_path = _resolve_login_image_path()
    if not img_path or (not os.path.exists(img_path)):
        return
    try:
        from PIL import Image, ImageTk  # type: ignore
    except Exception:
        return
    win = None
    try:
        try:
            dpi = float(root.winfo_fpixels("1i") or 96.0)
        except Exception:
            dpi = 96.0
        mm_limit = float(os.environ.get("LUGEST_LOGIN_IMAGE_MM", "70") or "70")
        mm_limit = max(35.0, min(120.0, mm_limit))
        max_px = int((mm_limit / 25.4) * dpi)
        max_px = max(120, max_px)

        ext = os.path.splitext(img_path)[1].strip().lower()
        is_video = ext in (".mp4", ".avi", ".mov", ".mkv", ".webm")

        splash_bg = str(os.environ.get("LUGEST_LOGIN_SPLASH_BG", "#f3f6fb") or "#f3f6fb").strip()
        if not splash_bg.startswith("#") or len(splash_bg) not in (4, 7):
            splash_bg = "#f3f6fb"
        if len(splash_bg) == 4:
            splash_bg = "#" + "".join(ch * 2 for ch in splash_bg[1:])
        try:
            bg_rgb = tuple(int(splash_bg[i:i + 2], 16) for i in (1, 3, 5))
        except Exception:
            splash_bg = "#f3f6fb"
            bg_rgb = (243, 246, 251)

        sw = int(root.winfo_screenwidth() or 1280)
        sh = int(root.winfo_screenheight() or 720)

        stop_flag = {"stop": False}
        win = Toplevel(root)
        win.overrideredirect(True)
        win.configure(bg=splash_bg)
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass

        if is_video:
            try:
                import cv2  # type: ignore
                import numpy as np  # type: ignore
            except Exception:
                return
            cap = cv2.VideoCapture(img_path)
            if not cap or not cap.isOpened():
                return
            fps = float(cap.get(cv2.CAP_PROP_FPS) or 0.0)
            if fps <= 1.0:
                fps = 24.0
            delay = max(8, int(1000.0 / fps))
            vw = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
            vh = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
            if vw <= 0 or vh <= 0:
                ok, frame0 = cap.read()
                if not ok or frame0 is None:
                    cap.release()
                    return
                vh, vw = frame0.shape[:2]
                cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            scale = min(max_px / float(vw), max_px / float(vh))
            nw = max(90, int(vw * scale))
            nh = max(90, int(vh * scale))
            x = max(0, (sw - nw) // 2)
            y = max(0, (sh - nh) // 2)
            win.geometry(f"{nw}x{nh}+{x}+{y}")
            lbl = Label(win, bg=splash_bg, bd=0, highlightthickness=0)
            lbl.pack(fill="both", expand=True)
            win.bind("<Button-1>", lambda _e=None: stop_flag.__setitem__("stop", True))
            win.bind("<Escape>", lambda _e=None: stop_flag.__setitem__("stop", True))
            win.bind("<space>", lambda _e=None: stop_flag.__setitem__("stop", True))
            dur = float(os.environ.get("LUGEST_LOGIN_IMAGE_SEC", "2.8") or "2.8")
            dur = max(0.6, min(12.0, dur))
            t0 = time.time()
            bg_arr = np.array(bg_rgb, dtype=np.float32)
            while True:
                if stop_flag["stop"]:
                    break
                if (time.time() - t0) >= dur:
                    break
                frame_t0 = time.time()
                ok, frame = cap.read()
                if not ok or frame is None:
                    break
                frame = cv2.resize(frame, (nw, nh), interpolation=cv2.INTER_AREA)
                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB).astype(np.float32)
                # Atenua fundo branco para aproximar ao fundo da app.
                whiteness = np.min(rgb, axis=2)
                a = np.clip((whiteness - 210.0) / 45.0, 0.0, 1.0)
                rgb = rgb * (1.0 - a[..., None]) + bg_arr * a[..., None]
                rgb = np.clip(rgb, 0, 255).astype("uint8")
                img = Image.fromarray(rgb)
                tk_img = ImageTk.PhotoImage(img)
                lbl.configure(image=tk_img)
                lbl.image = tk_img
                try:
                    win.update_idletasks()
                    win.update()
                except Exception:
                    break
                elapsed_ms = int((time.time() - frame_t0) * 1000.0)
                sleep_ms = max(1, delay - elapsed_ms)
                time.sleep(sleep_ms / 1000.0)
            cap.release()
            return

        img = Image.open(img_path).convert("RGBA")
        iw, ih = img.size
        if iw <= 0 or ih <= 0:
            return
        scale = min(max_px / float(iw), max_px / float(ih))
        nw = max(90, int(iw * scale))
        nh = max(90, int(ih * scale))
        img = img.resize((nw, nh), Image.LANCZOS)
        tk_img = ImageTk.PhotoImage(img)
        x = max(0, (sw - nw) // 2)
        y = max(0, (sh - nh) // 2)
        win.geometry(f"{nw}x{nh}+{x}+{y}")
        lbl = Label(win, image=tk_img, bg=splash_bg, bd=0, highlightthickness=0)
        lbl.image = tk_img
        lbl.pack(fill="both", expand=True)
        win.bind("<Button-1>", lambda _e=None: stop_flag.__setitem__("stop", True))
        win.bind("<Escape>", lambda _e=None: stop_flag.__setitem__("stop", True))
        win.bind("<space>", lambda _e=None: stop_flag.__setitem__("stop", True))

        dur = float(os.environ.get("LUGEST_LOGIN_IMAGE_SEC", "2.2") or "2.2")
        dur = max(0.6, min(12.0, dur))
        t0 = time.time()
        while True:
            if stop_flag["stop"]:
                break
            if (time.time() - t0) >= dur:
                break
            try:
                win.update_idletasks()
                win.update()
            except Exception:
                break
            time.sleep(0.02)
    except Exception:
        pass
    finally:
        try:
            if win is not None and win.winfo_exists():
                win.destroy()
        except Exception:
            pass


def main():
    root = Tk()
    _configure_font_compatibility(root)
    apply_window_icon(root)
    root.withdraw()
    try:
        login = LoginWindow(root)
        root.wait_window(login.win)
        if login.user is None:
            return
        _show_login_image_splash(root)
        app = App(root, login.user, prefetched_data=getattr(login, "loaded_data", None))
    except Exception as exc:
        try:
            messagebox.showerror("Erro de arranque", str(exc))
        except Exception:
            pass
        return
    root.deiconify()
    # Deixa a UI desenhar primeiro; o refresh inicial corre logo a seguir.
    root.after(20, app.refresh)

    def _save_heartbeat():
        try:
            flush_pending_save()
        except Exception:
            pass
        try:
            err = _consume_async_save_error()
            if err:
                messagebox.showwarning("Aviso MySQL", f"A última gravação apresentou erro:\n{err}")
        except Exception:
            pass
        try:
            root.after(_SAVE_HEARTBEAT_MS, _save_heartbeat)
        except Exception:
            pass

    def _on_root_close():
        try:
            flush_pending_save(force=True)
        except Exception:
            pass
        try:
            _drain_async_saves(timeout_sec=20.0)
        except Exception:
            pass
        try:
            _ASYNC_SAVE_STOP.set()
            _ASYNC_SAVE_EVENT.set()
        except Exception:
            pass
        try:
            if _ASYNC_SAVE_THREAD is not None and _ASYNC_SAVE_THREAD.is_alive():
                _ASYNC_SAVE_THREAD.join(timeout=1.0)
        except Exception:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_root_close)
    root.after(_SAVE_HEARTBEAT_MS, _save_heartbeat)
    root.mainloop()


def _local_project_python_path() -> str:
    if getattr(sys, "frozen", False):
        return ""
    try:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), ".venv", "Scripts", "python.exe")
    except Exception:
        return ""


def _python_supports_module(python_path: str, module_name: str) -> bool:
    path = str(python_path or "").strip()
    name = str(module_name or "").strip()
    if not path or not name or not os.path.exists(path):
        return False
    try:
        completed = subprocess.run(
            [path, "-c", f"import {name}"],
            cwd=os.path.dirname(os.path.abspath(__file__)),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=15,
            check=False,
        )
        return int(completed.returncode or 0) == 0
    except Exception:
        return False


def _delegate_qt_launch_to_local_venv(argv=None):
    preferred_python = _local_project_python_path()
    if not preferred_python:
        return None
    try:
        current_python = os.path.normcase(os.path.abspath(sys.executable))
        preferred_norm = os.path.normcase(os.path.abspath(preferred_python))
    except Exception:
        current_python = str(sys.executable or "").strip().lower()
        preferred_norm = str(preferred_python or "").strip().lower()
    if not preferred_norm or current_python == preferred_norm:
        return None
    if not _python_supports_module(preferred_python, "PySide6"):
        return None
    args = list(argv if argv is not None else sys.argv)
    script_path = os.path.abspath(__file__)
    cmd = [preferred_python, script_path, *args[1:]]
    try:
        sys.stderr.write(
            "Aviso: a reencaminhar o arranque Qt para a .venv local do projeto.\n"
        )
    except Exception:
        pass
    return int(subprocess.call(cmd, cwd=os.path.dirname(script_path)))


def _print_qt_setup_hint(missing_name: str = "PySide6") -> None:
    module_name = str(missing_name or "PySide6").strip() or "PySide6"
    try:
        sys.stderr.write(
            f"Erro de arranque Qt: dependencia em falta ({module_name}).\n"
            "Use a .venv local deste projeto:\n"
            "  py -m venv .venv\n"
            "  .venv\\Scripts\\python.exe -m pip install --upgrade pip\n"
            "  .venv\\Scripts\\python.exe -m pip install -r requirements-qt.txt\n"
            "Depois arranque com:\n"
            "  .venv\\Scripts\\python.exe main.py\n"
        )
    except Exception:
        pass


def run_qt_frontend(argv=None) -> int:
    args = list(argv if argv is not None else sys.argv)
    delegated_exit = _delegate_qt_launch_to_local_venv(args)
    if delegated_exit is not None:
        return int(delegated_exit)
    _assert_mysql_runtime_ready()
    # Garante que o backend Qt reutiliza este modulo em vez de reimportar outra copia.
    sys.modules.setdefault("main", sys.modules[__name__])
    try:
        from lugest_qt.app import main as qt_main
    except ModuleNotFoundError as exc:
        missing_name = str(getattr(exc, "name", "") or "").strip()
        if missing_name == "PySide6" or missing_name.startswith("PySide6."):
            _print_qt_setup_hint(missing_name)
            return 1
        raise

    return int(qt_main(args) or 0)


def run_desktop_entry(argv=None) -> int:
    args = list(argv if argv is not None else sys.argv)
    admin_cli_result = handle_admin_setup_cli(args)
    if admin_cli_result is not None:
        return int(admin_cli_result or 0)
    filtered_args = [args[0]] if args else ["main.py"]
    legacy_requested = False
    for raw in args[1:]:
        if raw == "--legacy-ui":
            legacy_requested = True
            continue
        if raw == "--qt-ui":
            continue
        filtered_args.append(raw)
    if legacy_requested:
        try:
            sys.stderr.write("Aviso: o modo --legacy-ui foi descontinuado. O desktop arranca sempre em Qt.\n")
        except Exception:
            pass
    return run_qt_frontend(filtered_args)


if __name__ == "__main__":
    raise SystemExit(run_desktop_entry())

