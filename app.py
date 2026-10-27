import html
import io
import os
import re
import unicodedata
from typing import List, Optional
from datetime import datetime
import pandas as pd
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -------------------------------------------------------------
# HELPERS DE LIMPIEZA Y NORMALIZACION
# -------------------------------------------------------------

def clean_col(s: str) -> str:
    """Normaliza un nombre de columna: quita acentos, simbolos y minusculas."""
    if not isinstance(s, str):
        return ""
    nfkd = unicodedata.normalize('NFKD', s)
    no_acc = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r'[^a-zA-Z0-9]', '', no_acc).lower()

def clean_phone_str(val) -> str:
    """Limpia y estandariza números telefónicos al formato venezolano 584..."""
    if pd.isna(val) or val is None or val == "":
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r'\D', '', s)
    if not digits:
        return ""
    if digits.startswith("58") and len(digits) >= 12:
        return digits
    elif len(digits) == 10 and digits.startswith("4"):
        return f"58{digits}"
    elif len(digits) == 11 and digits.startswith("04"):
        return f"58{digits[1:]}"
    return digits

def parse_number_value(val) -> float:
    """Convierte de forma segura cualquier formato numerico (US, Latino, con $, Bs, etc.) a float."""
    if pd.isna(val) or val == "" or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s == "-":
        return 0.0
    
    # Remover simbolos de monedas y espacios
    s = re.sub(r'[\$\s\€\£]|Bs\.?|USD', '', s, flags=re.IGNORECASE).strip()
    if not s:
        return 0.0
    
    # Si contiene coma y punto
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            # Formato latino/europeo: 1.234,56
            s = s.replace('.', '').replace(',', '.')
        else:
            # Formato US: 1,234.56
            s = s.replace(',', '')
    elif ',' in s:
        parts = s.split(',')
        if len(parts) == 2 and len(parts[1]) in [1, 2]:
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
            
    try:
        return float(s)
    except Exception:
        match = re.search(r'[-+]?\d*\.?\d+', s)
        return float(match.group(0)) if match else 0.0

def assign_tramo_mora(dias: int) -> str:
    """Asigna el tramo oficial de mora segun la antiguedad de la deuda."""
    if dias <= 30:
        return "1 a 30 días (Cobranza preventiva / temprana)"
    elif dias <= 60:
        return "31 a 60 días (Cobranza administrativa)"
    elif dias <= 90:
        return "61 a 90 días (Reestructuración / negociación)"
    elif dias <= 180:
        return "91 a 180 días (Pre-jurídica)"
    else:
        return "181+ días (Castigo / cartera pesada)"

# -------------------------------------------------------------
# CONFIGURACIÓN FIJA DE LOS 5 TRAMOS OFICIALES DE MORA
# -------------------------------------------------------------
OFFICIAL_TRAMOS_CONFIG = [
    {
        "id": "1-30",
        "grupo": "1 a 30 días (Cobranza preventiva / temprana)",
        "etiqueta": "1 a 30 días",
        "fase": "Cobranza preventiva / temprana",
        "min": 1,
        "max": 30,
        "color": "#10b981",
        "badge_class": "bg-emerald-50 text-emerald-800 border-emerald-200 border text-[10px] font-bold px-2 py-0.5 rounded-full"
    },
    {
        "id": "31-60",
        "grupo": "31 a 60 días (Cobranza administrativa)",
        "etiqueta": "31 a 60 días",
        "fase": "Cobranza administrativa",
        "min": 31,
        "max": 60,
        "color": "#3b82f6",
        "badge_class": "bg-blue-50 text-blue-800 border-blue-200 border text-[10px] font-bold px-2 py-0.5 rounded-full"
    },
    {
        "id": "61-90",
        "grupo": "61 a 90 días (Reestructuración / negociación)",
        "etiqueta": "61 a 90 días",
        "fase": "Reestructuración / negociación",
        "min": 61,
        "max": 90,
        "color": "#f59e0b",
        "badge_class": "bg-amber-50 text-amber-800 border-amber-200 border text-[10px] font-bold px-2 py-0.5 rounded-full"
    },
    {
        "id": "91-180",
        "grupo": "91 a 180 días (Pre-jurídica)",
        "etiqueta": "91 a 180 días",
        "fase": "Pre-jurídica (Mayor monto)",
        "min": 91,
        "max": 180,
        "color": "#f97316",
        "badge_class": "bg-orange-50 text-orange-800 border-orange-200 border text-[10px] font-bold px-2 py-0.5 rounded-full"
    },
    {
        "id": "181+",
        "grupo": "181+ días (Castigo / cartera pesada)",
        "etiqueta": "181+ días",
        "fase": "Castigo / cartera pesada",
        "min": 181,
        "max": 9999,
        "color": "#ef4444",
        "badge_class": "bg-rose-50 text-rose-800 border-rose-200 border text-[10px] font-bold px-2 py-0.5 rounded-full"
    }
]

def safe_read_bytes(path: str) -> bytes:
    """Lee un archivo en bytes de forma segura, con fallback para archivos bloqueados por Excel en Windows."""
    try:
        with open(path, "rb") as f:
            return f.read()
    except (PermissionError, IOError):
        if os.name == 'nt':
            try:
                import ctypes
                kernel32 = ctypes.windll.kernel32
                GENERIC_READ = 0x80000000
                FILE_SHARE_ALL = 1 | 2 | 4  # READ, WRITE, DELETE
                OPEN_EXISTING = 3
                FILE_ATTRIBUTE_NORMAL = 0x80
                h = kernel32.CreateFileW(path, GENERIC_READ, FILE_SHARE_ALL, None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
                if h != -1 and h != 0xFFFFFFFFFFFFFFFF and h != 0xFFFFFFFF:
                    size_high = ctypes.c_ulong(0)
                    size_low = kernel32.GetFileSize(h, ctypes.byref(size_high))
                    total_size = (size_high.value << 32) + size_low
                    buf = ctypes.create_string_buffer(total_size)
                    bytes_read = ctypes.c_ulong(0)
                    kernel32.ReadFile(h, buf, total_size, ctypes.byref(bytes_read), None)
                    kernel32.CloseHandle(h)
                    return buf.raw[:bytes_read.value]
            except Exception:
                pass
        raise

# Diccionario de patrones para mapear columnas de Cartera
TARGET_PATTERNS = {
    "Dias de Mora": [
        r"^diasdemora$", r"^diasmora$", r"^diasvencidos?$", r"^dias$", 
        r"^m.xdediasdemora$", r"^maxdediasdemora$", r".*diasdemora.*", r".*diasvencid.*", r".*mora.*dias.*"
    ],
    "Cedula Cliente": [
        r"^cedulacliente$", r"^cedula$", r"^ceduladeudor$", r"^identificacion$", 
        r"^dni$", r"^rut$", r"^etiquetasdefila$", r"^etiquetadefila$", r"^nrocedula$"
    ],
    "Nombre Cliente": [
        r"^nombrecliente$", r"^nombre$", r"^cliente$", r"^deudor$", r"^nombrey?apellido$"
    ],
    "Pendiente Tarifa": [
        r"^pendientetarifa$", r"^saldotarifa$", r"^tarifapendiente$", 
        r"^sumadependientetarifa$", r"^tarifa$"
    ],
    "Pendiente Cuota": [
        r"^pendientecuota$", r"^saldocuota$", r"^cuotapendiente$", 
        r"^sumadependientecuota$", r"^cuota$"
    ],
    "Pendiente Mora": [
        r"^pendientemora$", r"^saldomora$", r"^morapendiente$", 
        r"^sumadependientemora$", r"^montomora$"
    ],
    "Total": [
        r"^total$", r"^sumadetotal$", r"^montototal$", r"^saldototal$", r"^totalvencido$"
    ],
    "Programa": [
        r"^programa$", r"^campana/programa$", r"^plan$", r"^campana$", r"^campana2$"
    ],
    "Tipo": [
        r"^tipo$", r"^tipodecuota$", r"^estatus$", r"^estado$", r"^status$"
    ],
    "Identificador de credito": [
        r"^identificadordecredito$", r"^idcredito$", r"^credito$", r"^contrato$", r"^nrocredito$"
    ],
    "Canal": [
        r"^canal$", r"^sucursal$", r"^agencia$"
    ]
}

def map_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Mapea columnas inteligentemente garantizando que NUNCA se generen duplicados."""
    cols = []
    seen = {}
    for c in df.columns:
        c_str = str(c).replace('\ufeff', '').strip()
        if c_str in seen:
            seen[c_str] += 1
            cols.append(f"{c_str}_{seen[c_str]}")
        else:
            seen[c_str] = 0
            cols.append(c_str)
    df.columns = cols

    matched = {}
    used_cols = set()

    for target, patterns in TARGET_PATTERNS.items():
        for pat in patterns:
            found = None
            for col in df.columns:
                if col in used_cols:
                    continue
                cleaned = clean_col(str(col))
                if re.match(pat, cleaned):
                    found = col
                    break
            if found:
                matched[found] = target
                used_cols.add(found)
                break

    return df.rename(columns=matched)

def parse_bytes_to_dataframe(contents: bytes, filename: str) -> pd.DataFrame:
    """Lee bytes en memoria a DataFrame soportando CSV y Excel (.xlsx, .xls) con calamine y openpyxl."""
    filename_lower = filename.lower()
    
    # 1. Es un CSV
    if filename_lower.endswith(".csv"):
        encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
        separators = [None, ",", ";", "\t"]
        for enc in encodings:
            for sep in separators:
                try:
                    temp_df = pd.read_csv(io.BytesIO(contents), encoding=enc, sep=sep, engine="python")
                    if len(temp_df.columns) > 1:
                        return temp_df
                except Exception:
                    continue
        raise HTTPException(status_code=400, detail="No se pudo interpretar el archivo CSV. Verifica su formato y delimitadores.")
    
    # 2. Es Excel (.xlsx o .xls)
    engines = ["calamine", "openpyxl", "xlrd"]
    xl = None
    for eng in engines:
        try:
            xl = pd.ExcelFile(io.BytesIO(contents), engine=eng)
            break
        except Exception:
            continue
            
    if xl is None:
        try:
            html_dfs = pd.read_html(io.BytesIO(contents))
            if html_dfs:
                return html_dfs[0]
        except Exception:
            pass
        raise HTTPException(status_code=400, detail="No se pudo abrir el archivo Excel. Verifica que no esté dañado.")
        
    # Optimizacion: Si solo tiene 1 hoja (estandar en la gran mayoria de casos), se procesa directamente
    if len(xl.sheet_names) == 1:
        return xl.parse(xl.sheet_names[0])

    # Sistema de seleccion inteligente de hoja para archivos con multiples pestañas
    sheet_scores = []
    for s in xl.sheet_names:
        try:
            sample = xl.parse(s, nrows=4)
            cols = [clean_col(str(c)) for c in sample.columns]
            score = 0
            s_norm = clean_col(s)
            
            if any(k in s_norm for k in ["morageneral", "cartera", "base", "detalle", "datadeclientes", "clientes"]):
                score += 50
            elif any(k in s_norm for k in ["consolidado", "callcenter", "mora"]):
                score += 30
                
            has_dias = any(any(re.match(p, c) for p in TARGET_PATTERNS["Dias de Mora"]) for c in cols)
            has_cedula = any(any(re.match(p, c) for p in TARGET_PATTERNS["Cedula Cliente"]) for c in cols)
            has_tel = any(clean_col(c).startswith("tel") for c in cols)
            
            if has_dias:
                score += 60
            if has_cedula:
                score += 60
            if has_tel:
                score += 40
                
            score += min(len(sample.columns) * 2, 40)
            sheet_scores.append((s, score))
        except Exception:
            sheet_scores.append((s, -1))
            
    sheet_scores.sort(key=lambda x: x[1], reverse=True)
    chosen_sheet = sheet_scores[0][0] if sheet_scores and sheet_scores[0][1] > 0 else xl.sheet_names[0]
    return xl.parse(chosen_sheet)

def standardize_cartera_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Valida, estandariza y asegura tipos numericos para la cartera."""
    df = map_column_names(df)
    
    req_cols = ["Dias de Mora", "Cedula Cliente"]
    missing = [c for c in req_cols if c not in df.columns]
    if missing:
        cols_found = [str(c) for c in df.columns.tolist()][:15]
        raise HTTPException(
            status_code=400, 
            detail=f"Faltan columnas clave obligatorias en Cartera: {missing}. El archivo tiene estas columnas: {cols_found}"
        )
    
    df["Dias de Mora"] = df["Dias de Mora"].apply(parse_number_value).fillna(0).astype(int)
    df["Tramo_Mora"] = df["Dias de Mora"].apply(assign_tramo_mora)
    
    for c in ["Pendiente Tarifa", "Pendiente Cuota", "Pendiente Mora", "Total"]:
        if c in df.columns:
            df[c] = df[c].apply(parse_number_value).fillna(0.0)
        else:
            df[c] = 0.0
            
    tarifa_sum = float(df["Pendiente Tarifa"].sum())
    cuota_sum = float(df["Pendiente Cuota"].sum())
    mora_sum = float(df["Pendiente Mora"].sum())
    total_sum = float(df["Total"].sum())
    
    if (tarifa_sum + cuota_sum + mora_sum) == 0.0 and total_sum > 0.0:
        df["Pendiente Tarifa"] = df["Total"]
    elif total_sum == 0.0 and (tarifa_sum + cuota_sum + mora_sum) > 0.0:
        df["Total"] = df["Pendiente Tarifa"] + df["Pendiente Cuota"] + df["Pendiente Mora"]

    if "Programa" not in df.columns:
        df["Programa"] = "General"
    else:
        df["Programa"] = df["Programa"].fillna("General").astype(str).str.strip()
        df["Programa"] = df["Programa"].replace({"": "General", "nan": "General", "None": "General"})
        df["Programa"] = df["Programa"].str.replace(r"CAMPA.*AS", "CAMPAÑAS", regex=True)

    if "Tipo" not in df.columns:
        df["Tipo"] = "General"
    else:
        df["Tipo"] = df["Tipo"].fillna("General").astype(str).str.strip()
        df["Tipo"] = df["Tipo"].replace({"": "General", "nan": "General", "None": "General"})

    df["Cedula Cliente"] = df["Cedula Cliente"].fillna("").astype(str).str.strip()
    df["Cedula_Key"] = df["Cedula Cliente"].str.upper()

    if "Nombre Cliente" not in df.columns:
        df["Nombre Cliente"] = "Cliente Sin Nombre"
    else:
        df["Nombre Cliente"] = df["Nombre Cliente"].fillna("").astype(str).str.strip()

    return df

def standardize_contacts_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Estandariza columnas, teléfonos, campaña y cédula del archivo de contactos."""
    col_map = {}
    for c in df.columns:
        norm = clean_col(str(c))
        if any(re.match(p, norm) for p in [r"^cedula.*", r"^dni$", r"^identificacion$", r"^etiquetasdefila$"]):
            col_map[c] = "Cedula"
        elif any(re.match(p, norm) for p in [r"^telefono1$", r"^tel1$", r"^telf1$", r"^celular1$", r"^celular$", r"^telefono$"]):
            col_map[c] = "Telefono_1"
        elif any(re.match(p, norm) for p in [r"^telefono2$", r"^tel2$", r"^telf2$", r"^celular2$"]):
            col_map[c] = "Telefono_2"
        elif any(re.match(p, norm) for p in [r"^telefono3$", r"^tel3$", r"^telf3$", r"^celular3$"]):
            col_map[c] = "Telefono_3"
        elif any(re.match(p, norm) for p in [r"^correo.*", r"^email.*", r"^mail$"]):
            col_map[c] = "Correo"
        elif any(re.match(p, norm) for p in [r"^nombre.*", r"^cliente$"]):
            col_map[c] = "Nombre_Contacto"
        elif any(re.match(p, norm) for p in [r"^tipodecampana.*", r"^campana.*", r"^tipocampana.*"]):
            col_map[c] = "Tipo_Campana"
        elif any(re.match(p, norm) for p in [r"^programa.*", r"^plan.*"]):
            col_map[c] = "Programa_Contacto"

    df = df.rename(columns=col_map)
    if "Cedula" not in df.columns:
        raise HTTPException(status_code=400, detail="El archivo de contactos debe contener una columna de Cédula o Identificación.")

    for col in ["Telefono_1", "Telefono_2", "Telefono_3"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_phone_str)
        else:
            df[col] = ""

    for col in ["Correo", "Nombre_Contacto", "Tipo_Campana", "Programa_Contacto"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            df[col] = ""

    df["Cedula"] = df["Cedula"].fillna("").astype(str).str.strip().str.upper()
    df["Cedula_Key"] = df["Cedula"]
    return df.drop_duplicates(subset=["Cedula_Key"])

def detect_file_type(df: pd.DataFrame) -> str:
    """Detecta inteligentemente si un DataFrame corresponde a Cartera ('cartera') o a Contactos ('contacts')."""
    cleaned_cols = [clean_col(str(c)) for c in df.columns]
    
    # Indicadores de Cartera / Mora
    has_dias = any(any(re.match(p, c) for p in TARGET_PATTERNS["Dias de Mora"]) for c in cleaned_cols)
    has_cuota = any(any(re.match(p, c) for p in TARGET_PATTERNS["Pendiente Cuota"]) for c in cleaned_cols)
    has_mora = any(any(re.match(p, c) for p in TARGET_PATTERNS["Pendiente Mora"]) for c in cleaned_cols)
    
    # Indicadores de Contactos / Telefonos
    has_phone = any(any(re.match(p, c) for p in [r"^telefono.*", r"^tel\d?$", r"^celular.*", r"^telf.*"]) for c in cleaned_cols)
    has_email = any(any(re.match(p, c) for p in [r"^correo.*", r"^email.*", r"^mail$"]) for c in cleaned_cols)
    has_campana = any(any(re.match(p, c) for p in [r"^tipodecampana.*", r"^campana.*", r"^tipocampana.*"]) for c in cleaned_cols)
    
    if has_dias or (has_cuota and has_mora):
        return "cartera"
    elif has_phone or has_email or has_campana:
        return "contacts"
    return "unknown"

# -------------------------------------------------------------
# FASTAPI APP & STATE
# -------------------------------------------------------------

app = FastAPI(title="Motor de Reestructuracion, Segmentacion y Cruce de Cartera")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DATA_STORE = {
    "df_cartera_raw": None,
    "cartera_filename": "",
    "df_contacts_raw": None,
    "contacts_filename": "",
    "df_original": None,      # Enriquecida y unificada
    "df_filtered": None,      # Filtrada segun criterios
    "filter_info": {},
    "df_matched": None,       # Agregada a nivel deudor para Call Center
    "cartera_activa_manual": None # Monto manual de cartera activa para cálculo de % mora
}

DEFAULT_MORA_PATH = os.path.join(os.path.dirname(__file__), "mora.csv")
DEFAULT_CONTACTS_PATH = os.path.join(os.path.dirname(__file__), "Base de Datos Correos y Teléfonos al 06-08 (2).xlsx")

def enrich_and_merge_datasets():
    """Cruza y unifica la cartera con la base de contactos de forma instantanea."""
    if DATA_STORE["df_cartera_raw"] is None:
        return
        
    df_mora = DATA_STORE["df_cartera_raw"].copy()
    df_contacts = DATA_STORE["df_contacts_raw"]
    
    if df_contacts is not None and not df_contacts.empty:
        cols_to_merge = [
            "Cedula_Key", "Tipo_Campana", "Programa_Contacto", 
            "Telefono_1", "Telefono_2", "Telefono_3", "Correo", "Nombre_Contacto"
        ]
        avail_cols = [c for c in cols_to_merge if c in df_contacts.columns]
        
        merged = pd.merge(
            df_mora,
            df_contacts[avail_cols],
            on="Cedula_Key",
            how="left"
        )
        
        # Enriquecer Nombre si en cartera estaba vacio
        if "Nombre_Contacto" in merged.columns:
            mask_no_name = (merged["Nombre Cliente"].isna()) | (merged["Nombre Cliente"].astype(str).str.strip().isin(["", "Cliente Sin Nombre", "nan"]))
            merged.loc[mask_no_name, "Nombre Cliente"] = merged.loc[mask_no_name, "Nombre_Contacto"]
            
        # Reconciliar Programa si en cartera venia General
        if "Programa_Contacto" in merged.columns:
            mask_gen_prog = (merged["Programa"] == "General") & (merged["Programa_Contacto"].fillna("") != "")
            merged.loc[mask_gen_prog, "Programa"] = merged.loc[mask_gen_prog, "Programa_Contacto"]

        # Normalizar Tipo de Campaña
        if "Tipo_Campana" in merged.columns:
            merged["Tipo_Campana"] = merged["Tipo_Campana"].fillna("Sin Campaña Asignada").astype(str).str.strip()
            merged["Tipo_Campana"] = merged["Tipo_Campana"].replace({"": "Sin Campaña Asignada", "nan": "Sin Campaña Asignada"})
        else:
            merged["Tipo_Campana"] = "Sin Campaña Asignada"

        for c in ["Telefono_1", "Telefono_2", "Telefono_3", "Correo"]:
            if c in merged.columns:
                merged[c] = merged[c].fillna("").astype(str).str.strip()
            else:
                merged[c] = ""
    else:
        merged = df_mora.copy()
        merged["Tipo_Campana"] = "Sin Campaña Asignada"
        merged["Telefono_1"] = ""
        merged["Telefono_2"] = ""
        merged["Telefono_3"] = ""
        merged["Correo"] = ""

    DATA_STORE["df_original"] = merged
    DATA_STORE["df_filtered"] = merged
    build_debtor_matched_dataset(merged)

def build_debtor_matched_dataset(df: pd.DataFrame):
    """Consolida la cartera a nivel deudor (1 fila por deudor) para el modulo Call Center."""
    if df is None or df.empty:
        DATA_STORE["df_matched"] = None
        return
        
    df = df.copy()
    df["_Fila_Total"] = df["Pendiente Tarifa"] + df["Pendiente Cuota"] + df["Pendiente Mora"]
    if df["_Fila_Total"].sum() == 0 and "Total" in df.columns:
        df["_Fila_Total"] = df["Total"]
        
    debtors = df.groupby("Cedula_Key", as_index=False).agg(
        Cedula=("Cedula Cliente", "first"),
        Nombre_Cliente=("Nombre Cliente", "first"),
        Cuotas_Vencidas=("Cedula Cliente", "count"),
        Monto_Total_USD=("_Fila_Total", "sum"),
        Pendiente_Tarifa=("Pendiente Tarifa", "sum"),
        Pendiente_Cuota=("Pendiente Cuota", "sum"),
        Pendiente_Mora=("Pendiente Mora", "sum"),
        Max_Dias_Mora=("Dias de Mora", "max"),
        Tipo_Campana=("Tipo_Campana", "first"),
        Programa=("Programa", "first"),
        Tipo=("Tipo", lambda x: "/".join(sorted(set(str(v) for v in x if str(v) not in ["", "nan", "None", "General"]))) or "General"),
        Telefono_1=("Telefono_1", "first"),
        Telefono_2=("Telefono_2", "first"),
        Telefono_3=("Telefono_3", "first"),
        Correo=("Correo", "first")
    )

    debtors["Tramo_Mora"] = debtors["Max_Dias_Mora"].apply(assign_tramo_mora)
    debtors["Monto_Total_USD"] = debtors["Monto_Total_USD"].round(2)
    debtors["Pendiente_Tarifa"] = debtors["Pendiente_Tarifa"].round(2)
    debtors["Pendiente_Cuota"] = debtors["Pendiente_Cuota"].round(2)
    debtors["Pendiente_Mora"] = debtors["Pendiente_Mora"].round(2)
    
    debtors = debtors.sort_values(by="Monto_Total_USD", ascending=False)
    DATA_STORE["df_matched"] = debtors

def calculate_totals(df: pd.DataFrame):
    if df is None or df.empty:
        return {
            "total_cuotas": 0,
            "total_clientes": 0,
            "total_creditos": 0,
            "monto_total": 0.0,
            "pendiente_tarifa": 0.0,
            "pendiente_cuota": 0.0,
            "pendiente_mora": 0.0,
            "promedio_cliente": 0.0,
            "total_con_telefono": 0,
            "pct_con_telefono": 0.0,
            "total_con_correo": 0,
            "pct_con_correo": 0.0,
            "monto_contactable": 0.0,
            "pct_monto_contactable": 0.0
        }
    
    total_cuotas = len(df)
    total_clientes = int(df["Cedula Cliente"].nunique()) if "Cedula Cliente" in df.columns else total_cuotas
    total_creditos = int(df["Identificador de credito"].nunique()) if "Identificador de credito" in df.columns else 0
    p_tarifa = float(df["Pendiente Tarifa"].sum()) if "Pendiente Tarifa" in df.columns else 0.0
    p_cuota = float(df["Pendiente Cuota"].sum()) if "Pendiente Cuota" in df.columns else 0.0
    p_mora = float(df["Pendiente Mora"].sum()) if "Pendiente Mora" in df.columns else 0.0
    
    monto_total = p_tarifa + p_cuota + p_mora
    if monto_total == 0.0 and "Total" in df.columns:
        monto_total = float(df["Total"].sum())
        p_tarifa = monto_total

    promedio = monto_total / total_clientes if total_clientes > 0 else 0.0

    # Metricas de contactabilidad
    has_tel = (df["Telefono_1"].fillna("") != "") | (df["Telefono_2"].fillna("") != "") | (df["Telefono_3"].fillna("") != "")
    has_mail = df["Correo"].fillna("") != ""
    has_contact = has_tel | has_mail

    clientes_con_tel = df[has_tel]["Cedula Cliente"].nunique() if "Cedula Cliente" in df.columns else int(has_tel.sum())
    clientes_con_mail = df[has_mail]["Cedula Cliente"].nunique() if "Cedula Cliente" in df.columns else int(has_mail.sum())

    row_total = df["Pendiente Tarifa"] + df["Pendiente Cuota"] + df["Pendiente Mora"]
    if row_total.sum() == 0 and "Total" in df.columns:
        row_total = df["Total"]
        
    monto_contactable = float(row_total[has_contact].sum())

    pct_tel = round((clientes_con_tel / total_clientes * 100), 1) if total_clientes > 0 else 0.0
    pct_mail = round((clientes_con_mail / total_clientes * 100), 1) if total_clientes > 0 else 0.0
    pct_monto_contactable = round((monto_contactable / monto_total * 100), 1) if monto_total > 0 else 0.0

    return {
        "total_cuotas": total_cuotas,
        "total_clientes": total_clientes,
        "total_creditos": total_creditos,
        "monto_total": round(monto_total, 2),
        "pendiente_tarifa": round(p_tarifa, 2),
        "pendiente_cuota": round(p_cuota, 2),
        "pendiente_mora": round(p_mora, 2),
        "promedio_cliente": round(promedio, 2),
        "total_con_telefono": int(clientes_con_tel),
        "pct_con_telefono": pct_tel,
        "total_con_correo": int(clientes_con_mail),
        "pct_con_correo": pct_mail,
        "monto_contactable": round(monto_contactable, 2),
        "pct_monto_contactable": pct_monto_contactable,
    }

def get_cutoff_date(df: pd.DataFrame) -> str:
    """Extrae la fecha de corte de la cartera o la fecha actual."""
    if df is not None and not df.empty and "Fecha Vencimiento" in df.columns:
        try:
            dates = pd.to_datetime(df["Fecha Vencimiento"], errors="coerce").dropna()
            if not dates.empty:
                return dates.max().strftime("%d-%m-%Y")
        except Exception:
            pass
    return datetime.now().strftime("%d-%m-%Y")

def calculate_global_tramos_summary(df: pd.DataFrame, cartera_activa: Optional[float] = None):
    """Calcula el Resumen Fijo Oficial por Tramos de Mora para la Cartera Total,
    calculando los porcentajes respecto a la Cartera Activa Total si ha sido ingresada."""
    if df is None or df.empty:
        return {
            "cutoff_date": datetime.now().strftime("%d-%m-%Y"),
            "total_cartera_usd": 0.0,
            "total_cuotas": 0,
            "total_clientes": 0,
            "cartera_activa_usd": None,
            "total_porcentaje": None,
            "tramos": []
        }

    df = df.copy()
    if "Pendiente Tarifa" in df.columns and df["Pendiente Tarifa"].sum() > 0:
        df["_Tot"] = df["Pendiente Tarifa"]
    elif "Total" in df.columns:
        df["_Tot"] = df["Total"]
    elif "Pendiente Cuota" in df.columns and "Pendiente Mora" in df.columns and (df["Pendiente Cuota"].sum() + df["Pendiente Mora"].sum()) > 0:
        df["_Tot"] = df["Pendiente Cuota"] + df["Pendiente Mora"]
    else:
        df["_Tot"] = 0.0

    df_active = df[df["_Tot"] > 0].copy() if (df["_Tot"] > 0).any() else df.copy()

    total_cartera_usd = round(float(df_active["_Tot"].sum()), 2)
    total_cuotas = len(df_active)
    total_clientes = int(df_active["Cedula Cliente"].nunique()) if "Cedula Cliente" in df_active.columns else total_cuotas

    # Denominador para el cálculo de porcentaje sobre la Cartera Activa Total disponible
    target_ca = cartera_activa if (cartera_activa is not None and cartera_activa > 0) else DATA_STORE.get("cartera_activa_manual")
    denom = float(target_ca) if (target_ca is not None and target_ca > 0) else None
    total_pct = round((total_cartera_usd / denom * 100), 2) if denom else None

    tramos = []
    for t_cfg in OFFICIAL_TRAMOS_CONFIG:
        sub = df_active[df_active["Tramo_Mora"] == t_cfg["grupo"]]
        cuotas_cnt = len(sub)
        clientes_cnt = int(sub["Cedula Cliente"].nunique()) if "Cedula Cliente" in sub.columns else cuotas_cnt
        p_tarifa = round(float(sub["Pendiente Tarifa"].sum()), 2) if "Pendiente Tarifa" in sub.columns else 0.0
        p_cuota = round(float(sub["Pendiente Cuota"].sum()), 2) if "Pendiente Cuota" in sub.columns else 0.0
        p_mora = round(float(sub["Pendiente Mora"].sum()), 2) if "Pendiente Mora" in sub.columns else 0.0
        monto_grp = round(float(sub["_Tot"].sum()), 2)
        pct = round((monto_grp / denom * 100), 2) if denom else None

        tramos.append({
            "id": t_cfg["id"],
            "fase": t_cfg["fase"],
            "etiqueta": t_cfg["etiqueta"],
            "grupo": t_cfg["grupo"],
            "color": t_cfg["color"],
            "badge_class": t_cfg["badge_class"],
            "min_dias": t_cfg["min"],
            "max_dias": t_cfg["max"],
            "cuotas": cuotas_cnt,
            "clientes": clientes_cnt,
            "pendiente_tarifa": p_tarifa,
            "pendiente_cuota": p_cuota,
            "pendiente_mora": p_mora,
            "cuota_mas_mora": round(p_cuota + p_mora, 2),
            "total_monto": monto_grp,
            "porcentaje": pct
        })

    return {
        "cutoff_date": get_cutoff_date(df),
        "total_cartera_usd": total_cartera_usd,
        "total_cuotas": total_cuotas,
        "total_clientes": total_clientes,
        "cartera_activa_usd": denom,
        "total_porcentaje": total_pct,
        "tramos": tramos
    }

def normalize_campaign_name(name: str) -> str:
    """Normaliza y unifica los nombres de campaña con variaciones regionales en macro-familias."""
    s = str(name).strip() if name is not None else ""
    s_lower = s.lower()
    
    if not s or s_lower in ["nan", "none", "", "sin campaña", "sin campana asignada"]:
        return "Sin Campaña Asignada"
        
    if "pre aprob" in s_lower or "pre-aprob" in s_lower or "preaprob" in s_lower:
        return "Pre-aprobados"
    if "web no fibex" in s_lower:
        return "Web no Fibex"
    if "vuelve a casa" in s_lower:
        return "Vuelve a Casa"
    if "nuevo" in s_lower:
        return "Nuevos"
    if "referid" in s_lower:
        return "Referidos"
    if "invitad" in s_lower:
        return "Invitados"
    if "asesor" in s_lower:
        return "Asesores de Seguros"
    if "empleados aragua" in s_lower:
        return "Empleados Aragua"
    if "emplead" in s_lower:
        return "Empleados"
    if "aliados club fibex" in s_lower or "club fibex" in s_lower:
        return "Aliados Club Fibex"
    if "aliados integral" in s_lower or "aliados" in s_lower:
        return "Aliados Integrales"
    if "barrido" in s_lower:
        return "Barridos de Zona - Caracas"
    if "clientes web" in s_lower or "web fibex" in s_lower:
        return "Web Fibex"
    if "internacional" in s_lower:
        return "Trabajadores La Internacional"
    if "warrior" in s_lower:
        return "Inv. Warrior Tegnologi"
    if "triple j" in s_lower:
        return "Triple J"
    if "turicopy" in s_lower:
        return "TuriCopy"
    if "piloto" in s_lower:
        return "Piloto"
    if "fiber hogar" in s_lower:
        return "Fiber Hogar Telecom"
    if "inter i" in s_lower:
        return "Inter IVU"
    if "vip" in s_lower:
        return "VIP"
    if "proveed" in s_lower:
        return "Proveedores"
    if "agencia" in s_lower:
        return "Agencia"
    if "jac" in s_lower:
        return "Financiadora JAC"
    if "cobraca" in s_lower:
        return "Grupo Cobraca"
    if "gonzalez" in s_lower:
        return "Grupo Gonzalez Castillo"
    if "avila" in s_lower:
        return "Cartera Avila Pay"
        
    return s

def calculate_campanas_tramos_matrix(df: pd.DataFrame):
    """Calcula la matriz bidimensional de Campañas vs 5 Tramos Oficiales de Mora."""
    if df is None or df.empty:
        return {
            "cutoff_date": datetime.now().strftime("%d-%m-%Y"),
            "total_cartera_usd": 0.0,
            "total_clientes": 0,
            "total_cuotas": 0,
            "matrix": [],
            "totals_by_tramo": {t["id"]: {"monto": 0.0, "clientes": 0, "cuotas": 0} for t in OFFICIAL_TRAMOS_CONFIG}
        }
        
    df = df.copy()
    if "Pendiente Tarifa" in df.columns and df["Pendiente Tarifa"].sum() > 0:
        df["_Tot"] = df["Pendiente Tarifa"]
    elif "Total" in df.columns:
        df["_Tot"] = df["Total"]
    elif "Pendiente Cuota" in df.columns and "Pendiente Mora" in df.columns and (df["Pendiente Cuota"].sum() + df["Pendiente Mora"].sum()) > 0:
        df["_Tot"] = df["Pendiente Cuota"] + df["Pendiente Mora"]
    else:
        df["_Tot"] = 0.0

    df_active = df[df["_Tot"] > 0].copy() if (df["_Tot"] > 0).any() else df.copy()

    total_cartera = float(df_active["_Tot"].sum())
    total_clientes_global = int(df_active["Cedula Cliente"].nunique()) if "Cedula Cliente" in df_active.columns else len(df_active)
    total_cuotas_global = len(df_active)
    
    col_campana = "Tipo_Campana" if "Tipo_Campana" in df_active.columns else "Programa"
    df_active["_Macro_Campana"] = df_active[col_campana].apply(normalize_campaign_name)

    matrix = []
    totals_by_tramo = {t["id"]: {"monto": 0.0, "clientes": 0, "cuotas": 0} for t in OFFICIAL_TRAMOS_CONFIG}

    # Totales globales por tramo
    for t_cfg in OFFICIAL_TRAMOS_CONFIG:
        sub_t = df_active[df_active["Tramo_Mora"] == t_cfg["grupo"]]
        totals_by_tramo[t_cfg["id"]]["monto"] = round(float(sub_t["_Tot"].sum()), 2)
        totals_by_tramo[t_cfg["id"]]["cuotas"] = len(sub_t)
        totals_by_tramo[t_cfg["id"]]["clientes"] = int(sub_t["Cedula Cliente"].nunique()) if "Cedula Cliente" in sub_t.columns else len(sub_t)

    # Detalle por campaña
    for camp, grp in df_active.groupby("_Macro_Campana"):
        camp_tot = float(grp["_Tot"].sum())
        camp_cl = int(grp["Cedula Cliente"].nunique()) if "Cedula Cliente" in grp.columns else len(grp)
        camp_cuotas = len(grp)
        
        tramos_dict = {}
        for t_cfg in OFFICIAL_TRAMOS_CONFIG:
            sub = grp[grp["Tramo_Mora"] == t_cfg["grupo"]]
            m = float(sub["_Tot"].sum())
            cl = int(sub["Cedula Cliente"].nunique()) if "Cedula Cliente" in sub.columns else len(sub)
            cu = len(sub)
            tramos_dict[t_cfg["id"]] = {
                "monto": round(m, 2),
                "clientes": cl,
                "cuotas": cu
            }
            
        pct = round(camp_tot / total_cartera * 100, 2) if total_cartera > 0 else 0.0
        matrix.append({
            "campana": camp,
            "total_monto": round(camp_tot, 2),
            "total_clientes": camp_cl,
            "total_cuotas": camp_cuotas,
            "porcentaje": pct,
            "tramos": tramos_dict
        })

    # Ordenar por deuda descendente (las campañas más críticas primero)
    matrix.sort(key=lambda x: x["total_monto"], reverse=True)

    return {
        "cutoff_date": get_cutoff_date(df),
        "total_cartera_usd": round(total_cartera, 2),
        "total_clientes": total_clientes_global,
        "total_cuotas": total_cuotas_global,
        "matrix": matrix,
        "totals_by_tramo": totals_by_tramo
    }

def get_metadata(df: pd.DataFrame):
    has_contacts = (DATA_STORE["df_contacts_raw"] is not None and not DATA_STORE["df_contacts_raw"].empty)
    total_contacts = len(DATA_STORE["df_contacts_raw"]) if has_contacts else 0

    if df is None or df.empty:
        return {
            "total_rows": 0, "total_clientes": 0, "programas": [],
            "tipos_campana": [], "tipos": [], "canales": [], "tramos": {},
            "min_dias": 0, "max_dias": 0, "has_contacts": has_contacts,
            "total_contacts": total_contacts,
            "cartera_filename": DATA_STORE["cartera_filename"],
            "contacts_filename": DATA_STORE["contacts_filename"]
        }
        
    programas = sorted([p for p in df["Programa"].dropna().unique().tolist() if str(p).strip()]) if "Programa" in df.columns else []
    tipos_campana = sorted([t for t in df["Tipo_Campana"].dropna().unique().tolist() if str(t).strip()]) if "Tipo_Campana" in df.columns else []
    tipos = sorted([t for t in df["Tipo"].dropna().unique().tolist() if str(t).strip()]) if "Tipo" in df.columns else []
    canales = sorted([c for c in df["Canal"].dropna().unique().tolist() if str(c).strip()]) if "Canal" in df.columns else []
    
    tramos_counts = df["Tramo_Mora"].value_counts().to_dict() if "Tramo_Mora" in df.columns else {}

    min_dias = int(df["Dias de Mora"].min()) if "Dias de Mora" in df.columns and not df.empty else 0
    max_dias = int(df["Dias de Mora"].max()) if "Dias de Mora" in df.columns and not df.empty else 0

    return {
        "total_rows": len(df),
        "total_clientes": int(df["Cedula Cliente"].nunique()) if "Cedula Cliente" in df.columns else 0,
        "programas": programas,
        "tipos_campana": tipos_campana,
        "tipos": tipos,
        "canales": canales,
        "tramos": tramos_counts,
        "min_dias": min_dias,
        "max_dias": max_dias,
        "has_contacts": has_contacts,
        "total_contacts": total_contacts,
        "cartera_filename": DATA_STORE["cartera_filename"],
        "contacts_filename": DATA_STORE["contacts_filename"]
    }

class FilterRequest(BaseModel):
    min_dias: int = 1
    max_dias: int = 999
    programas: Optional[List[str]] = None
    tipos_campana: Optional[List[str]] = None
    tipos: Optional[List[str]] = None
    only_with_phone: bool = False
    group_by: str = "Tipo_Campana"
    cartera_activa: Optional[float] = None

class SetCarteraActivaRequest(BaseModel):
    cartera_activa: Optional[float] = None

# -------------------------------------------------------------
# RUTAS DE LA API
# -------------------------------------------------------------

@app.get("/api/status")
def get_status():
    cartera_loaded = DATA_STORE["df_cartera_raw"] is not None
    contacts_loaded = DATA_STORE["df_contacts_raw"] is not None
    sample_available = os.path.exists(DEFAULT_MORA_PATH)
    contacts_sample_available = os.path.exists(DEFAULT_CONTACTS_PATH)
    
    if not cartera_loaded and not contacts_loaded:
        return {
            "loaded": False,
            "contacts_loaded": False,
            "sample_available": sample_available,
            "contacts_sample_available": contacts_sample_available
        }
    
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    meta = get_metadata(active_df)
    totals = calculate_totals(active_df)
    resumen_global = calculate_global_tramos_summary(active_df)
    matriz_campanas = calculate_campanas_tramos_matrix(active_df)
    
    return {
        "loaded": cartera_loaded,
        "contacts_loaded": contacts_loaded,
        "cartera_filename": DATA_STORE["cartera_filename"],
        "contacts_filename": DATA_STORE["contacts_filename"],
        "metadata": meta,
        "totals": totals,
        "resumen_global_tramos": resumen_global,
        "matriz_campanas_tramos": matriz_campanas,
        "sample_available": sample_available,
        "contacts_sample_available": contacts_sample_available
    }

@app.post("/api/load-sample")
def load_sample():
    """Carga los archivos de ejemplo del servidor (mora.csv y contactos) unificandolos al instante."""
    if not os.path.exists(DEFAULT_MORA_PATH):
        raise HTTPException(status_code=404, detail="Archivo mora.csv no encontrado en el servidor.")
    
    # 1. Cargar Cartera
    try:
        cartera_bytes = safe_read_bytes(DEFAULT_MORA_PATH)
        df_cartera = parse_bytes_to_dataframe(cartera_bytes, "mora.csv")
        df_cartera = standardize_cartera_dataframe(df_cartera)
        DATA_STORE["df_cartera_raw"] = df_cartera
        DATA_STORE["cartera_filename"] = "mora.csv (Cartera Principal)"
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No se pudo leer mora.csv: {e}")

    # 2. Cargar Contactos si existe
    contacts_msg = ""
    if os.path.exists(DEFAULT_CONTACTS_PATH):
        try:
            contacts_bytes = safe_read_bytes(DEFAULT_CONTACTS_PATH)
            df_contacts = parse_bytes_to_dataframe(contacts_bytes, "Base de Datos Correos y Teléfonos al 06-08 (2).xlsx")
            df_contacts = standardize_contacts_dataframe(df_contacts)
            DATA_STORE["df_contacts_raw"] = df_contacts
            DATA_STORE["contacts_filename"] = "Base de Contactos y Campañas (Ejemplo)"
            contacts_msg = f" + {len(df_contacts):,} contactos unificados."
        except Exception as e:
            print("Aviso: No se pudo cargar archivo de contactos de ejemplo:", e)

    # 3. Unificar y Enriquecer
    enrich_and_merge_datasets()
    
    active_df = DATA_STORE["df_original"]
    resumen_global = calculate_global_tramos_summary(active_df)
    matriz_campanas = calculate_campanas_tramos_matrix(active_df)
    return {
        "message": f"Cartera cargada exitosamente: {len(df_cartera):,} cuotas{contacts_msg}",
        "metadata": get_metadata(active_df),
        "totals": calculate_totals(active_df),
        "resumen_global_tramos": resumen_global,
        "matriz_campanas_tramos": matriz_campanas
    }

@app.post("/api/upload")
async def upload_cartera(file: UploadFile = File(...)):
    """Sube el archivo de cartera (CSV o Excel). Si se detecta que es contactos, se auto-enruta inteligentemente."""
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
        
    try:
        raw_df = parse_bytes_to_dataframe(contents, file.filename)
        detected = detect_file_type(raw_df)
        
        # Auto-enrutamiento inteligente: si el usuario subió la base de contactos en la opción de cartera
        if detected == "contacts":
            df_contacts = standardize_contacts_dataframe(raw_df)
            DATA_STORE["df_contacts_raw"] = df_contacts
            DATA_STORE["contacts_filename"] = file.filename
            
            enrich_and_merge_datasets()
            
            active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
            totals = calculate_totals(active_df) if active_df is not None else {}
            meta = get_metadata(active_df) if active_df is not None else {}
            resumen_global = calculate_global_tramos_summary(active_df) if active_df is not None else {}
            matriz_campanas = calculate_campanas_tramos_matrix(active_df) if active_df is not None else {}
            
            return {
                "auto_routed": "contacts",
                "message": f"Detección inteligente: '{file.filename}' es la Base de Contactos ({len(df_contacts):,} clientes) y fue procesada automáticamente en su sección.",
                "total_contacts": len(df_contacts),
                "filename": file.filename,
                "metadata": meta,
                "totals": totals,
                "resumen_global_tramos": resumen_global,
                "matriz_campanas_tramos": matriz_campanas
            }

        df = standardize_cartera_dataframe(raw_df)
        DATA_STORE["df_cartera_raw"] = df
        DATA_STORE["cartera_filename"] = file.filename
        
        # Cruzar con contactos si ya estaban cargados
        enrich_and_merge_datasets()
        
        active_df = DATA_STORE["df_original"]
        resumen_global = calculate_global_tramos_summary(active_df)
        matriz_campanas = calculate_campanas_tramos_matrix(active_df)
        return {
            "message": f"Archivo de cartera '{file.filename}' procesado con éxito.",
            "metadata": get_metadata(active_df),
            "totals": calculate_totals(active_df),
            "resumen_global_tramos": resumen_global,
            "matriz_campanas_tramos": matriz_campanas
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error al procesar archivo de cartera: {str(e)}")

@app.post("/api/contacts/upload")
async def upload_contacts(file: UploadFile = File(...)):
    """Sube la base de contactos y campañas. Si se detecta que es cartera, se auto-enruta inteligentemente."""
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
        
    try:
        raw_df = parse_bytes_to_dataframe(contents, file.filename)
        detected = detect_file_type(raw_df)
        
        # Auto-enrutamiento inteligente: si el usuario subió la cartera en la opción de contactos
        if detected == "cartera":
            df = standardize_cartera_dataframe(raw_df)
            DATA_STORE["df_cartera_raw"] = df
            DATA_STORE["cartera_filename"] = file.filename
            
            enrich_and_merge_datasets()
            
            active_df = DATA_STORE["df_original"]
            resumen_global = calculate_global_tramos_summary(active_df)
            matriz_campanas = calculate_campanas_tramos_matrix(active_df)
            return {
                "auto_routed": "cartera",
                "message": f"Detección inteligente: '{file.filename}' es el Archivo de Cartera ({len(df):,} cuotas) y fue procesado automáticamente en su sección.",
                "metadata": get_metadata(active_df),
                "totals": calculate_totals(active_df),
                "resumen_global_tramos": resumen_global,
                "matriz_campanas_tramos": matriz_campanas
            }

        df = standardize_contacts_dataframe(raw_df)
        DATA_STORE["df_contacts_raw"] = df
        DATA_STORE["contacts_filename"] = file.filename
        
        # Cruzar inmediatamente con cartera
        enrich_and_merge_datasets()
        
        active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
        totals = calculate_totals(active_df) if active_df is not None else {}
        meta = get_metadata(active_df) if active_df is not None else {}
        resumen_global = calculate_global_tramos_summary(active_df) if active_df is not None else {}
        matriz_campanas = calculate_campanas_tramos_matrix(active_df) if active_df is not None else {}
        
        return {
            "message": f"Base de contactos '{file.filename}' cargada: {len(df):,} clientes indexados.",
            "total_contacts": len(df),
            "filename": file.filename,
            "metadata": meta,
            "totals": totals,
            "resumen_global_tramos": resumen_global,
            "matriz_campanas_tramos": matriz_campanas
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error al procesar archivo de contactos: {str(e)}")

@app.post("/api/process")
def process_data(req: FilterRequest):
    """Aplica filtros dinamicos de tramos, campanas, tipos y agrupa los resultados. Requiere ambos archivos obligatorios."""
    has_cartera = DATA_STORE["df_cartera_raw"] is not None
    has_contacts = DATA_STORE["df_contacts_raw"] is not None
    
    if not has_cartera or not has_contacts:
        missing = []
        if not has_cartera:
            missing.append("Archivo de Cartera / Mora (cuotas y saldos)")
        if not has_contacts:
            missing.append("Base de Contactos y Clientes (teléfonos y campañas)")
        raise HTTPException(
            status_code=400, 
            detail=f"Falta archivo obligatorio: {', '.join(missing)}. Debes subir ambos archivos para poder procesar la cartera."
        )
    
    df_master = DATA_STORE["df_original"].copy()
    
    if req.cartera_activa is not None:
        DATA_STORE["cartera_activa_manual"] = req.cartera_activa

    # 1. Resumen Global Fijo por Tramos (Invariable ante cualquier filtro operativo)
    resumen_global = calculate_global_tramos_summary(df_master, req.cartera_activa)
    matriz_campanas = calculate_campanas_tramos_matrix(df_master)
    
    # 2. Aplicar filtros a la cartera para la vista operativa detallada
    df = df_master.copy()
    df = df[(df["Dias de Mora"] >= req.min_dias) & (df["Dias de Mora"] <= req.max_dias)]
    
    if req.tipos_campana and len(req.tipos_campana) > 0 and "Tipo_Campana" in df.columns:
        df = df[df["Tipo_Campana"].isin(req.tipos_campana)]

    if req.programas and len(req.programas) > 0 and "Programa" in df.columns:
        df = df[df["Programa"].isin(req.programas)]
        
    if req.tipos and len(req.tipos) > 0 and "Tipo" in df.columns:
        df = df[df["Tipo"].isin(req.tipos)]

    if req.only_with_phone:
        has_tel = (df["Telefono_1"].fillna("") != "") | (df["Telefono_2"].fillna("") != "") | (df["Telefono_3"].fillna("") != "")
        df = df[has_tel]
    
    DATA_STORE["df_filtered"] = df
    DATA_STORE["filter_info"] = req.dict()
    
    build_debtor_matched_dataset(df)
    kpis = calculate_totals(df)
    
    # Vista previa: Primeras 100 filas
    cols_priority = [
        "Cedula Cliente", "Nombre Cliente", "Dias de Mora", "Tramo_Mora",
        "Pendiente Tarifa", "Pendiente Cuota", "Pendiente Mora", "Total",
        "Tipo_Campana", "Programa", "Telefono_1", "Telefono_2", "Correo", "Tipo"
    ]
    cols_to_show = [c for c in cols_priority if c in df.columns] + [c for c in df.columns if c not in cols_priority and not c.startswith("_") and c != "Cedula_Key"]
    
    preview_df = df[cols_to_show].head(100).fillna("")
    preview = []
    for r in preview_df.to_dict(orient="records"):
        tel1 = str(r.get("Telefono_1", ""))
        wa_link = f"https://wa.me/{tel1}" if tel1 and len(tel1) >= 10 else ""
        r["wa_link"] = wa_link
        preview.append(r)
    
    return {
        "kpis": kpis,
        "resumen_global_tramos": resumen_global,
        "matriz_campanas_tramos": matriz_campanas,
        "agrupacion": resumen_global["tramos"],
        "group_by_used": "Tramo Oficial de Mora",
        "preview_count": len(preview),
        "total_filtered": len(df),
        "preview": preview
    }

@app.post("/api/tramos/set-cartera-activa")
def set_cartera_activa_endpoint(req: SetCarteraActivaRequest):
    """Guarda o actualiza el monto manual de Cartera Activa y recalcula el resumen de tramos."""
    DATA_STORE["cartera_activa_manual"] = req.cartera_activa
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is not None and not active_df.empty:
        summary = calculate_global_tramos_summary(active_df, req.cartera_activa)
        return {"status": "ok", "cartera_activa": req.cartera_activa, "summary": summary}
    return {"status": "ok", "cartera_activa": req.cartera_activa, "summary": None}

@app.get("/api/download/excel")
def download_excel():
    """Descarga la cartera filtrada enriquecida con todas las columnas unificadas."""
    if DATA_STORE["df_filtered"] is None or DATA_STORE["df_filtered"].empty:
        raise HTTPException(status_code=400, detail="No hay datos filtrados para descargar.")
    
    df = DATA_STORE["df_filtered"].copy()
    
    # Ordenar columnas logicamente
    preferred_order = [
        "Cedula Cliente", "Nombre Cliente", "Dias de Mora", "Tramo_Mora",
        "Pendiente Tarifa", "Pendiente Cuota", "Pendiente Mora", "Total",
        "Tipo_Campana", "Programa", "Telefono_1", "Telefono_2", "Telefono_3", "Correo",
        "Tipo", "Identificador de credito", "Canal", "Fecha Vencimiento", "Cuota Numero"
    ]
    cols = [c for c in preferred_order if c in df.columns] + [c for c in df.columns if c not in preferred_order and not c.startswith("_") and c != "Cedula_Key"]
    df_clean = df[cols]
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_clean.to_excel(writer, index=False, sheet_name="Cartera Enriquecida")
    
    output.seek(0)
    filename = f"cartera_desglose_enriquecida_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/api/download/pdf")
def download_pdf():
    """Genera reporte ejecutivo formal en PDF con metricas financieras y de contactabilidad."""
    if DATA_STORE["df_filtered"] is None or DATA_STORE["df_filtered"].empty:
        raise HTTPException(status_code=400, detail="No hay datos filtrados para generar el reporte.")
    
    df = DATA_STORE["df_filtered"]
    kpis = calculate_totals(df)
    f_info = DATA_STORE.get("filter_info", {})
    
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#1e293b"),
        spaceAfter=6
    )
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Normal'],
        fontSize=9,
        leading=13,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=14
    )
    kpi_title = ParagraphStyle(
        'KPITitle',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor("#64748b"),
        alignment=1
    )
    kpi_val = ParagraphStyle(
        'KPIVal',
        parent=styles['Heading2'],
        fontSize=13,
        textColor=colors.HexColor("#0f172a"),
        alignment=1
    )
    table_cell = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor("#1e293b")
    )
    table_header = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.white,
        fontName="Helvetica-Bold"
    )
    
    story = []
    
    # Header
    story.append(Paragraph("REPORTE EJECUTIVO DE CARTERA VENCIDA Y CONTACTABILIDAD", title_style))
    sub_text = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Rango mora: {f_info.get('min_dias', 1)} a {f_info.get('max_dias', 60)} días | Archivo Cartera: {html.escape(str(DATA_STORE.get('cartera_filename', 'Cartera')))}"
    story.append(Paragraph(sub_text, subtitle_style))
    story.append(Spacer(1, 10))
    
    # KPI Grid Table 1: Financiero
    kpi_data = [
        [
            Paragraph("TOTAL CLIENTES", kpi_title),
            Paragraph("TOTAL CUOTAS EN MORA", kpi_title),
            Paragraph("MONTO TOTAL VENCIDO", kpi_title),
            Paragraph("PROMEDIO X CLIENTE", kpi_title)
        ],
        [
            Paragraph(f"{kpis['total_clientes']:,}", kpi_val),
            Paragraph(f"{kpis['total_cuotas']:,}", kpi_val),
            Paragraph(f"${kpis['monto_total']:,.2f} USD", kpi_val),
            Paragraph(f"${kpis['promedio_cliente']:,.2f} USD", kpi_val)
        ]
    ]
    t_kpi = Table(kpi_data, colWidths=[135, 135, 140, 130])
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#cbd5e1")),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 10))
    
    # KPI Grid Table 2: Contactabilidad
    contact_data = [
        [
            Paragraph("CLIENTES CON TELÉFONO", kpi_title),
            Paragraph("COBERTURA WHATSAPP", kpi_title),
            Paragraph("MONTO CONTACTABLE", kpi_title),
            Paragraph("% DEUDA CONTACTABLE", kpi_title)
        ],
        [
            Paragraph(f"{kpis['total_con_telefono']:,}", kpi_val),
            Paragraph(f"{kpis['pct_con_telefono']}%", kpi_val),
            Paragraph(f"${kpis['monto_contactable']:,.2f} USD", kpi_val),
            Paragraph(f"{kpis['pct_monto_contactable']}%", kpi_val)
        ]
    ]
    t_contact = Table(contact_data, colWidths=[135, 135, 140, 130])
    t_contact.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f0fdf4")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#bbf7d0")),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#bbf7d0")),
        ('TOPPADDING', (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
    ]))
    story.append(t_contact)
    story.append(Spacer(1, 15))
    
    # Subtotales
    story.append(Paragraph(f"<b>Composición de la Deuda:</b> Tarifa Pendiente: <b>${kpis['pendiente_tarifa']:,.2f}</b> | Cuota Pendiente: <b>${kpis['pendiente_cuota']:,.2f}</b> | Recargo Mora: <b>${kpis['pendiente_mora']:,.2f}</b>", styles['Normal']))
    story.append(Spacer(1, 15))
    
    # Distribución Fija por Tramos Oficiales
    df_master = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else df
    global_summary = calculate_global_tramos_summary(df_master)
    cutoff = global_summary.get("cutoff_date", datetime.now().strftime('%d-%m-%Y'))
    total_cartera_val = global_summary.get("total_cartera_usd", kpis['monto_total'])

    story.append(Paragraph(f"<b>Resumen Agrupado por Tramos de Mora (al {cutoff})</b>", styles['Heading3']))
    story.append(Paragraph(f"<b>Cartera Activa Total: ${total_cartera_val:,.2f} USD</b> | Cuotas: {global_summary.get('total_cuotas', 0):,} | Clientes: {global_summary.get('total_clientes', 0):,}", styles['Normal']))
    story.append(Spacer(1, 8))

    table_rows = [
        [
            Paragraph("Fase / Categoría", table_header),
            Paragraph("Rango / Etiqueta", table_header),
            Paragraph("Cant. Cuotas", table_header),
            Paragraph("Clientes", table_header),
            Paragraph("Monto (USD)", table_header),
            Paragraph("% Cartera", table_header)
        ]
    ]
    
    for t_data in global_summary.get("tramos", []):
        table_rows.append([
            Paragraph(html.escape(t_data["fase"]), table_cell),
            Paragraph(html.escape(t_data["etiqueta"]), table_cell),
            Paragraph(f"{t_data['cuotas']:,}", table_cell),
            Paragraph(f"{t_data['clientes']:,}", table_cell),
            Paragraph(f"${t_data['total_monto']:,.2f}", table_cell),
            Paragraph(f"{t_data['porcentaje']:.2f}%", table_cell)
        ])
    
    # Fila de Totales Generales
    table_rows.append([
        Paragraph("<b>Total Cartera</b>", table_cell),
        Paragraph("<b>Gestión total</b>", table_cell),
        Paragraph(f"<b>{global_summary.get('total_cuotas', 0):,}</b>", table_cell),
        Paragraph(f"<b>{global_summary.get('total_clientes', 0):,}</b>", table_cell),
        Paragraph(f"<b>${total_cartera_val:,.2f}</b>", table_cell),
        Paragraph("<b>100.0%</b>", table_cell)
    ])
    
    t_grp = Table(table_rows, colWidths=[150, 110, 70, 70, 95, 55])
    t_grp.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e293b")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_grp)
    
    doc.build(story)
    buf.seek(0)
    
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=reporte_cartera_ejecutivo.pdf"}
    )

@app.get("/api/tramos/download/excel")
def download_tramos_excel(cartera_activa: Optional[float] = None):
    """Descarga el Resumen Agrupado por Tramos de Mora en un archivo Excel formal."""
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is None or active_df.empty:
        raise HTTPException(status_code=400, detail="No hay cartera activa cargada.")
        
    if cartera_activa is not None and cartera_activa > 0:
        DATA_STORE["cartera_activa_manual"] = cartera_activa
    manual_ca = cartera_activa or DATA_STORE.get("cartera_activa_manual")
    global_summary = calculate_global_tramos_summary(active_df, manual_ca)
    cutoff = global_summary.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = global_summary.get("total_cartera_usd", 0.0)
    total_pct = global_summary.get("total_porcentaje")
    
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen por Tramos"
    ws.views.sheetView[0].showGridLines = True

    # Estilos elegantes CobranzApp
    font_title = Font(name="Calibri", size=14, bold=True, color="0F172A")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="64748B")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=11, color="1E293B")
    font_data_bold = Font(name="Calibri", size=11, bold=True, color="1E293B")
    font_total = Font(name="Calibri", size=11, bold=True, color="0F172A")

    fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_banner = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_ca_label = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    total_border = Border(
        left=Side(style='thin', color='94A3B8'),
        right=Side(style='thin', color='94A3B8'),
        top=Side(style='medium', color='1E293B'),
        bottom=Side(style='double', color='1E293B')
    )

    # Titulo
    ws["A1"] = "COBRANZAPP - RESUMEN AGRUPADO POR TRAMOS DE MORA"
    ws["A1"].font = font_title
    
    ws["A2"] = f"Fecha de corte: al {cutoff} | Tarifa Pendiente Total: ${total_usd:,.2f} USD | Emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = font_subtitle

    # Banner Ejecutivo
    ws.merge_cells("A4:B4")
    ws["A4"] = "TARIFA PENDIENTE TOTAL (USD)"
    ws["A4"].font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A4"].fill = fill_banner

    ws["C4"] = total_usd
    ws["C4"].font = Font(name="Calibri", size=14, bold=True, color="047857")
    ws["C4"].number_format = '$#,##0.00'
    ws["C4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C4"].fill = fill_banner

    ws["D4"] = "CARTERA ACTIVA (USD):"
    ws["D4"].font = Font(name="Calibri", size=10, bold=True, color="92400E")
    ws["D4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["D4"].fill = fill_ca_label

    ws.merge_cells("E4:F4")
    if manual_ca and manual_ca > 0:
        ws["E4"] = manual_ca
        ws["E4"].font = Font(name="Calibri", size=14, bold=True, color="B45309")
        ws["E4"].number_format = '$#,##0.00'
    else:
        ws["E4"] = "No especificada"
        ws["E4"].font = Font(name="Calibri", size=11, italic=True, color="64748B")
    ws["E4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["E4"].fill = fill_ca_label

    # Encabezados de Tabla en fila 6
    headers = ["Fase / Categoría", "Rango / Etiqueta", "Cant. Cuotas", "Clientes", "Monto Tarifa (USD)", "% Cartera Activa"]
    start_row = 6
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center" if col_idx in [2, 6] else ("right" if col_idx in [3, 4, 5] else "left"), vertical="center")
        cell.border = thin_border
    ws.row_dimensions[start_row].height = 24

    # Filas de datos
    curr_row = start_row + 1
    for idx, t in enumerate(global_summary.get("tramos", [])):
        r = curr_row + idx
        ws.row_dimensions[r].height = 20
        c1 = ws.cell(row=r, column=1, value=t["fase"])
        c2 = ws.cell(row=r, column=2, value=t["etiqueta"])
        c3 = ws.cell(row=r, column=3, value=t["cuotas"])
        c4 = ws.cell(row=r, column=4, value=t["clientes"])
        c5 = ws.cell(row=r, column=5, value=t["total_monto"])
        if t["porcentaje"] is not None:
            c6 = ws.cell(row=r, column=6, value=t["porcentaje"] / 100.0)
            c6.number_format = '0.00%'
        else:
            c6 = ws.cell(row=r, column=6, value="-")

        c1.font = font_data
        c2.font = font_data_bold
        c3.font = font_data
        c4.font = font_data
        c5.font = font_data_bold
        c6.font = font_data_bold

        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="right", vertical="center")
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c6.alignment = Alignment(horizontal="right" if t["porcentaje"] is not None else "center", vertical="center")

        c3.number_format = '#,##0'
        c4.number_format = '#,##0'
        c5.number_format = '$#,##0.00'

        row_fill = fill_alt if idx % 2 == 1 else None
        for col_idx in range(1, 7):
            cell = ws.cell(row=r, column=col_idx)
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill

    # Fila Total
    tot_row = curr_row + len(global_summary.get("tramos", []))
    ws.row_dimensions[tot_row].height = 24

    tc1 = ws.cell(row=tot_row, column=1, value="TOTAL CARTERA")
    tc2 = ws.cell(row=tot_row, column=2, value="Gestión total")
    tc3 = ws.cell(row=tot_row, column=3, value=global_summary.get("total_cuotas", 0))
    tc4 = ws.cell(row=tot_row, column=4, value=global_summary.get("total_clientes", 0))
    tc5 = ws.cell(row=tot_row, column=5, value=total_usd)
    if total_pct is not None:
        tc6 = ws.cell(row=tot_row, column=6, value=total_pct / 100.0)
        tc6.number_format = '0.00%'
    else:
        tc6 = ws.cell(row=tot_row, column=6, value="-")

    for cell in [tc1, tc2, tc3, tc4, tc5, tc6]:
        cell.font = font_total
        cell.fill = fill_total
        cell.border = total_border

    tc1.alignment = Alignment(horizontal="left", vertical="center")
    tc2.alignment = Alignment(horizontal="center", vertical="center")
    tc3.alignment = Alignment(horizontal="right", vertical="center")
    tc4.alignment = Alignment(horizontal="right", vertical="center")
    tc5.alignment = Alignment(horizontal="right", vertical="center")
    tc6.alignment = Alignment(horizontal="right" if total_pct is not None else "center", vertical="center")

    tc3.number_format = '#,##0'
    tc4.number_format = '#,##0'
    tc5.number_format = '$#,##0.00'

    # Auto ancho de columnas
    col_widths = {1: 34, 2: 20, 3: 16, 4: 16, 5: 22, 6: 15}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"resumen_tramos_mora_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/api/tramos/download/pdf")
def download_tramos_pdf(cartera_activa: Optional[float] = None):
    """Genera un reporte PDF ejecutivo exclusivamente del Resumen Agrupado por Tramos de Mora."""
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is None or active_df.empty:
        raise HTTPException(status_code=400, detail="No hay cartera activa cargada.")
        
    if cartera_activa is not None and cartera_activa > 0:
        DATA_STORE["cartera_activa_manual"] = cartera_activa
    manual_ca = cartera_activa or DATA_STORE.get("cartera_activa_manual")
    global_summary = calculate_global_tramos_summary(active_df, manual_ca)
    cutoff = global_summary.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = global_summary.get("total_cartera_usd", 0.0)
    total_cuotas = global_summary.get("total_cuotas", 0)
    total_clientes = global_summary.get("total_clientes", 0)
    total_pct = global_summary.get("total_porcentaje")
    promedio = total_usd / total_clientes if total_clientes > 0 else 0.0

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'TramosTitle',
        parent=styles['Heading1'],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold",
        spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        'TramosSubtitle',
        parent=styles['Normal'],
        fontSize=9,
        leading=13,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=14
    )
    banner_label = ParagraphStyle(
        'BannerLabel',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#64748b"),
        alignment=1,
        fontName="Helvetica-Bold"
    )
    banner_val = ParagraphStyle(
        'BannerVal',
        parent=styles['Heading2'],
        fontSize=14,
        leading=17,
        textColor=colors.HexColor("#1e293b"),
        alignment=1,
        fontName="Helvetica-Bold"
    )
    table_cell = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#1e293b")
    )
    table_cell_center = ParagraphStyle(
        'TableCellCenter',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#1e293b"),
        alignment=1
    )
    table_cell_right = ParagraphStyle(
        'TableCellRight',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#1e293b"),
        alignment=2
    )
    table_header = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#ffffff"),
        fontName="Helvetica-Bold"
    )
    table_header_right = ParagraphStyle(
        'TableHeaderRight',
        parent=table_header,
        alignment=2
    )
    table_header_center = ParagraphStyle(
        'TableHeaderCenter',
        parent=table_header,
        alignment=1
    )

    story = []

    # Encabezado
    story.append(Paragraph("COBRANZAPP | RESUMEN AGRUPADO POR TRAMOS DE MORA", title_style))
    story.append(Paragraph(f"<b>Fecha de corte:</b> al {cutoff} &nbsp;|&nbsp; <b>Fecha de emisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')} &nbsp;|&nbsp; <b>Cartera:</b> {html.escape(str(DATA_STORE.get('cartera_filename', 'Cartera Activa')))}", subtitle_style))
    story.append(Spacer(1, 8))

    # Tarjetas ejecutivas superiores
    if manual_ca and manual_ca > 0:
        kpi_cards = [
            [
                Paragraph("TARIFA PENDIENTE TOTAL", banner_label),
                Paragraph("CARTERA ACTIVA TOTAL", banner_label),
                Paragraph("% MORA SOBRE CARTERA", banner_label),
                Paragraph("TOTAL CLIENTES", banner_label)
            ],
            [
                Paragraph(f"${total_usd:,.2f} USD", banner_val),
                Paragraph(f"${manual_ca:,.2f} USD", banner_val),
                Paragraph(f"{total_pct:.2f}%" if total_pct is not None else "-", banner_val),
                Paragraph(f"{total_clientes:,}", banner_val)
            ]
        ]
    else:
        kpi_cards = [
            [
                Paragraph("TARIFA PENDIENTE TOTAL", banner_label),
                Paragraph("CARTERA ACTIVA TOTAL", banner_label),
                Paragraph("TOTAL CLIENTES", banner_label),
                Paragraph("CANTIDAD DE CUOTAS", banner_label)
            ],
            [
                Paragraph(f"${total_usd:,.2f} USD", banner_val),
                Paragraph("No especificada", banner_val),
                Paragraph(f"{total_clientes:,}", banner_val),
                Paragraph(f"{total_cuotas:,}", banner_val)
            ]
        ]

    t_kpis = Table(kpi_cards, colWidths=[135, 135, 135, 135])
    t_kpis.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#cbd5e1")),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_kpis)
    story.append(Spacer(1, 16))

    story.append(Paragraph("<b>Distribución de Tarifa Pendiente por los 5 Tramos Oficiales de Mora:</b>", styles['Normal']))
    story.append(Spacer(1, 6))

    # Tabla de Tramos Oficiales
    table_rows = [
        [
            Paragraph("Fase / Categoría", table_header),
            Paragraph("Rango / Etiqueta", table_header_center),
            Paragraph("Cant. Cuotas", table_header_right),
            Paragraph("Clientes", table_header_right),
            Paragraph("Monto Tarifa (USD)", table_header_right),
            Paragraph("% Cartera Activa", table_header_right)
        ]
    ]

    for t in global_summary.get("tramos", []):
        pct_display = f"<b>{t['porcentaje']:.2f}%</b>" if t["porcentaje"] is not None else "-"
        table_rows.append([
            Paragraph(f"<b>{html.escape(t['fase'])}</b>", table_cell),
            Paragraph(html.escape(t["etiqueta"]), table_cell_center),
            Paragraph(f"{t['cuotas']:,}", table_cell_right),
            Paragraph(f"{t['clientes']:,}", table_cell_right),
            Paragraph(f"${t['total_monto']:,.2f}", table_cell_right),
            Paragraph(pct_display, table_cell_right if t["porcentaje"] is not None else table_cell_center)
        ])

    # Fila Total Consolidada
    total_pct_display = f"<b>{total_pct:.2f}%</b>" if total_pct is not None else "-"
    table_rows.append([
        Paragraph("<b>TOTAL CARTERA</b>", table_cell),
        Paragraph("<b>Gestión total</b>", table_cell_center),
        Paragraph(f"<b>{total_cuotas:,}</b>", table_cell_right),
        Paragraph(f"<b>{total_clientes:,}</b>", table_cell_right),
        Paragraph(f"<b>${total_usd:,.2f}</b>", table_cell_right),
        Paragraph(total_pct_display, table_cell_right if total_pct is not None else table_cell_center)
    ])

    t_table = Table(table_rows, colWidths=[150, 95, 75, 75, 95, 50])
    t_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e293b")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_table)
    story.append(Spacer(1, 16))

    # Notas ejecutivas estratégicas
    tramos_list = global_summary.get("tramos", [])
    max_tramo = max(tramos_list, key=lambda x: x["total_monto"]) if tramos_list else None
    if max_tramo:
        pct_note = f" ({max_tramo['porcentaje']:.2f}% de la cartera activa total)" if max_tramo.get("porcentaje") is not None else ""
        notes_text = (
            f"<b>Observación Estratégica:</b> El tramo con mayor concentración de cartera es <b>{max_tramo['etiqueta']} ({max_tramo['fase']})</b> "
            f"con un saldo acumulado de <b>${max_tramo['total_monto']:,.2f} USD</b>{pct_note} "
            f"distribuido en {max_tramo['cuotas']:,} cuotas y {max_tramo['clientes']:,} clientes únicos. "
            f"Se sugiere enfocar las acciones intensivas de negociación y cobranza pre-jurídica en este segmento para maximizar la recuperación de liquidez."
        )
        p_notes = Paragraph(notes_text, styles['Normal'])
        t_notes = Table([[p_notes]], colWidths=[540])
        t_notes.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#eff6ff")),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#bfdbfe")),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(t_notes)

    doc.build(story)
    buf.seek(0)
    
    filename = f"resumen_tramos_mora_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/api/campanas-tramos/download/excel")
def download_campanas_tramos_excel():
    """Descarga la Matriz Completa de Campañas vs 5 Tramos Oficiales en Excel con formato corporativo."""
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is None or active_df.empty:
        raise HTTPException(status_code=400, detail="No hay cartera activa cargada.")

    matrix_data = calculate_campanas_tramos_matrix(active_df)
    cutoff = matrix_data.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = matrix_data.get("total_cartera_usd", 0.0)

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Desglose Tramos y Campañas"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=14, bold=True, color="0F172A")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="64748B")
    font_group_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_sub_hdr = Font(name="Calibri", size=9, bold=True, color="334155")
    font_data = Font(name="Calibri", size=10, color="1E293B")
    font_data_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_total = Font(name="Calibri", size=10, bold=True, color="0F172A")

    fill_tramo_hdrs = [
        PatternFill(start_color="065F46", end_color="065F46", fill_type="solid"), # 1-30d
        PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid"), # 31-60d
        PatternFill(start_color="92400E", end_color="92400E", fill_type="solid"), # 61-90d
        PatternFill(start_color="9A3412", end_color="9A3412", fill_type="solid"), # 91-180d
        PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid"), # 181+d
        PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")  # Total
    ]
    fill_sub_hdr = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    total_border = Border(
        left=Side(style='thin', color='94A3B8'),
        right=Side(style='thin', color='94A3B8'),
        top=Side(style='medium', color='0F172A'),
        bottom=Side(style='double', color='0F172A')
    )

    ws["A1"] = "COBRANZAPP - DESGLOSE POR TRAMOS Y CATEGORÍAS (CAMPAÑAS)"
    ws["A1"].font = font_title
    ws["A2"] = f"Fecha de corte: al {cutoff} | Tarifa Pendiente Total: ${total_usd:,.2f} USD | Emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = font_subtitle

    # Fila 4: Nivel 1 de Encabezados
    ws.merge_cells("A4:A5")
    ws["A4"] = "Campaña Comercial"
    ws["A4"].font = font_group_hdr
    ws["A4"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A4"].border = thin_border

    tramo_spans = [
        ("B4:C4", "B", "C", "1 a 30 días (Preventiva)", fill_tramo_hdrs[0]),
        ("D4:E4", "D", "E", "31 a 60 días (Administrativa)", fill_tramo_hdrs[1]),
        ("F4:G4", "F", "G", "61 a 90 días (Negociación)", fill_tramo_hdrs[2]),
        ("H4:I4", "H", "I", "91 a 180 días (Pre-jurídica)", fill_tramo_hdrs[3]),
        ("J4:K4", "J", "K", "181+ días (Castigo)", fill_tramo_hdrs[4]),
        ("L4:N4", "L", "N", "Total Tarifa Consolidada", fill_tramo_hdrs[5]),
    ]

    for span, col_start, col_end, title, fill in tramo_spans:
        ws.merge_cells(span)
        cell = ws[f"{col_start}4"]
        cell.value = title
        cell.font = font_group_hdr
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Fila 5: Nivel 2 de Encabezados
    sub_headers = [
        ("B5", "Tarifa ($)", "right"), ("C5", "Clientes", "right"),
        ("D5", "Tarifa ($)", "right"), ("E5", "Clientes", "right"),
        ("F5", "Tarifa ($)", "right"), ("G5", "Clientes", "right"),
        ("H5", "Tarifa ($)", "right"), ("I5", "Clientes", "right"),
        ("J5", "Tarifa ($)", "right"), ("K5", "Clientes", "right"),
        ("L5", "Total Tarifa ($)", "right"), ("M5", "Total Clientes", "right"),
        ("N5", "% Tarifa", "right")
    ]
    for cell_id, text, align in sub_headers:
        c = ws[cell_id]
        c.value = text
        c.font = font_sub_hdr
        c.fill = fill_sub_hdr
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border = thin_border

    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 20

    curr_row = 6
    for idx, row in enumerate(matrix_data.get("matrix", [])):
        r = curr_row + idx
        ws.row_dimensions[r].height = 19
        
        c_camp = ws.cell(row=r, column=1, value=row["campana"])
        c_camp.font = font_data_bold
        c_camp.alignment = Alignment(horizontal="left", vertical="center")

        col_idx = 2
        for t_cfg in OFFICIAL_TRAMOS_CONFIG:
            t_data = row["tramos"].get(t_cfg["id"], {"monto": 0.0, "clientes": 0})
            cm = ws.cell(row=r, column=col_idx, value=t_data["monto"])
            cc = ws.cell(row=r, column=col_idx + 1, value=t_data["clientes"])

            cm.font = font_data
            cc.font = font_data
            cm.alignment = Alignment(horizontal="right", vertical="center")
            cc.alignment = Alignment(horizontal="right", vertical="center")
            cm.number_format = '$#,##0.00' if t_data["monto"] > 0 else '"-"'
            cc.number_format = '#,##0' if t_data["clientes"] > 0 else '"-"'
            col_idx += 2

        c_tot_m = ws.cell(row=r, column=12, value=row["total_monto"])
        c_tot_c = ws.cell(row=r, column=13, value=row["total_clientes"])
        c_tot_p = ws.cell(row=r, column=14, value=row["porcentaje"] / 100.0)

        c_tot_m.font = font_data_bold
        c_tot_c.font = font_data_bold
        c_tot_p.font = font_data_bold

        c_tot_m.alignment = Alignment(horizontal="right", vertical="center")
        c_tot_c.alignment = Alignment(horizontal="right", vertical="center")
        c_tot_p.alignment = Alignment(horizontal="right", vertical="center")

        c_tot_m.number_format = '$#,##0.00'
        c_tot_c.number_format = '#,##0'
        c_tot_p.number_format = '0.00%'

        row_fill = fill_alt if idx % 2 == 1 else None
        for c_i in range(1, 15):
            cell = ws.cell(row=r, column=c_i)
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill

    # Fila TOTAL CONSOLIDADO
    tot_row = curr_row + len(matrix_data.get("matrix", []))
    ws.row_dimensions[tot_row].height = 24

    tc_camp = ws.cell(row=tot_row, column=1, value="TOTAL CONSOLIDADO")
    tc_camp.font = font_total
    tc_camp.alignment = Alignment(horizontal="left", vertical="center")

    totals_tramos = matrix_data.get("totals_by_tramo", {})
    col_idx = 2
    for t_cfg in OFFICIAL_TRAMOS_CONFIG:
        t_data = totals_tramos.get(t_cfg["id"], {"monto": 0.0, "clientes": 0})
        cm = ws.cell(row=tot_row, column=col_idx, value=t_data["monto"])
        cc = ws.cell(row=tot_row, column=col_idx + 1, value=t_data["clientes"])

        cm.font = font_total
        cc.font = font_total
        cm.alignment = Alignment(horizontal="right", vertical="center")
        cc.alignment = Alignment(horizontal="right", vertical="center")
        cm.number_format = '$#,##0.00'
        cc.number_format = '#,##0'
        col_idx += 2

    tc_tot_m = ws.cell(row=tot_row, column=12, value=total_usd)
    tc_tot_c = ws.cell(row=tot_row, column=13, value=matrix_data.get("total_clientes", 0))
    tc_tot_p = ws.cell(row=tot_row, column=14, value=1.0)

    for c in [tc_tot_m, tc_tot_c, tc_tot_p]:
        c.font = font_total
        c.alignment = Alignment(horizontal="right", vertical="center")

    tc_tot_m.number_format = '$#,##0.00'
    tc_tot_c.number_format = '#,##0'
    tc_tot_p.number_format = '0.00%'

    for c_i in range(1, 15):
        cell = ws.cell(row=tot_row, column=c_i)
        cell.fill = fill_total
        cell.border = total_border

    widths = {
        1: 32, 2: 15, 3: 11, 4: 15, 5: 11, 6: 15, 7: 11,
        8: 15, 9: 11, 10: 15, 11: 11, 12: 17, 13: 13, 14: 12
    }
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"desglose_campanas_tramos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/api/campanas-tramos/download/pdf")
def download_campanas_tramos_pdf():
    """Genera reporte ejecutivo en PDF en formato apaisado (Landscape) con la Matriz de Campañas vs 5 Tramos."""
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is None or active_df.empty:
        raise HTTPException(status_code=400, detail="No hay cartera activa cargada.")

    matrix_data = calculate_campanas_tramos_matrix(active_df)
    cutoff = matrix_data.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = matrix_data.get("total_cartera_usd", 0.0)
    total_clientes = matrix_data.get("total_clientes", 0)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'MatTitle',
        parent=styles['Heading1'],
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold",
        spaceAfter=2
    )
    subtitle_style = ParagraphStyle(
        'MatSubtitle',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=10
    )
    hdr_group = ParagraphStyle(
        'HdrGroup',
        parent=styles['Normal'],
        fontSize=7.5,
        leading=9,
        textColor=colors.white,
        fontName="Helvetica-Bold",
        alignment=1
    )
    hdr_sub = ParagraphStyle(
        'HdrSub',
        parent=styles['Normal'],
        fontSize=7,
        leading=8,
        textColor=colors.HexColor("#334155"),
        fontName="Helvetica-Bold",
        alignment=2
    )
    cell_camp = ParagraphStyle(
        'CellCamp',
        parent=styles['Normal'],
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold"
    )
    cell_val = ParagraphStyle(
        'CellVal',
        parent=styles['Normal'],
        fontSize=7,
        leading=8.5,
        textColor=colors.HexColor("#1e293b"),
        alignment=2
    )
    cell_tot = ParagraphStyle(
        'CellTot',
        parent=styles['Normal'],
        fontSize=7,
        leading=8.5,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold",
        alignment=2
    )

    story = []
    story.append(Paragraph("COBRANZAPP | DESGLOSE POR TRAMOS Y CATEGORÍAS (CAMPAÑAS)", title_style))
    story.append(Paragraph(
        f"<b>Fecha de corte:</b> al {cutoff} &nbsp;|&nbsp; "
        f"<b>Tarifa Pendiente Total:</b> ${total_usd:,.2f} USD &nbsp;|&nbsp; "
        f"<b>Clientes con Tarifa:</b> {total_clientes:,} &nbsp;|&nbsp; "
        f"<b>Emisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        subtitle_style
    ))

    table_data = [
        [
            Paragraph("<b>Campaña Comercial</b>", hdr_group),
            Paragraph("<b>1 a 30 días</b>", hdr_group), "",
            Paragraph("<b>31 a 60 días</b>", hdr_group), "",
            Paragraph("<b>61 a 90 días</b>", hdr_group), "",
            Paragraph("<b>91 a 180 días</b>", hdr_group), "",
            Paragraph("<b>181+ días</b>", hdr_group), "",
            Paragraph("<b>Total Tarifa</b>", hdr_group), "", ""
        ],
        [
            "",
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("% Tarifa", hdr_sub)
        ]
    ]

    for row in matrix_data.get("matrix", []):
        row_cells = [Paragraph(html.escape(row["campana"]), cell_camp)]
        for t_cfg in OFFICIAL_TRAMOS_CONFIG:
            t_data = row["tramos"].get(t_cfg["id"], {"monto": 0.0, "clientes": 0})
            m_str = f"${t_data['monto']:,.2f}" if t_data["monto"] > 0 else "-"
            c_str = f"{t_data['clientes']:,}" if t_data["clientes"] > 0 else "-"
            row_cells.append(Paragraph(m_str, cell_val))
            row_cells.append(Paragraph(c_str, cell_val))
        row_cells.append(Paragraph(f"${row['total_monto']:,.2f}", cell_tot))
        row_cells.append(Paragraph(f"{row['total_clientes']:,}", cell_tot))
        row_cells.append(Paragraph(f"{row['porcentaje']:.2f}%", cell_tot))
        table_data.append(row_cells)

    totals_tramos = matrix_data.get("totals_by_tramo", {})
    total_row_cells = [Paragraph("<b>TOTAL CONSOLIDADO</b>", cell_camp)]
    for t_cfg in OFFICIAL_TRAMOS_CONFIG:
        t_data = totals_tramos.get(t_cfg["id"], {"monto": 0.0, "clientes": 0})
        total_row_cells.append(Paragraph(f"<b>${t_data['monto']:,.2f}</b>", cell_tot))
        total_row_cells.append(Paragraph(f"<b>{t_data['clientes']:,}</b>", cell_tot))
    total_row_cells.append(Paragraph(f"<b>${total_usd:,.2f}</b>", cell_tot))
    total_row_cells.append(Paragraph(f"<b>{total_clientes:,}</b>", cell_tot))
    total_row_cells.append(Paragraph("<b>100.0%</b>", cell_tot))
    table_data.append(total_row_cells)

    col_widths = [140, 48, 28, 48, 28, 48, 28, 52, 28, 48, 28, 56, 32, 38]
    t_mat = Table(table_data, colWidths=col_widths, repeatRows=2)
    t_mat.setStyle(TableStyle([
        ('SPAN', (0,0), (0,1)),
        ('SPAN', (1,0), (2,0)),
        ('SPAN', (3,0), (4,0)),
        ('SPAN', (5,0), (6,0)),
        ('SPAN', (7,0), (8,0)),
        ('SPAN', (9,0), (10,0)),
        ('SPAN', (11,0), (13,0)),

        ('BACKGROUND', (0,0), (0,1), colors.HexColor("#0f172a")),
        ('BACKGROUND', (1,0), (2,0), colors.HexColor("#065f46")),
        ('BACKGROUND', (3,0), (4,0), colors.HexColor("#1e40af")),
        ('BACKGROUND', (5,0), (6,0), colors.HexColor("#92400e")),
        ('BACKGROUND', (7,0), (8,0), colors.HexColor("#9a3412")),
        ('BACKGROUND', (9,0), (10,0), colors.HexColor("#991b1b")),
        ('BACKGROUND', (11,0), (13,0), colors.HexColor("#0f172a")),
        ('BACKGROUND', (1,1), (-1,1), colors.HexColor("#f1f5f9")),

        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0,2), (-1,-2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#e2e8f0")),

        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t_mat)

    doc.build(story)
    buf.seek(0)

    filename = f"desglose_campanas_tramos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# -------------------------------------------------------------
# RUTAS DE CALL CENTER / DIALER / WHATSAPP
# -------------------------------------------------------------

@app.get("/api/contacts/status")
def get_contacts_status():
    contacts_sample_available = os.path.exists(DEFAULT_CONTACTS_PATH)
    if DATA_STORE["df_contacts_raw"] is None:
        return {
            "loaded": False,
            "sample_available": contacts_sample_available,
            "sample_name": "Base de Datos Correos y Teléfonos al 06-08 (2).xlsx"
        }
    return {
        "loaded": True,
        "filename": DATA_STORE["contacts_filename"],
        "total_contacts": len(DATA_STORE["df_contacts_raw"]),
        "sample_available": contacts_sample_available
    }

@app.post("/api/contacts/load-sample")
def load_sample_contacts():
    if not os.path.exists(DEFAULT_CONTACTS_PATH):
        raise HTTPException(status_code=404, detail="Archivo de contactos de ejemplo no encontrado en el servidor.")
        
    try:
        data = safe_read_bytes(DEFAULT_CONTACTS_PATH)
        df = parse_bytes_to_dataframe(data, "Base de Datos Correos y Teléfonos al 06-08 (2).xlsx")
        df = standardize_contacts_dataframe(df)
        DATA_STORE["df_contacts_raw"] = df
        DATA_STORE["contacts_filename"] = "Base de Contactos (Ejemplo Servidor)"
        
        enrich_and_merge_datasets()
        
        active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else df
        return {
            "message": f"Base de contactos cargada: {len(df):,} clientes listos para marcación y WhatsApp.",
            "total_contacts": len(df),
            "filename": DATA_STORE["contacts_filename"],
            "metadata": get_metadata(active_df),
            "totals": calculate_totals(active_df)
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al cargar base de contactos: {str(e)}")

@app.post("/api/contacts/cross-match")
def cross_match_contacts():
    """Retorna el listado consolidado a nivel deudor con teléfonos y enlaces directos de WhatsApp."""
    if DATA_STORE["df_filtered"] is None:
        raise HTTPException(status_code=400, detail="Primero debes cargar la cartera.")
        
    debtors = DATA_STORE.get("df_matched")
    if debtors is None or debtors.empty:
        build_debtor_matched_dataset(DATA_STORE["df_filtered"])
        debtors = DATA_STORE.get("df_matched")
        
    if debtors is None or debtors.empty:
        return {"kpis": {}, "preview": []}
        
    has_tel = (debtors["Telefono_1"] != "") | (debtors["Telefono_2"] != "") | (debtors["Telefono_3"] != "")
    has_mail = debtors["Correo"] != ""
    has_any = has_tel | has_mail
    has_none = ~has_any
    
    total_deudores = len(debtors)
    total_con_tel = int(has_tel.sum())
    total_con_mail = int(has_mail.sum())
    total_con_any = int(has_any.sum())
    total_sin_contacto = int(has_none.sum())
    
    monto_total_cartera = float(debtors["Monto_Total_USD"].sum())
    monto_contactable = float(debtors.loc[has_any, "Monto_Total_USD"].sum())
    monto_sin_contacto = float(debtors.loc[has_none, "Monto_Total_USD"].sum())
    
    pct_cobertura = round((total_con_any / total_deudores * 100), 1) if total_deudores > 0 else 0
    pct_monto_contactable = round((monto_contactable / monto_total_cartera * 100), 1) if monto_total_cartera > 0 else 0
    
    preview = []
    for _, r in debtors.head(100).iterrows():
        tel_principal = r["Telefono_1"] or r["Telefono_2"] or r["Telefono_3"]
        wa_link = f"https://wa.me/{tel_principal}" if tel_principal and len(tel_principal) >= 10 else ""
        preview.append({
            "cedula": r["Cedula"],
            "nombre": r["Nombre_Cliente"],
            "cuotas": int(r["Cuotas_Vencidas"]),
            "monto_total": float(r["Monto_Total_USD"]),
            "tarifa": float(r["Pendiente_Tarifa"]),
            "cuota": float(r["Pendiente_Cuota"]),
            "mora": float(r["Pendiente_Mora"]),
            "dias_mora": int(r["Max_Dias_Mora"]),
            "tramo_mora": r["Tramo_Mora"],
            "tipo_campana": r["Tipo_Campana"],
            "programa": r["Programa"],
            "telefono_1": r["Telefono_1"],
            "telefono_2": r["Telefono_2"],
            "telefono_3": r["Telefono_3"],
            "correo": r["Correo"],
            "wa_link": wa_link
        })
        
    return {
        "kpis": {
            "total_deudores": total_deudores,
            "total_con_telefono": total_con_tel,
            "total_con_correo": total_con_mail,
            "total_con_contacto": total_con_any,
            "total_sin_contacto": total_sin_contacto,
            "porcentaje_cobertura": pct_cobertura,
            "monto_total": round(monto_total_cartera, 2),
            "monto_contactable": round(monto_contactable, 2),
            "monto_sin_contacto": round(monto_sin_contacto, 2),
            "porcentaje_monto_contactable": pct_monto_contactable
        },
        "preview": preview
    }

@app.get("/api/contacts/download/excel")
def download_contacts_excel():
    """Descarga el listado consolidado a nivel deudor para marcadores predictivos y WhatsApp."""
    debtors = DATA_STORE.get("df_matched")
    if debtors is None or debtors.empty:
        if DATA_STORE["df_filtered"] is not None:
            build_debtor_matched_dataset(DATA_STORE["df_filtered"])
            debtors = DATA_STORE.get("df_matched")
            
    if debtors is None or debtors.empty:
        raise HTTPException(status_code=400, detail="No hay datos de cartera para exportar a Call Center.")
        
    df = debtors.copy()
    if "Cedula_Key" in df.columns:
        df = df.drop(columns=["Cedula_Key"])
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Deudores Call Center")
        
    output.seek(0)
    filename = f"cartera_deudores_callcenter_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/health")
def health_check():
    return {"status": "ok", "service": "cobranzapp"}

@app.get("/", response_class=HTMLResponse)
def get_dashboard():
    with open(os.path.join(os.path.dirname(__file__), "templates", "index.html"), "r", encoding="utf-8") as f:
        return f.read()

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")
    uvicorn.run("app:app", host=host, port=port)
