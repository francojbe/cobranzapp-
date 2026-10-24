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

from reportlab.lib.pagesizes import letter
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
        return "1 a 30 días (Preventiva)"
    elif dias <= 60:
        return "31 a 60 días (Administrativa)"
    elif dias <= 90:
        return "61 a 90 días (Reestructuración)"
    elif dias <= 180:
        return "91 a 180 días (Pre-jurídica)"
    else:
        return "181+ días (Castigo / Pesada)"

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
        
    # Sistema de seleccion inteligente de hoja
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
    "df_matched": None        # Agregada a nivel deudor para Call Center
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
        
    debtors = df.groupby("Cedula_Key").agg(
        Cedula=("Cedula Cliente", "first"),
        Nombre_Cliente=("Nombre Cliente", "first"),
        Cuotas_Vencidas=("Cedula Cliente", "count"),
        Monto_Total_USD=("_Fila_Total", "sum"),
        Pendiente_Tarifa=("Pendiente Tarifa", "sum"),
        Pendiente_Cuota=("Pendiente Cuota", "sum"),
        Pendiente_Mora=("Pendiente Mora", "sum"),
        Max_Dias_Mora=("Dias de Mora", "max"),
        Tramo_Mora=("Dias de Mora", lambda x: assign_tramo_mora(int(x.max()))),
        Tipo_Campana=("Tipo_Campana", lambda x: next((str(v) for v in x if str(v) not in ["", "nan", "None"]), "Sin Campaña Asignada")),
        Programa=("Programa", lambda x: next((str(v) for v in x if str(v) not in ["", "nan", "None"]), "General")),
        Tipo=("Tipo", lambda x: "/".join(sorted(set(str(v) for v in x if str(v) not in ["", "nan", "None", "General"]))) or "General"),
        Telefono_1=("Telefono_1", lambda x: next((str(v) for v in x if str(v).strip()), "")),
        Telefono_2=("Telefono_2", lambda x: next((str(v) for v in x if str(v).strip()), "")),
        Telefono_3=("Telefono_3", lambda x: next((str(v) for v in x if str(v).strip()), "")),
        Correo=("Correo", lambda x: next((str(v) for v in x if str(v).strip()), ""))
    ).reset_index()

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

def get_metadata(df: pd.DataFrame):
    if df is None or df.empty:
        return {
            "total_rows": 0, "total_clientes": 0, "programas": [],
            "tipos_campana": [], "tipos": [], "canales": [], "tramos": {},
            "min_dias": 0, "max_dias": 0, "has_contacts": False,
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
    
    has_contacts = (DATA_STORE["df_contacts_raw"] is not None and not DATA_STORE["df_contacts_raw"].empty)

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
            "sample_available": sample_available,
            "contacts_sample_available": contacts_sample_available
        }
    
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    meta = get_metadata(active_df)
    totals = calculate_totals(active_df)
    
    return {
        "loaded": cartera_loaded,
        "contacts_loaded": contacts_loaded,
        "cartera_filename": DATA_STORE["cartera_filename"],
        "contacts_filename": DATA_STORE["contacts_filename"],
        "metadata": meta,
        "totals": totals,
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
    return {
        "message": f"Cartera cargada exitosamente: {len(df_cartera):,} cuotas{contacts_msg}",
        "metadata": get_metadata(active_df),
        "totals": calculate_totals(active_df)
    }

@app.post("/api/upload")
async def upload_cartera(file: UploadFile = File(...)):
    """Sube el archivo de cartera (CSV o Excel). Si ya habia contactos cargados, cruza inmediatamente."""
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
        
    try:
        df = parse_bytes_to_dataframe(contents, file.filename)
        df = standardize_cartera_dataframe(df)
        
        DATA_STORE["df_cartera_raw"] = df
        DATA_STORE["cartera_filename"] = file.filename
        
        # Cruzar con contactos si ya estaban cargados
        enrich_and_merge_datasets()
        
        active_df = DATA_STORE["df_original"]
        return {
            "message": f"Archivo de cartera '{file.filename}' procesado con éxito.",
            "metadata": get_metadata(active_df),
            "totals": calculate_totals(active_df)
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error al procesar archivo de cartera: {str(e)}")

@app.post("/api/contacts/upload")
async def upload_contacts(file: UploadFile = File(...)):
    """Sube la base de contactos y campañas. Si ya habia cartera cargada, cruza inmediatamente."""
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
        
    try:
        df = parse_bytes_to_dataframe(contents, file.filename)
        df = standardize_contacts_dataframe(df)
        
        DATA_STORE["df_contacts_raw"] = df
        DATA_STORE["contacts_filename"] = file.filename
        
        # Cruzar inmediatamente con cartera
        enrich_and_merge_datasets()
        
        active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
        totals = calculate_totals(active_df) if active_df is not None else {}
        meta = get_metadata(active_df) if active_df is not None else {}
        
        return {
            "message": f"Base de contactos '{file.filename}' cargada: {len(df):,} clientes indexados.",
            "total_contacts": len(df),
            "filename": file.filename,
            "metadata": meta,
            "totals": totals
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error al procesar archivo de contactos: {str(e)}")

@app.post("/api/process")
def process_data(req: FilterRequest):
    """Aplica filtros dinamicos de tramos, campanas, tipos y agrupa los resultados."""
    if DATA_STORE["df_original"] is None:
        raise HTTPException(status_code=400, detail="No se ha cargado ningún archivo de cartera.")
    
    df = DATA_STORE["df_original"].copy()
    
    # 1. Filtro por dias de mora
    df = df[(df["Dias de Mora"] >= req.min_dias) & (df["Dias de Mora"] <= req.max_dias)]
    
    # 2. Filtro por Tipo de Campaña
    if req.tipos_campana and len(req.tipos_campana) > 0 and "Tipo_Campana" in df.columns:
        df = df[df["Tipo_Campana"].isin(req.tipos_campana)]

    # 3. Filtro por Programas
    if req.programas and len(req.programas) > 0 and "Programa" in df.columns:
        df = df[df["Programa"].isin(req.programas)]
        
    # 4. Filtro por Tipos
    if req.tipos and len(req.tipos) > 0 and "Tipo" in df.columns:
        df = df[df["Tipo"].isin(req.tipos)]

    # 5. Filtro opcional solo con telefono
    if req.only_with_phone:
        has_tel = (df["Telefono_1"].fillna("") != "") | (df["Telefono_2"].fillna("") != "") | (df["Telefono_3"].fillna("") != "")
        df = df[has_tel]
    
    DATA_STORE["df_filtered"] = df
    DATA_STORE["filter_info"] = req.dict()
    
    # Actualizar conjunto para Call Center
    build_debtor_matched_dataset(df)
    
    kpis = calculate_totals(df)
    
    # Calculo de agrupacion dinamica
    group_col = req.group_by if req.group_by in df.columns else "Tipo_Campana"
    if group_col not in df.columns:
        group_col = "Programa" if "Programa" in df.columns else "Tramo_Mora"
        
    agrupacion = []
    if not df.empty and group_col in df.columns:
        if (df["Pendiente Tarifa"].sum() + df["Pendiente Cuota"].sum() + df["Pendiente Mora"].sum()) > 0:
            df["_Total_Row"] = df["Pendiente Tarifa"] + df["Pendiente Cuota"] + df["Pendiente Mora"]
        elif "Total" in df.columns:
            df["_Total_Row"] = df["Total"]
        else:
            df["_Total_Row"] = 0.0

        grouped = df.groupby(group_col).agg(
            cuotas=("Cedula Cliente", "count"),
            clientes=("Cedula Cliente", "nunique"),
            p_tarifa=("Pendiente Tarifa", "sum"),
            p_cuota=("Pendiente Cuota", "sum"),
            p_mora=("Pendiente Mora", "sum"),
            total=("_Total_Row", "sum")
        ).reset_index()
        
        total_global = kpis["monto_total"]
        
        for _, r in grouped.iterrows():
            monto_grp = round(float(r["total"]), 2)
            pct = round((monto_grp / total_global * 100), 1) if total_global > 0 else 0
            agrupacion.append({
                "grupo": str(r[group_col]),
                "cuotas": int(r["cuotas"]),
                "clientes": int(r["clientes"]),
                "pendiente_tarifa": round(float(r["p_tarifa"]), 2),
                "pendiente_cuota": round(float(r["p_cuota"]), 2),
                "pendiente_mora": round(float(r["p_mora"]), 2),
                "total_monto": monto_grp,
                "porcentaje": pct
            })
        
        agrupacion = sorted(agrupacion, key=lambda x: x["total_monto"], reverse=True)
    
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
        "agrupacion": agrupacion,
        "group_by_used": group_col,
        "preview_count": len(preview),
        "total_filtered": len(df),
        "preview": preview
    }

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
    
    # Agrupacion
    grp_col = f_info.get('group_by', 'Tipo_Campana')
    if grp_col not in df.columns:
        grp_col = 'Programa' if 'Programa' in df.columns else 'Tramo_Mora'
        
    story.append(Paragraph(f"<b>Distribución de Deuda Agrupada por: {html.escape(grp_col)}</b>", styles['Heading3']))
    story.append(Spacer(1, 8))
    
    df["_Tot"] = df["Pendiente Tarifa"] + df["Pendiente Cuota"] + df["Pendiente Mora"]
    if df["_Tot"].sum() == 0 and "Total" in df.columns:
        df["_Tot"] = df["Total"]

    grp = df.groupby(grp_col).agg(
        cuotas=("Cedula Cliente", "count"),
        clientes=("Cedula Cliente", "nunique"),
        total=("_Tot", "sum")
    ).reset_index().sort_values(by="total", ascending=False)
    
    tot_g = kpis["monto_total"]
    table_rows = [
        [
            Paragraph(f"Grupo ({grp_col})", table_header),
            Paragraph("Clientes", table_header),
            Paragraph("Cuotas", table_header),
            Paragraph("Monto Vencido", table_header),
            Paragraph("% Cartera", table_header)
        ]
    ]
    
    for _, r in grp.head(25).iterrows():
        m_val = float(r["total"])
        pct = (m_val / tot_g * 100) if tot_g > 0 else 0
        table_rows.append([
            Paragraph(html.escape(str(r[grp_col])), table_cell),
            Paragraph(f"{int(r['clientes']):,}", table_cell),
            Paragraph(f"{int(r['cuotas']):,}", table_cell),
            Paragraph(f"${m_val:,.2f}", table_cell),
            Paragraph(f"{pct:.1f}%", table_cell)
        ])
    
    t_grp = Table(table_rows, colWidths=[200, 75, 75, 110, 80])
    t_grp.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2563eb")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
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
