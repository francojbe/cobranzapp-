import html
import io
import os
import re
import unicodedata
from typing import List, Optional
from datetime import datetime
import pandas as pd
from fastapi import FastAPI, File, UploadFile, HTTPException
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
        # e.g. "1234,56" o "1,234"
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

# Diccionario de patrones para mapear columnas de forma no ambigua
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
    ],
    "Telefono": [
        r"^telefono$", r"^telf$", r"^celular$", r"^telefonos$", r"^movil$"
    ],
    "Correo": [
        r"^correo$", r"^email$", r"^correos$", r"^mail$"
    ]
}

def map_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Mapea columnas inteligentemente garantizando que NUNCA se generen duplicados."""
    # 1. Deduplicar nombres originales si ya venian repetidos en el archivo
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

    # 2. Asignar maximo una columna del archivo por cada columna estandar
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

    df = df.rename(columns=matched)
    return df

def standardize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Valida, estandariza y asegura tipos numericos para todo el flujo."""
    df = map_column_names(df)
    
    # 1. Validar columnas minimas obligatorias
    req_cols = ["Dias de Mora", "Cedula Cliente"]
    missing = [c for c in req_cols if c not in df.columns]
    if missing:
        cols_found = [str(c) for c in df.columns.tolist()][:15]
        raise HTTPException(
            status_code=400, 
            detail=f"Faltan columnas clave obligatorias: {missing}. El archivo tiene estas columnas: {cols_found}"
        )
    
    # 2. Dias de Mora como enteros limpios
    df["Dias de Mora"] = df["Dias de Mora"].apply(parse_number_value).fillna(0).astype(int)
    
    # 3. Columnas de montos
    for c in ["Pendiente Tarifa", "Pendiente Cuota", "Pendiente Mora", "Total"]:
        if c in df.columns:
            df[c] = df[c].apply(parse_number_value).fillna(0.0)
        else:
            df[c] = 0.0
            
    # Manejar caso de archivo con Total pero sin desglose secundario
    tarifa_sum = float(df["Pendiente Tarifa"].sum())
    cuota_sum = float(df["Pendiente Cuota"].sum())
    mora_sum = float(df["Pendiente Mora"].sum())
    total_sum = float(df["Total"].sum())
    
    if (tarifa_sum + cuota_sum + mora_sum) == 0.0 and total_sum > 0.0:
        df["Pendiente Tarifa"] = df["Total"]
    elif total_sum == 0.0 and (tarifa_sum + cuota_sum + mora_sum) > 0.0:
        df["Total"] = df["Pendiente Tarifa"] + df["Pendiente Cuota"] + df["Pendiente Mora"]

    # 4. Normalizar Programa
    if "Programa" not in df.columns:
        df["Programa"] = "General"
    else:
        df["Programa"] = df["Programa"].fillna("General").astype(str).str.strip()
        df["Programa"] = df["Programa"].replace({"": "General", "nan": "General", "None": "General"})
        df["Programa"] = df["Programa"].str.replace(r"CAMPA.*AS", "CAMPAÑAS", regex=True)

    # 5. Normalizar Tipo
    if "Tipo" not in df.columns:
        df["Tipo"] = "General"
    else:
        df["Tipo"] = df["Tipo"].fillna("General").astype(str).str.strip()
        df["Tipo"] = df["Tipo"].replace({"": "General", "nan": "General", "None": "General"})

    # 6. Normalizar Cedula Cliente
    df["Cedula Cliente"] = df["Cedula Cliente"].fillna("").astype(str).str.strip()

    # 7. Normalizar Nombre Cliente
    if "Nombre Cliente" not in df.columns:
        df["Nombre Cliente"] = "Cliente Sin Nombre"
    else:
        df["Nombre Cliente"] = df["Nombre Cliente"].fillna("").astype(str).str.strip()

    return df

# -------------------------------------------------------------
# FASTAPI APP & STATE
# -------------------------------------------------------------

app = FastAPI(title="Motor de Reestructuracion y Segmentacion de Cartera")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DATA_STORE = {
    "df_original": None,
    "df_filtered": None,
    "filter_info": {},
    "filename": "",
    "df_contacts": None,
    "df_matched": None,
    "contacts_filename": ""
}

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

def standardize_contacts_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Estandariza columnas y cédula del archivo de contactos."""
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

    df = df.rename(columns=col_map)
    if "Cedula" not in df.columns:
        raise HTTPException(status_code=400, detail="El archivo de contactos debe contener una columna de Cédula o Identificación.")

    for col in ["Telefono_1", "Telefono_2", "Telefono_3", "Correo", "Nombre_Contacto"]:
        if col not in df.columns:
            df[col] = ""
        else:
            if "Telefono" in col:
                df[col] = df[col].apply(clean_phone_str)
            else:
                df[col] = df[col].fillna("").astype(str).str.strip()

    df["Cedula"] = df["Cedula"].fillna("").astype(str).str.strip().str.upper()
    df["Cedula_Key"] = df["Cedula"]
    return df.drop_duplicates(subset=["Cedula_Key"])

DEFAULT_CSV_PATH = os.path.join(os.path.dirname(__file__), "mora.csv")

def calculate_totals(df: pd.DataFrame):
    if df.empty:
        return {
            "total_cuotas": 0,
            "total_clientes": 0,
            "total_creditos": 0,
            "monto_total": 0.0,
            "pendiente_tarifa": 0.0,
            "pendiente_cuota": 0.0,
            "pendiente_mora": 0.0,
            "promedio_cliente": 0.0,
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

    return {
        "total_cuotas": total_cuotas,
        "total_clientes": total_clientes,
        "total_creditos": total_creditos,
        "monto_total": round(monto_total, 2),
        "pendiente_tarifa": round(p_tarifa, 2),
        "pendiente_cuota": round(p_cuota, 2),
        "pendiente_mora": round(p_mora, 2),
        "promedio_cliente": round(promedio, 2),
    }

def get_metadata(df: pd.DataFrame):
    programas = sorted([p for p in df["Programa"].dropna().unique().tolist() if p]) if "Programa" in df.columns else []
    tipos = sorted([t for t in df["Tipo"].dropna().unique().tolist() if t]) if "Tipo" in df.columns else []
    canales = sorted([c for c in df["Canal"].dropna().unique().tolist() if c]) if "Canal" in df.columns else []
    
    min_dias = int(df["Dias de Mora"].min()) if "Dias de Mora" in df.columns and not df.empty else 0
    max_dias = int(df["Dias de Mora"].max()) if "Dias de Mora" in df.columns and not df.empty else 0

    return {
        "total_rows": len(df),
        "total_clientes": int(df["Cedula Cliente"].nunique()) if "Cedula Cliente" in df.columns else 0,
        "programas": programas,
        "tipos": tipos,
        "canales": canales,
        "min_dias": min_dias,
        "max_dias": max_dias,
    }

class FilterRequest(BaseModel):
    min_dias: int = 1
    max_dias: int = 60
    programas: Optional[List[str]] = None
    tipos: Optional[List[str]] = None
    group_by: str = "Programa"

# -------------------------------------------------------------
# RUTAS DE LA API
# -------------------------------------------------------------

@app.get("/api/status")
def get_status():
    if DATA_STORE["df_original"] is None:
        if os.path.exists(DEFAULT_CSV_PATH):
            return {"loaded": False, "sample_available": True, "sample_path": "mora.csv"}
        return {"loaded": False, "sample_available": False}
    
    meta = get_metadata(DATA_STORE["df_original"])
    totals = calculate_totals(DATA_STORE["df_original"])
    return {
        "loaded": True,
        "filename": DATA_STORE["filename"],
        "metadata": meta,
        "totals": totals
    }

@app.post("/api/load-sample")
def load_sample():
    if not os.path.exists(DEFAULT_CSV_PATH):
        raise HTTPException(status_code=404, detail="Archivo mora.csv no encontrado en el servidor.")
    
    try:
        df = pd.read_csv(DEFAULT_CSV_PATH, encoding="utf-8-sig")
    except Exception:
        try:
            df = pd.read_csv(DEFAULT_CSV_PATH, encoding="latin1")
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"No se pudo leer mora.csv: {e}")
    
    df = standardize_dataframe(df)
    
    DATA_STORE["df_original"] = df
    DATA_STORE["df_filtered"] = df
    DATA_STORE["filename"] = "mora.csv (Archivo Base de Prueba)"
    
    return {
        "message": "Archivo de prueba cargado exitosamente.",
        "metadata": get_metadata(df),
        "totals": calculate_totals(df)
    }

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
    
    sheet_selected_name = ""
    try:
        filename_lower = file.filename.lower()
        
        # 1. Es un CSV
        if filename_lower.endswith(".csv"):
            df = None
            encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
            separators = [None, ",", ";", "\t"]
            
            for enc in encodings:
                for sep in separators:
                    try:
                        temp_df = pd.read_csv(io.BytesIO(contents), encoding=enc, sep=sep, engine="python")
                        # Verificar si se separó en más de 1 columna
                        if len(temp_df.columns) > 1:
                            df = temp_df
                            break
                    except Exception:
                        continue
                if df is not None:
                    break
                    
            if df is None:
                raise HTTPException(status_code=400, detail="No se pudo interpretar el formato ni la codificación del archivo CSV.")
                
        # 2. Es un archivo Excel (.xls o .xlsx)
        elif filename_lower.endswith((".xls", ".xlsx")):
            # Motores en orden de velocidad: calamine (Rust 10x más rápido), openpyxl, xlrd
            engines = ["calamine", "openpyxl", "xlrd"]
            xl = None
            for eng in engines:
                try:
                    xl = pd.ExcelFile(io.BytesIO(contents), engine=eng)
                    break
                except Exception:
                    continue
                    
            if xl is None:
                # Intentar lectura HTML (algunos sistemas exportan tablas HTML con extension .xls)
                try:
                    html_dfs = pd.read_html(io.BytesIO(contents))
                    if html_dfs:
                        df = html_dfs[0]
                except Exception:
                    pass
                if 'df' not in locals() or df is None:
                    raise HTTPException(status_code=400, detail="No se pudo abrir el archivo Excel. Verifica que no esté dañado.")
            else:
                # Sistema de puntuación para seleccionar la hoja con más datos relevantes
                sheet_scores = []
                for s in xl.sheet_names:
                    try:
                        sample = xl.parse(s, nrows=4)
                        cols = [clean_col(str(c)) for c in sample.columns]
                        score = 0
                        s_norm = clean_col(s)
                        
                        # Prioridad de nombre de hoja
                        if any(k in s_norm for k in ["morageneral", "cartera", "base", "detalle"]):
                            score += 50
                        elif any(k in s_norm for k in ["clientes", "consolidado", "callcenter", "mora"]):
                            score += 30
                            
                        # Prioridad si contiene las columnas requeridas
                        has_dias = any(any(re.match(p, c) for p in TARGET_PATTERNS["Dias de Mora"]) for c in cols)
                        has_cedula = any(any(re.match(p, c) for p in TARGET_PATTERNS["Cedula Cliente"]) for c in cols)
                        
                        if has_dias:
                            score += 60
                        if has_cedula:
                            score += 60
                            
                        score += min(len(sample.columns) * 2, 40)
                        sheet_scores.append((s, score))
                    except Exception:
                        sheet_scores.append((s, -1))
                        
                sheet_scores.sort(key=lambda x: x[1], reverse=True)
                chosen_sheet = sheet_scores[0][0] if sheet_scores and sheet_scores[0][1] > 0 else xl.sheet_names[0]
                
                sheet_selected_name = f" (Hoja: '{chosen_sheet.strip()}')"
                df = xl.parse(chosen_sheet)
                
        else:
            # Archivo con extension desconocida: intentar Excel y luego CSV
            try:
                df = pd.read_excel(io.BytesIO(contents))
            except Exception:
                try:
                    df = pd.read_csv(io.BytesIO(contents), encoding="latin1")
                except Exception:
                    raise HTTPException(status_code=400, detail="Formato de archivo no soportado. Usa .csv, .xlsx o .xls.")
        
        # Procesar y estandarizar datos
        df = standardize_dataframe(df)
        
        DATA_STORE["df_original"] = df
        DATA_STORE["df_filtered"] = df
        DATA_STORE["filename"] = file.filename + sheet_selected_name
        
        return {
            "message": f"Archivo cargado con éxito.{sheet_selected_name}",
            "metadata": get_metadata(df),
            "totals": calculate_totals(df)
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error al procesar el archivo: {str(e)}")

@app.post("/api/process")
def process_data(req: FilterRequest):
    if DATA_STORE["df_original"] is None:
        raise HTTPException(status_code=400, detail="No se ha cargado ningún archivo de cartera.")
    
    df = DATA_STORE["df_original"].copy()
    
    # 1. Filtro por dias de mora
    df = df[(df["Dias de Mora"] >= req.min_dias) & (df["Dias de Mora"] <= req.max_dias)]
    
    # 2. Filtro por programas (campañas)
    if req.programas and len(req.programas) > 0 and "Programa" in df.columns:
        df = df[df["Programa"].isin(req.programas)]
        
    # 3. Filtro por tipos
    if req.tipos and len(req.tipos) > 0 and "Tipo" in df.columns:
        df = df[df["Tipo"].isin(req.tipos)]
    
    DATA_STORE["df_filtered"] = df
    DATA_STORE["filter_info"] = req.dict()
    
    kpis = calculate_totals(df)
    
    # Calculo de agrupacion dinamica
    group_col = req.group_by if req.group_by in df.columns else "Programa"
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
        
        # Ordenar de mayor a menor monto
        agrupacion = sorted(agrupacion, key=lambda x: x["total_monto"], reverse=True)
    
    # Vista previa: Primeras 100 filas
    cols_to_show = [c for c in df.columns if not c.startswith("_")]
    preview_df = df[cols_to_show].head(100).fillna("")
    preview = preview_df.to_dict(orient="records")
    
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
    if DATA_STORE["df_filtered"] is None or DATA_STORE["df_filtered"].empty:
        raise HTTPException(status_code=400, detail="No hay datos filtrados para descargar.")
    
    df = DATA_STORE["df_filtered"].copy()
    clean_cols = [c for c in df.columns if not c.startswith("_")]
    df_clean = df[clean_cols]
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_clean.to_excel(writer, index=False, sheet_name="Cartera Desglosada")
    
    output.seek(0)
    filename = f"cartera_desglose_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/api/download/pdf")
def download_pdf():
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
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=15
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
        fontSize=14,
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
    story.append(Paragraph("REPORTE EJECUTIVO DE CARTERA VENCIDA", title_style))
    sub_text = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Rango mora: {f_info.get('min_dias', 1)} a {f_info.get('max_dias', 60)} días | Archivo: {html.escape(str(DATA_STORE.get('filename', 'Cartera')))}"
    story.append(Paragraph(sub_text, subtitle_style))
    story.append(Spacer(1, 10))
    
    # KPI Grid Table
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
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 20))
    
    # Subtotales
    story.append(Paragraph(f"<b>Desglose del Monto Total Vencido:</b> Tarifa Pendiente: <b>${kpis['pendiente_tarifa']:,.2f}</b> | Cuota Pendiente: <b>${kpis['pendiente_cuota']:,.2f}</b> | Recargo Mora: <b>${kpis['pendiente_mora']:,.2f}</b>", styles['Normal']))
    story.append(Spacer(1, 15))
    
    # Agrupacion
    story.append(Paragraph("<b>Distribución de Deuda por Segmento:</b>", styles['Heading3']))
    story.append(Spacer(1, 8))
    
    grp_col = f_info.get('group_by', 'Programa')
    if grp_col not in df.columns:
        grp_col = 'Programa'
    
    if (df["Pendiente Tarifa"].sum() + df["Pendiente Cuota"].sum() + df["Pendiente Mora"].sum()) > 0:
        df["_Tot"] = df["Pendiente Tarifa"] + df["Pendiente Cuota"] + df["Pendiente Mora"]
    elif "Total" in df.columns:
        df["_Tot"] = df["Total"]
    else:
        df["_Tot"] = 0.0

    grp = df.groupby(grp_col).agg(
        cuotas=("Cedula Cliente", "count"),
        clientes=("Cedula Cliente", "nunique"),
        total=("_Tot", "sum")
    ).reset_index().sort_values(by="total", ascending=False)
    
    tot_g = kpis["monto_total"]
    
    table_rows = [
        [
            Paragraph("Campaña / Grupo", table_header),
            Paragraph("Clientes", table_header),
            Paragraph("Cuotas", table_header),
            Paragraph("Monto Vencido", table_header),
            Paragraph("% Total", table_header)
        ]
    ]
    
    for _, r in grp.iterrows():
        m_val = float(r["total"])
        pct = (m_val / tot_g * 100) if tot_g > 0 else 0
        table_rows.append([
            Paragraph(html.escape(str(r[grp_col])), table_cell),
            Paragraph(f"{int(r['clientes']):,}", table_cell),
            Paragraph(f"{int(r['cuotas']):,}", table_cell),
            Paragraph(f"${m_val:,.2f}", table_cell),
            Paragraph(f"{pct:.1f}%", table_cell)
        ])
    
    t_grp = Table(table_rows, colWidths=[180, 80, 80, 120, 80])
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
# RUTAS DE CRUCE DE CONTACTOS (CALL CENTER / WHATSAPP)
# -------------------------------------------------------------

@app.get("/api/contacts/status")
def get_contacts_status():
    excel_path = os.path.join(os.path.dirname(__file__), "Base de Datos Correos y Teléfonos al 06-08 (2).xlsx")
    sample_available = os.path.exists(excel_path)
    
    if DATA_STORE["df_contacts"] is None:
        return {
            "loaded": False,
            "sample_available": sample_available,
            "sample_name": "Base de Datos Correos y Teléfonos al 06-08 (2).xlsx"
        }
    
    return {
        "loaded": True,
        "filename": DATA_STORE["contacts_filename"],
        "total_contacts": len(DATA_STORE["df_contacts"]),
        "sample_available": sample_available
    }

@app.post("/api/contacts/load-sample")
def load_sample_contacts():
    excel_path = os.path.join(os.path.dirname(__file__), "Base de Datos Correos y Teléfonos al 06-08 (2).xlsx")
    if not os.path.exists(excel_path):
        raise HTTPException(status_code=404, detail="Archivo de contactos de ejemplo no encontrado en el servidor.")
        
    for eng in ["calamine", "openpyxl", "xlrd"]:
        try:
            df = pd.read_excel(excel_path, engine=eng)
            break
        except Exception:
            continue
            
    df = standardize_contacts_dataframe(df)
    DATA_STORE["df_contacts"] = df
    DATA_STORE["contacts_filename"] = "Base de Contactos (Ejemplo Servidor)"
    
    return {
        "message": "Base de contactos cargada exitosamente.",
        "total_contacts": len(df),
        "filename": DATA_STORE["contacts_filename"]
    }

@app.post("/api/contacts/upload")
async def upload_contacts(file: UploadFile = File(...)):
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
        
    filename_lower = file.filename.lower()
    df = None
    try:
        if filename_lower.endswith(".csv"):
            for enc in ["utf-8-sig", "utf-8", "latin1"]:
                for sep in [None, ",", ";", "\t"]:
                    try:
                        temp_df = pd.read_csv(io.BytesIO(contents), encoding=enc, sep=sep, engine="python")
                        if len(temp_df.columns) > 1:
                            df = temp_df
                            break
                    except Exception:
                        continue
                if df is not None:
                    break
        elif filename_lower.endswith((".xlsx", ".xls")):
            # Motores en orden de rendimiento (calamine lee 200k filas en segundos)
            for eng in ["calamine", "openpyxl", "xlrd"]:
                try:
                    df = pd.read_excel(io.BytesIO(contents), engine=eng)
                    break
                except Exception:
                    continue
        else:
            raise HTTPException(status_code=400, detail="Formato no soportado. Usa .xlsx, .xls o .csv")
            
        if df is None:
            raise HTTPException(status_code=400, detail="No se pudo interpretar el archivo de contactos. Verifica que no esté dañado.")
            
        df = standardize_contacts_dataframe(df)
        DATA_STORE["df_contacts"] = df
        DATA_STORE["contacts_filename"] = file.filename
        
        return {
            "message": f"Base de contactos cargada: {len(df):,} clientes listos para cruce.",
            "total_contacts": len(df),
            "filename": file.filename
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error al procesar archivo de contactos: {str(e)}")

@app.post("/api/contacts/cross-match")
def cross_match_contacts():
    if DATA_STORE["df_filtered"] is None:
        raise HTTPException(status_code=400, detail="Primero debes cargar y procesar la cartera.")
    if DATA_STORE["df_contacts"] is None:
        raise HTTPException(status_code=400, detail="Primero debes cargar la base de contactos.")
        
    df_cartera = DATA_STORE["df_filtered"].copy()
    df_contacts = DATA_STORE["df_contacts"].copy()
    
    # Consolidar cartera a nivel de cliente
    df_cartera["Cedula_Key"] = df_cartera["Cedula Cliente"].astype(str).str.strip().str.upper()
    df_cartera["_Fila_Total"] = df_cartera["Pendiente Tarifa"] + df_cartera["Pendiente Cuota"] + df_cartera["Pendiente Mora"]
    
    debtors = df_cartera.groupby("Cedula_Key").agg(
        Cedula=("Cedula Cliente", "first"),
        Nombre_Cliente=("Nombre Cliente", "first"),
        Cuotas_Vencidas=("Cedula Cliente", "count"),
        Monto_Total_USD=("_Fila_Total", "sum"),
        Pendiente_Tarifa=("Pendiente Tarifa", "sum"),
        Pendiente_Cuota=("Pendiente Cuota", "sum"),
        Pendiente_Mora=("Pendiente Mora", "sum"),
        Max_Dias_Mora=("Dias de Mora", "max"),
        Programa=("Programa", "first"),
        Tipo=("Tipo", lambda x: "/".join(sorted(set(str(v) for v in x if str(v) != 'nan'))))
    ).reset_index()
    
    debtors["Monto_Total_USD"] = debtors["Monto_Total_USD"].round(2)
    debtors["Pendiente_Tarifa"] = debtors["Pendiente_Tarifa"].round(2)
    debtors["Pendiente_Cuota"] = debtors["Pendiente_Cuota"].round(2)
    debtors["Pendiente_Mora"] = debtors["Pendiente_Mora"].round(2)
    
    # Cruzar mediante Cedula_Key
    merged = pd.merge(
        debtors,
        df_contacts[["Cedula_Key", "Telefono_1", "Telefono_2", "Telefono_3", "Correo"]],
        on="Cedula_Key",
        how="left"
    )
    
    for col in ["Telefono_1", "Telefono_2", "Telefono_3", "Correo"]:
        merged[col] = merged[col].fillna("").astype(str).str.strip()
        
    merged = merged.sort_values(by="Monto_Total_USD", ascending=False)
    DATA_STORE["df_matched"] = merged
    
    # Métricas de cobertura
    has_tel = (merged["Telefono_1"] != "") | (merged["Telefono_2"] != "") | (merged["Telefono_3"] != "")
    has_mail = merged["Correo"] != ""
    has_both = has_tel & has_mail
    has_any = has_tel | has_mail
    has_none = ~has_any
    
    total_deudores = len(merged)
    total_con_tel = int(has_tel.sum())
    total_con_mail = int(has_mail.sum())
    total_con_any = int(has_any.sum())
    total_sin_contacto = int(has_none.sum())
    
    monto_total_cartera = float(merged["Monto_Total_USD"].sum())
    monto_contactable = float(merged.loc[has_any, "Monto_Total_USD"].sum())
    monto_sin_contacto = float(merged.loc[has_none, "Monto_Total_USD"].sum())
    
    pct_cobertura = round((total_con_any / total_deudores * 100), 1) if total_deudores > 0 else 0
    pct_monto_contactable = round((monto_contactable / monto_total_cartera * 100), 1) if monto_total_cartera > 0 else 0
    
    preview = []
    for _, r in merged.head(100).iterrows():
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
    if DATA_STORE["df_matched"] is None or DATA_STORE["df_matched"].empty:
        raise HTTPException(status_code=400, detail="No se ha realizado el cruce de cartera con contactos.")
        
    df = DATA_STORE["df_matched"].copy()
    if "Cedula_Key" in df.columns:
        df = df.drop(columns=["Cedula_Key"])
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Contactos Call Center")
        
    output.seek(0)
    filename = f"cartera_enriquecida_callcenter_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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
