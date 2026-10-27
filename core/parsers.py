import io
import os
import re
import unicodedata
import pandas as pd
from fastapi import HTTPException
from core.config import TARGET_PATTERNS


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
    """Estandariza columnas, teléfonos, campaña y cédula del archivo de contactos o campañas."""
    col_map = {}
    for c in df.columns:
        norm = clean_col(str(c))
        if any(re.match(p, norm) for p in [r"^cedula.*", r"^dni$", r"^identificacion$", r"^etiquetasdefila$", r"^clienteci$", r"^ci$"]):
            col_map[c] = "Cedula"
        elif any(re.match(p, norm) for p in [r"^telefono1$", r"^tel1$", r"^telf1$", r"^celular1$", r"^celular$", r"^telefono$", r".*telefonomovil.*", r".*telefonoppal.*"]):
            col_map[c] = "Telefono_1"
        elif any(re.match(p, norm) for p in [r"^telefono2$", r"^tel2$", r"^telf2$", r"^celular2$"]):
            col_map[c] = "Telefono_2"
        elif any(re.match(p, norm) for p in [r"^telefono3$", r"^tel3$", r"^telf3$", r"^celular3$"]):
            col_map[c] = "Telefono_3"
        elif any(re.match(p, norm) for p in [r"^correo.*", r"^email.*", r"^mail$", r".*correoppal.*"]):
            col_map[c] = "Correo"
        elif any(re.match(p, norm) for p in [r"^nombre.*", r"^cliente$"]):
            col_map[c] = "Nombre_Contacto"
        elif any(re.match(p, norm) for p in [r"^tipodecampana.*", r"^campana.*", r"^tipocampana.*"]):
            col_map[c] = "Tipo_Campana"
        elif any(re.match(p, norm) for p in [r"^programa.*", r"^plan.*"]):
            col_map[c] = "Programa_Contacto"
        elif any(re.match(p, norm) for p in [r"^canal.*"]):
            col_map[c] = "Canal_Contacto"

    df = df.rename(columns=col_map)
    if "Cedula" not in df.columns:
        raise HTTPException(status_code=400, detail="El archivo de contactos o campañas debe contener una columna de Cédula (ej. 'Cedula' o 'cliente_ci').")

    df["Cedula"] = df["Cedula"].fillna("").astype(str).str.strip().str.upper()
    df = df[~df["Cedula"].isin(["", "NAN", "NULL", "NONE"])].copy()

    for col in ["Telefono_1", "Telefono_2", "Telefono_3"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_phone_str)
        else:
            df[col] = ""

    for col in ["Correo", "Nombre_Contacto", "Tipo_Campana", "Programa_Contacto", "Canal_Contacto"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "NULL": "", "None": "", "NAN": ""})
        else:
            df[col] = ""

    df["Cedula_Key"] = df["Cedula"]
    return df.drop_duplicates(subset=["Cedula_Key"])


def detect_file_type(df: pd.DataFrame) -> str:
    """Detecta inteligentemente si un DataFrame corresponde a Cartera ('cartera') o a Contactos/Campañas ('contacts')."""
    cleaned_cols = [clean_col(str(c)) for c in df.columns]

    # Indicadores exclusivos de Cartera / Mora (montos de tarifas, cuotas y mora)
    has_dias = any(any(re.match(p, c) for p in TARGET_PATTERNS["Dias de Mora"]) for c in cleaned_cols)
    has_tarifa = any(any(re.match(p, c) for p in TARGET_PATTERNS["Pendiente Tarifa"]) for c in cleaned_cols)
    has_cuota = any(any(re.match(p, c) for p in TARGET_PATTERNS["Pendiente Cuota"]) for c in cleaned_cols)
    has_mora = any(any(re.match(p, c) for p in TARGET_PATTERNS["Pendiente Mora"]) for c in cleaned_cols)
    has_debt_amounts = has_tarifa or has_cuota or has_mora

    # Indicadores de Contactos o Campañas
    has_phone = any(any(re.match(p, c) for p in [r"^telefono.*", r"^tel\d?$", r"^celular.*", r"^telf.*", r".*telefonomovil.*"]) for c in cleaned_cols)
    has_email = any(any(re.match(p, c) for p in [r"^correo.*", r"^email.*", r"^mail$", r".*correoppal.*"]) for c in cleaned_cols)
    has_campana = any(any(re.match(p, c) for p in [r"^tipodecampana.*", r"^campana.*", r"^tipocampana.*"]) for c in cleaned_cols)
    has_cliente_ci = any(c == "clienteci" for c in cleaned_cols)

    if has_debt_amounts or (has_dias and not has_campana and not has_phone and not has_cliente_ci):
        return "cartera"
    elif has_phone or has_email or has_campana or has_cliente_ci:
        return "contacts"
    return "unknown"
