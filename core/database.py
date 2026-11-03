import base64
import hashlib
import hmac
import json
import secrets
import sqlite3
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from typing import Optional

import pandas as pd
from core.config import (
    DATA_STORE,
    DB_PATH,
    SESSION_SECRET,
    SUPABASE_SERVICE_KEY,
    SUPABASE_URL,
)


# -------------------------------------------------------------
# SEGURIDAD: HASH DE CONTRASEÑAS Y TOKENS DE SESIÓN
# -------------------------------------------------------------

def hash_password(password: str, salt: Optional[str] = None) -> str:
    if not salt:
        salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120_000,
    )
    return f"pbkdf2_sha256${salt}${dk.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    if not stored_hash or "$" not in stored_hash:
        return False
    parts = stored_hash.split("$")
    if len(parts) != 3 or parts[0] != "pbkdf2_sha256":
        return False
    _, salt, _ = parts
    expected = hash_password(password, salt=salt)
    return hmac.compare_digest(expected, stored_hash)


def create_session_token(user_data: dict, max_age_sec: int = 86400 * 7) -> str:
    payload = {
        "email": user_data.get("email", ""),
        "nombre": user_data.get("nombre", ""),
        "rol": user_data.get("rol", "admin"),
        "exp": int(time.time()) + max_age_sec,
    }
    raw_json = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    b64_payload = base64.urlsafe_b64encode(raw_json).decode("ascii").rstrip("=")
    sig = hmac.new(
        SESSION_SECRET.encode("utf-8"),
        b64_payload.encode("ascii"),
        hashlib.sha256,
    ).hexdigest()
    return f"{b64_payload}.{sig}"


def verify_session_token(token: Optional[str]) -> Optional[dict]:
    if not token or "." not in token:
        return None
    try:
        b64_payload, sig = token.rsplit(".", 1)
        expected_sig = hmac.new(
            SESSION_SECRET.encode("utf-8"),
            b64_payload.encode("ascii"),
            hashlib.sha256,
        ).hexdigest()
        if not hmac.compare_digest(sig, expected_sig):
            return None
        padding = "=" * (-len(b64_payload) % 4)
        payload = json.loads(base64.urlsafe_b64decode(b64_payload + padding).decode("utf-8"))
        if payload.get("exp", 0) < int(time.time()):
            return None
        return payload
    except Exception:
        return None


# -------------------------------------------------------------
# CLIENTE REST SUPABASE (POSTGREST)
# -------------------------------------------------------------

def _supabase_request(
    method: str,
    endpoint: str,
    data=None,
    prefer: Optional[str] = None,
    timeout: int = 12,
):
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        return None
    url = f"{SUPABASE_URL}/rest/v1/{endpoint.lstrip('/')}"
    headers = {
        "apikey": SUPABASE_SERVICE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "Content-Type": "application/json",
        "User-Agent": "CobranzApp/2.0",
    }
    if prefer:
        headers["Prefer"] = prefer

    body_bytes = None
    if data is not None:
        body_bytes = json.dumps(data, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(url, data=body_bytes, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8").strip()
        if not raw:
            return []
        return json.loads(raw)


# -------------------------------------------------------------
# INICIALIZACIÓN DE TABLAS Y SEMILLA DE USUARIOS
# -------------------------------------------------------------

def ensure_db_table():
    """Garantiza que existan las tablas locales en SQLite y los usuarios base en Supabase y SQLite."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS master_clientes (
            cedula_key TEXT PRIMARY KEY,
            nombre TEXT DEFAULT '',
            telefono_1 TEXT DEFAULT '',
            telefono_2 TEXT DEFAULT '',
            telefono_3 TEXT DEFAULT '',
            correo TEXT DEFAULT '',
            tipo_campana TEXT DEFAULT '',
            programa TEXT DEFAULT '',
            canal TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            nombre TEXT NOT NULL DEFAULT '',
            password_hash TEXT NOT NULL,
            rol TEXT NOT NULL DEFAULT 'admin',
            activo INTEGER NOT NULL DEFAULT 1,
            created_at TEXT DEFAULT '',
            last_login TEXT DEFAULT ''
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS historial_cargas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT DEFAULT '',
            fecha_corte TEXT DEFAULT '',
            modulo TEXT NOT NULL DEFAULT 'general',
            modulo_label TEXT NOT NULL DEFAULT 'Mora Ubi',
            nombre_archivo TEXT NOT NULL DEFAULT '',
            origen TEXT NOT NULL DEFAULT 'web',
            usuario_email TEXT DEFAULT '',
            cartera_activa_usd REAL,
            total_mora_usd REAL NOT NULL DEFAULT 0,
            pendiente_tarifa_usd REAL NOT NULL DEFAULT 0,
            pendiente_cuota_usd REAL NOT NULL DEFAULT 0,
            pendiente_mora_usd REAL NOT NULL DEFAULT 0,
            porcentaje_mora REAL,
            total_cuotas INTEGER NOT NULL DEFAULT 0,
            total_clientes INTEGER NOT NULL DEFAULT 0,
            clientes_con_telefono INTEGER NOT NULL DEFAULT 0,
            cobertura_contacto_pct REAL NOT NULL DEFAULT 0,
            resumen_tramos TEXT DEFAULT '{}',
            resumen_fibex TEXT DEFAULT '{}',
            filtros_aplicados TEXT DEFAULT '{}'
        )
    """)
    conn.commit()

    # Sembrar usuarios por defecto en SQLite si no existen
    default_users = [
        ("admin@credix.com", "Administrador Credix", "Credix2026*", "admin"),
        ("francojbe@gmail.com", "Franco (SuperAdmin)", "Admin2026*", "admin"),
    ]
    now_iso = datetime.now(timezone.utc).isoformat()
    for email, nombre, raw_pwd, rol in default_users:
        cur.execute("SELECT id FROM usuarios WHERE lower(email) = lower(?)", (email,))
        if not cur.fetchone():
            cur.execute(
                "INSERT INTO usuarios (email, nombre, password_hash, rol, activo, created_at) VALUES (?, ?, ?, ?, 1, ?)",
                (email.lower(), nombre, hash_password(raw_pwd), rol, now_iso),
            )
    conn.commit()
    conn.close()


def ensure_supabase_default_users():
    """Garantiza que los usuarios iniciales existan también en Supabase public.usuarios."""
    try:
        existing = _supabase_request("GET", "usuarios?select=email", timeout=8)
        existing_emails = {str(u.get("email", "")).lower() for u in (existing or [])}
        default_users = [
            ("admin@credix.com", "Administrador Credix", "Credix2026*", "admin"),
            ("francojbe@gmail.com", "Franco (SuperAdmin)", "Admin2026*", "admin"),
        ]
        to_insert = []
        for email, nombre, raw_pwd, rol in default_users:
            if email.lower() not in existing_emails:
                to_insert.append({
                    "email": email.lower(),
                    "nombre": nombre,
                    "password_hash": hash_password(raw_pwd),
                    "rol": rol,
                    "activo": True,
                })
        if to_insert:
            _supabase_request(
                "POST",
                "usuarios?on_conflict=email",
                data=to_insert,
                prefer="resolution=merge-duplicates,return=minimal",
                timeout=10,
            )
    except Exception as e:
        print("Aviso al sincronizar usuarios iniciales en Supabase:", e)


# -------------------------------------------------------------
# BASE MAESTRA DE CLIENTES (SQLITE + SUPABASE)
# -------------------------------------------------------------

def load_contacts_from_db() -> pd.DataFrame:
    """Lee la tabla master_clientes de SQLite a memoria (caché ultrarrápido) o descarga desde Supabase si está vacía."""
    ensure_db_table()
    conn = sqlite3.connect(DB_PATH)
    df_db = pd.read_sql_query("""
        SELECT
            cedula_key AS Cedula_Key,
            cedula_key AS Cedula,
            nombre AS Nombre_Contacto,
            telefono_1 AS Telefono_1,
            telefono_2 AS Telefono_2,
            telefono_3 AS Telefono_3,
            correo AS Correo,
            tipo_campana AS Tipo_Campana,
            programa AS Programa_Contacto,
            canal AS Canal_Contacto
        FROM master_clientes
    """, conn)
    conn.close()

    if not df_db.empty:
        DATA_STORE["df_contacts_raw"] = df_db
        DATA_STORE["contacts_filename"] = f"Base Supabase Activa ({len(df_db):,} clientes)"
    return df_db


def _sync_contacts_batches_to_supabase(rows: list):
    """Sincroniza en lotes hacia Supabase public.master_clientes."""
    if not rows or not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        return
    batch_size = 1500
    batches = []
    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i + batch_size]
        payload = [
            {
                "cedula_key": r[0],
                "nombre": r[1] or "",
                "telefono_1": r[2] or "",
                "telefono_2": r[3] or "",
                "telefono_3": r[4] or "",
                "correo": r[5] or "",
                "tipo_campana": r[6] or "",
                "programa": r[7] or "",
                "canal": r[8] or "",
                "updated_at": r[9] or "",
            }
            for r in chunk
        ]
        batches.append(payload)

    def _post_batch(b):
        try:
            _supabase_request(
                "POST",
                "master_clientes?on_conflict=cedula_key",
                data=b,
                prefer="resolution=merge-duplicates,return=minimal",
                timeout=25,
            )
        except Exception as err:
            print("Aviso al sincronizar lote de contactos con Supabase:", err)

    with ThreadPoolExecutor(max_workers=4) as ex:
        list(ex.map(_post_batch, batches))


def upsert_contacts_to_db(df: pd.DataFrame) -> dict:
    """Inserta clientes nuevos o actualiza campos no vacíos en SQLite y Supabase sin perder información previa."""
    ensure_db_table()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM master_clientes")
    count_before = cur.fetchone()[0]

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for _, r in df.iterrows():
        ced = str(r.get("Cedula_Key", "")).strip().upper()
        if not ced or ced in ("NAN", "NULL", "NONE"):
            continue
        rows.append((
            ced,
            str(r.get("Nombre_Contacto", "")).strip(),
            str(r.get("Telefono_1", "")).strip(),
            str(r.get("Telefono_2", "")).strip(),
            str(r.get("Telefono_3", "")).strip(),
            str(r.get("Correo", "")).strip(),
            str(r.get("Tipo_Campana", "")).strip(),
            str(r.get("Programa_Contacto", "")).strip(),
            str(r.get("Canal_Contacto", "")).strip(),
            now_str
        ))

    cur.executemany("""
        INSERT INTO master_clientes (
            cedula_key, nombre, telefono_1, telefono_2, telefono_3,
            correo, tipo_campana, programa, canal, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(cedula_key) DO UPDATE SET
            nombre = CASE WHEN excluded.nombre != '' THEN excluded.nombre ELSE master_clientes.nombre END,
            telefono_1 = CASE WHEN excluded.telefono_1 != '' THEN excluded.telefono_1 ELSE master_clientes.telefono_1 END,
            telefono_2 = CASE WHEN excluded.telefono_2 != '' THEN excluded.telefono_2 ELSE master_clientes.telefono_2 END,
            telefono_3 = CASE WHEN excluded.telefono_3 != '' THEN excluded.telefono_3 ELSE master_clientes.telefono_3 END,
            correo = CASE WHEN excluded.correo != '' THEN excluded.correo ELSE master_clientes.correo END,
            tipo_campana = CASE WHEN excluded.tipo_campana != '' AND excluded.tipo_campana != 'Sin Campaña Asignada' THEN excluded.tipo_campana ELSE master_clientes.tipo_campana END,
            programa = CASE WHEN excluded.programa != '' THEN excluded.programa ELSE master_clientes.programa END,
            canal = CASE WHEN excluded.canal != '' THEN excluded.canal ELSE master_clientes.canal END,
            updated_at = excluded.updated_at
    """, rows)
    conn.commit()

    cur.execute("SELECT COUNT(*) FROM master_clientes")
    count_after = cur.fetchone()[0]
    conn.close()

    # Sincronizar también en Supabase public.master_clientes
    try:
        _sync_contacts_batches_to_supabase(rows)
    except Exception as e:
        print("Aviso al sincronizar contactos con Supabase:", e)

    # Refrescar DATA_STORE desde la base de datos actualizada
    load_contacts_from_db()

    return {
        "processed": len(rows),
        "added": max(0, count_after - count_before),
        "total_in_db": count_after
    }


# -------------------------------------------------------------
# AUTENTICACIÓN Y GESTIÓN DE USUARIOS (SUPABASE + SQLITE)
# -------------------------------------------------------------

def authenticate_user(email: str, password: str) -> Optional[dict]:
    """Valida credenciales contra Supabase public.usuarios (con respaldo en SQLite local)."""
    clean_email = (email or "").strip().lower()
    if not clean_email or not password:
        return None

    ensure_db_table()
    now_iso = datetime.now(timezone.utc).isoformat()

    # 1. Intentar autenticar contra Supabase
    try:
        encoded_email = urllib.parse.quote(clean_email)
        rows = _supabase_request(
            "GET",
            f"usuarios?email=eq.{encoded_email}&activo=eq.true&limit=1",
            timeout=8,
        )
        if rows and len(rows) > 0:
            user = rows[0]
            if verify_password(password, user.get("password_hash", "")):
                try:
                    _supabase_request(
                        "PATCH",
                        f"usuarios?id=eq.{user['id']}",
                        data={"last_login": now_iso},
                        prefer="return=minimal",
                        timeout=5,
                    )
                except Exception:
                    pass
                return {
                    "id": user.get("id"),
                    "email": user.get("email"),
                    "nombre": user.get("nombre") or clean_email.split("@")[0],
                    "rol": user.get("rol") or "admin",
                    "source": "supabase",
                }
            return None
    except Exception as e:
        print("Aviso al consultar usuario en Supabase (usando respaldo local):", e)

    # 2. Respaldo en SQLite local
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "SELECT id, email, nombre, password_hash, rol, activo FROM usuarios WHERE lower(email) = ?",
        (clean_email,),
    )
    row = cur.fetchone()
    if row and row[5] == 1 and verify_password(password, row[3]):
        cur.execute("UPDATE usuarios SET last_login = ? WHERE id = ?", (now_iso, row[0]))
        conn.commit()
        conn.close()
        return {
            "id": row[0],
            "email": row[1],
            "nombre": row[2],
            "rol": row[4],
            "source": "sqlite",
        }
    conn.close()
    return None


def list_users() -> list:
    """Lista los usuarios registrados en Supabase public.usuarios."""
    try:
        rows = _supabase_request(
            "GET",
            "usuarios?select=id,email,nombre,rol,activo,created_at,last_login&order=id.asc",
            timeout=8,
        )
        if rows is not None:
            return rows
    except Exception:
        pass

    ensure_db_table()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT id, email, nombre, rol, activo, created_at, last_login FROM usuarios ORDER BY id ASC")
    res = [
        {
            "id": r[0],
            "email": r[1],
            "nombre": r[2],
            "rol": r[3],
            "activo": bool(r[4]),
            "created_at": r[5],
            "last_login": r[6],
        }
        for r in cur.fetchall()
    ]
    conn.close()
    return res


def create_or_update_user(email: str, nombre: str, password: str, rol: str = "analista") -> dict:
    """Crea o actualiza un usuario en Supabase public.usuarios y en SQLite local."""
    clean_email = (email or "").strip().lower()
    clean_name = (nombre or "").strip() or clean_email.split("@")[0]
    pwd_hash = hash_password(password)
    now_iso = datetime.now(timezone.utc).isoformat()

    ensure_db_table()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO usuarios (email, nombre, password_hash, rol, activo, created_at)
        VALUES (?, ?, ?, ?, 1, ?)
        ON CONFLICT(email) DO UPDATE SET
            nombre = excluded.nombre,
            password_hash = excluded.password_hash,
            rol = excluded.rol,
            activo = 1
    """, (clean_email, clean_name, pwd_hash, rol, now_iso))
    conn.commit()
    conn.close()

    try:
        _supabase_request(
            "POST",
            "usuarios?on_conflict=email",
            data=[{
                "email": clean_email,
                "nombre": clean_name,
                "password_hash": pwd_hash,
                "rol": rol,
                "activo": True,
            }],
            prefer="resolution=merge-duplicates,return=representation",
            timeout=10,
        )
    except Exception as e:
        print("Aviso al guardar usuario en Supabase:", e)

    return {"email": clean_email, "nombre": clean_name, "rol": rol}


# -------------------------------------------------------------
# HISTÓRICO DE CARGAS DE MORA (SUPABASE public.historial_cargas)
# -------------------------------------------------------------

def check_existing_cutoff_snapshot(
    modulo: str,
    fecha_corte: str,
    total_cuotas: int,
    total_mora_usd: float,
    total_clientes: int,
    nombre_archivo: str = "",
) -> Optional[dict]:
    """Verifica si ya existe un registro en Supabase para la misma (fecha_corte, modulo) con datos distintos."""
    mod_clean = "avila" if modulo == "avila" else "general"
    mod_label = "Mora Ávila" if mod_clean == "avila" else "Mora Ubi"
    try:
        encoded_date = urllib.parse.quote(str(fecha_corte))
        existing_matches = _supabase_request(
            "GET",
            f"historial_cargas?modulo=eq.{mod_clean}&fecha_corte=eq.{encoded_date}&order=id.desc&limit=5",
            timeout=8,
        )
        if existing_matches:
            em = existing_matches[0]
            same_cuotas = (int(em.get("total_cuotas") or 0) == int(total_cuotas))
            same_mora = abs(float(em.get("total_mora_usd") or 0.0) - float(total_mora_usd)) < 0.1
            same_clientes = (int(em.get("total_clientes") or 0) == int(total_clientes))
            if not (same_cuotas and same_mora and same_clientes):
                return {
                    "existing_id": em.get("id"),
                    "fecha_corte": fecha_corte,
                    "modulo": mod_clean,
                    "modulo_label": mod_label,
                    "existing_archivo": em.get("nombre_archivo") or "corte_anterior.csv",
                    "existing_mora_usd": round(float(em.get("total_mora_usd") or 0.0), 2),
                    "existing_cuotas": int(em.get("total_cuotas") or 0),
                    "existing_clientes": int(em.get("total_clientes") or 0),
                    "new_archivo": nombre_archivo or "nuevo_archivo.csv",
                    "new_mora_usd": round(float(total_mora_usd), 2),
                    "new_cuotas": int(total_cuotas),
                    "new_clientes": int(total_clientes),
                }
    except Exception as e:
        print("Aviso al verificar corte existente:", e)
    return None


def save_history_snapshot(
    modulo: str,
    nombre_archivo: str,
    origen: str = "web",
    usuario_email: str = "sistema@credix.com",
    cartera_activa_usd: Optional[float] = None,
    kpis: Optional[dict] = None,
    resumen_tramos: Optional[dict] = None,
    resumen_fibex: Optional[dict] = None,
    filtros: Optional[dict] = None,
    overwrite_existing: bool = False,
) -> Optional[dict]:
    """Guarda una instantánea histórica del corte de mora en Supabase public.historial_cargas (y respaldo SQLite)."""
    kpis = kpis or {}
    resumen_tramos = resumen_tramos or {}
    resumen_fibex = resumen_fibex or {}
    filtros = filtros or {}

    mod_clean = "avila" if modulo == "avila" else "general"
    mod_label = "Mora Ávila" if mod_clean == "avila" else "Mora Ubi"
    fecha_corte = str(resumen_tramos.get("cutoff_date") or datetime.now().strftime("%d-%m-%Y"))

    ca_val = cartera_activa_usd if cartera_activa_usd is not None else resumen_tramos.get("cartera_activa_usd")
    total_mora = float(resumen_tramos.get("total_cartera_usd") or kpis.get("monto_total") or 0.0)
    p_tarifa = float(kpis.get("pendiente_tarifa") or 0.0)
    p_cuota = float(kpis.get("pendiente_cuota") or 0.0)
    p_mora = float(kpis.get("pendiente_mora") or 0.0)
    pct_mora = resumen_tramos.get("total_porcentaje")
    if pct_mora is None and ca_val and float(ca_val) > 0:
        pct_mora = round((total_mora / float(ca_val)) * 100, 2)

    total_cuotas = int(resumen_tramos.get("total_cuotas") or kpis.get("total_cuotas") or 0)
    total_clientes = int(resumen_tramos.get("total_clientes") or kpis.get("total_clientes") or 0)
    clientes_tel = int(kpis.get("total_con_telefono") or 0)
    cobertura_pct = float(kpis.get("pct_con_telefono") or 0.0)

    record = {
        "fecha_corte": fecha_corte,
        "modulo": mod_clean,
        "modulo_label": mod_label,
        "nombre_archivo": str(nombre_archivo or "mora.csv"),
        "origen": str(origen or "web"),
        "usuario_email": str(usuario_email or "sistema@credix.com"),
        "cartera_activa_usd": round(float(ca_val), 2) if (ca_val is not None and float(ca_val) > 0) else None,
        "total_mora_usd": round(total_mora, 2),
        "pendiente_tarifa_usd": round(p_tarifa, 2),
        "pendiente_cuota_usd": round(p_cuota, 2),
        "pendiente_mora_usd": round(p_mora, 2),
        "porcentaje_mora": round(float(pct_mora), 2) if pct_mora is not None else None,
        "total_cuotas": total_cuotas,
        "total_clientes": total_clientes,
        "clientes_con_telefono": clientes_tel,
        "cobertura_contacto_pct": round(cobertura_pct, 1),
        "resumen_tramos": resumen_tramos,
        "resumen_fibex": resumen_fibex,
        "filtros_aplicados": filtros,
    }

    # 1. Guardar en Supabase public.historial_cargas con Detección Inteligente de Duplicados
    try:
        # Buscar si ya existe un corte idéntico (mismo módulo, misma fecha de corte, mismas cuotas y saldo)
        encoded_date = urllib.parse.quote(fecha_corte)
        existing_matches = _supabase_request(
            "GET",
            f"historial_cargas?modulo=eq.{mod_clean}&fecha_corte=eq.{encoded_date}&order=created_at.desc&limit=5",
            timeout=8,
        )
        if existing_matches:
            em = existing_matches[0]
            same_cuotas = (int(em.get("total_cuotas") or 0) == record["total_cuotas"])
            same_mora = abs(float(em.get("total_mora_usd") or 0) - record["total_mora_usd"]) < 0.1
            same_clientes = (int(em.get("total_clientes") or 0) == record["total_clientes"])

            if (same_cuotas and same_mora and same_clientes) or overwrite_existing:
                # Actualizar el registro existente de esa fecha de corte (y limpiar duplicados previos si los hubiera)
                updated = _supabase_request(
                    "PATCH",
                    f"historial_cargas?id=eq.{em['id']}",
                    data=record,
                    prefer="return=representation",
                    timeout=8,
                )
                for extra in existing_matches[1:]:
                    try:
                        _supabase_request("DELETE", f"historial_cargas?id=eq.{extra['id']}", prefer="return=minimal", timeout=5)
                    except Exception:
                        pass
                return updated[0] if updated else em
            else:
                # Existe un corte de la misma fecha y módulo pero con datos diferentes y aún no se ha confirmado sobrescribir
                return {
                    "status": "conflict",
                    "conflict": {
                        "existing_id": em.get("id"),
                        "fecha_corte": fecha_corte,
                        "modulo": mod_clean,
                        "modulo_label": mod_label,
                        "existing_archivo": em.get("nombre_archivo") or "corte_anterior.csv",
                        "existing_mora_usd": round(float(em.get("total_mora_usd") or 0.0), 2),
                        "existing_cuotas": int(em.get("total_cuotas") or 0),
                        "existing_clientes": int(em.get("total_clientes") or 0),
                        "new_archivo": record["nombre_archivo"],
                        "new_mora_usd": record["total_mora_usd"],
                        "new_cuotas": record["total_cuotas"],
                        "new_clientes": record["total_clientes"],
                    },
                }

        inserted = _supabase_request(
            "POST",
            "historial_cargas",
            data=[record],
            prefer="return=representation",
            timeout=10,
        )
        if inserted and len(inserted) > 0:
            return inserted[0]
    except Exception as e:
        print("Aviso al guardar historial en Supabase (guardando en SQLite local):", e)

    # 2. Respaldo en SQLite local
    try:
        ensure_db_table()
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        now_iso = datetime.now(timezone.utc).isoformat()
        cur.execute("""
            INSERT INTO historial_cargas (
                created_at, fecha_corte, modulo, modulo_label, nombre_archivo,
                origen, usuario_email, cartera_activa_usd, total_mora_usd,
                pendiente_tarifa_usd, pendiente_cuota_usd, pendiente_mora_usd,
                porcentaje_mora, total_cuotas, total_clientes,
                clientes_con_telefono, cobertura_contacto_pct,
                resumen_tramos, resumen_fibex, filtros_aplicados
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            now_iso,
            record["fecha_corte"],
            record["modulo"],
            record["modulo_label"],
            record["nombre_archivo"],
            record["origen"],
            record["usuario_email"],
            record["cartera_activa_usd"],
            record["total_mora_usd"],
            record["pendiente_tarifa_usd"],
            record["pendiente_cuota_usd"],
            record["pendiente_mora_usd"],
            record["porcentaje_mora"],
            record["total_cuotas"],
            record["total_clientes"],
            record["clientes_con_telefono"],
            record["cobertura_contacto_pct"],
            json.dumps(record["resumen_tramos"], ensure_ascii=False),
            json.dumps(record["resumen_fibex"], ensure_ascii=False),
            json.dumps(record["filtros_aplicados"], ensure_ascii=False),
        ))
        conn.commit()
        new_id = cur.lastrowid
        conn.close()
        record["id"] = new_id
        record["created_at"] = now_iso
        return record
    except Exception as e:
        print("Error en respaldo SQLite de historial:", e)
        return None


def _compute_module_comparison(items: list, mod_id: str) -> dict:
    mod_items = [x for x in items if x.get("modulo") == mod_id]
    if not mod_items:
        return {
            "has_data": False,
            "latest_mora_usd": 0.0,
            "latest_cartera_activa_usd": None,
            "latest_porcentaje_mora": None,
            "latest_clientes": 0,
            "latest_cuotas": 0,
            "latest_fecha": None,
            "prev_mora_usd": None,
            "delta_mora_usd": None,
            "delta_mora_pct": None,
            "total_cargas": 0,
        }
    latest = mod_items[0]
    prev = mod_items[1] if len(mod_items) > 1 else None
    lat_mora = float(latest.get("total_mora_usd") or 0.0)
    prev_mora = float(prev.get("total_mora_usd") or 0.0) if prev else None
    delta_usd = round(lat_mora - prev_mora, 2) if prev_mora is not None else None
    delta_pct = round(((lat_mora - prev_mora) / prev_mora) * 100, 2) if (prev_mora and prev_mora > 0) else None

    return {
        "has_data": True,
        "latest_id": latest.get("id"),
        "latest_mora_usd": lat_mora,
        "latest_cartera_activa_usd": latest.get("cartera_activa_usd"),
        "latest_porcentaje_mora": latest.get("porcentaje_mora"),
        "latest_clientes": int(latest.get("total_clientes") or 0),
        "latest_cuotas": int(latest.get("total_cuotas") or 0),
        "latest_fecha": latest.get("created_at"),
        "latest_fecha_corte": latest.get("fecha_corte"),
        "latest_archivo": latest.get("nombre_archivo"),
        "prev_mora_usd": prev_mora,
        "delta_mora_usd": delta_usd,
        "delta_mora_pct": delta_pct,
        "total_cargas": len(mod_items),
    }


def get_history_list(modulo: Optional[str] = None, limit: int = 100) -> dict:
    """Obtiene el listado histórico de cargas de mora desde Supabase public.historial_cargas."""
    select_cols = (
        "id,created_at,fecha_corte,modulo,modulo_label,nombre_archivo,origen,"
        "usuario_email,cartera_activa_usd,total_mora_usd,pendiente_tarifa_usd,"
        "pendiente_cuota_usd,pendiente_mora_usd,porcentaje_mora,total_cuotas,"
        "total_clientes,clientes_con_telefono,cobertura_contacto_pct,resumen_tramos"
    )
    query = f"historial_cargas?select={select_cols}&order=created_at.desc&limit={int(limit)}"
    if modulo in ("general", "avila"):
        query += f"&modulo=eq.{modulo}"

    items = []
    source = "supabase"
    try:
        rows = _supabase_request("GET", query, timeout=10)
        if rows is None:
            raise RuntimeError("Supabase no disponible (verificar SUPABASE_SERVICE_KEY en variables de entorno)")
        items = rows
    except Exception as e:
        print("Aviso al consultar historial en Supabase:", e)
        source = "sqlite"
        ensure_db_table()
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        sql = """
            SELECT id, created_at, fecha_corte, modulo, modulo_label, nombre_archivo,
                   origen, usuario_email, cartera_activa_usd, total_mora_usd,
                   pendiente_tarifa_usd, pendiente_cuota_usd, pendiente_mora_usd,
                   porcentaje_mora, total_cuotas, total_clientes,
                   clientes_con_telefono, cobertura_contacto_pct, resumen_tramos
            FROM historial_cargas
        """
        params = []
        if modulo in ("general", "avila"):
            sql += " WHERE modulo = ?"
            params.append(modulo)
        sql += " ORDER BY id DESC LIMIT ?"
        params.append(int(limit))
        cur.execute(sql, tuple(params))
        for r in cur.fetchall():
            try:
                rt = json.loads(r[18]) if r[18] else {}
            except Exception:
                rt = {}
            items.append({
                "id": r[0],
                "created_at": r[1],
                "fecha_corte": r[2],
                "modulo": r[3],
                "modulo_label": r[4],
                "nombre_archivo": r[5],
                "origen": r[6],
                "usuario_email": r[7],
                "cartera_activa_usd": r[8],
                "total_mora_usd": r[9],
                "pendiente_tarifa_usd": r[10],
                "pendiente_cuota_usd": r[11],
                "pendiente_mora_usd": r[12],
                "porcentaje_mora": r[13],
                "total_cuotas": r[14],
                "total_clientes": r[15],
                "clientes_con_telefono": r[16],
                "cobertura_contacto_pct": r[17],
                "resumen_tramos": rt,
            })
        conn.close()

    # Para el resumen comparativo global, si se filtró por módulo, usamos items o consultamos todos
    all_items = items
    if modulo in ("general", "avila") and source == "supabase":
        try:
            all_rows = _supabase_request(
                "GET",
                f"historial_cargas?select={select_cols}&order=created_at.desc&limit=100",
                timeout=8,
            )
            if all_rows is not None:
                all_items = all_rows
        except Exception:
            pass

    return {
        "source": source,
        "total_records": len(items),
        "summary": {
            "general": _compute_module_comparison(all_items, "general"),
            "avila": _compute_module_comparison(all_items, "avila"),
            "total_snapshots": len(all_items),
        },
        "items": items,
    }


def get_history_detail(item_id: int) -> Optional[dict]:
    """Obtiene el detalle completo de una carga histórica (incluyendo Tramos y Fibex)."""
    try:
        rows = _supabase_request(
            "GET",
            f"historial_cargas?id=eq.{int(item_id)}&limit=1",
            timeout=8,
        )
        if rows and len(rows) > 0:
            return rows[0]
    except Exception:
        pass

    ensure_db_table()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT * FROM historial_cargas WHERE id = ?", (int(item_id),))
    r = cur.fetchone()
    conn.close()
    if not r:
        return None
    return {
        "id": r[0],
        "created_at": r[1],
        "fecha_corte": r[2],
        "modulo": r[3],
        "modulo_label": r[4],
        "nombre_archivo": r[5],
        "origen": r[6],
        "usuario_email": r[7],
        "cartera_activa_usd": r[8],
        "total_mora_usd": r[9],
        "pendiente_tarifa_usd": r[10],
        "pendiente_cuota_usd": r[11],
        "pendiente_mora_usd": r[12],
        "porcentaje_mora": r[13],
        "total_cuotas": r[14],
        "total_clientes": r[15],
        "clientes_con_telefono": r[16],
        "cobertura_contacto_pct": r[17],
        "resumen_tramos": json.loads(r[18]) if r[18] else {},
        "resumen_fibex": json.loads(r[19]) if r[19] else {},
        "filtros_aplicados": json.loads(r[20]) if r[20] else {},
    }


def delete_history_item(item_id: int) -> bool:
    """Elimina un registro del historial en Supabase y SQLite."""
    deleted = False
    try:
        _supabase_request(
            "DELETE",
            f"historial_cargas?id=eq.{int(item_id)}",
            prefer="return=minimal",
            timeout=8,
        )
        deleted = True
    except Exception as e:
        print("Aviso al eliminar historial en Supabase:", e)

    try:
        ensure_db_table()
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("DELETE FROM historial_cargas WHERE id = ?", (int(item_id),))
        conn.commit()
        conn.close()
        deleted = True
    except Exception:
        pass
    return deleted
