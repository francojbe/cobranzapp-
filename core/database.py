import sqlite3
from datetime import datetime
import pandas as pd
from core.config import DB_PATH, DATA_STORE


def ensure_db_table():
    """Garantiza que exista la tabla master_clientes en SQLite."""
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
    conn.commit()
    conn.close()


def load_contacts_from_db() -> pd.DataFrame:
    """Lee la tabla master_clientes de SQLite a memoria y actualiza DATA_STORE."""
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
        if not DATA_STORE["contacts_filename"]:
            DATA_STORE["contacts_filename"] = f"Base de Datos Activa ({len(df_db):,} clientes)"
    return df_db


def upsert_contacts_to_db(df: pd.DataFrame) -> dict:
    """Inserta clientes nuevos o actualiza campos no vacíos en la base de datos SQLite sin perder información previa."""
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

    # Refrescar DATA_STORE desde la base de datos actualizada
    load_contacts_from_db()

    return {
        "processed": len(rows),
        "added": max(0, count_after - count_before),
        "total_in_db": count_after
    }
