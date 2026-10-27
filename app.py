import os
from typing import Optional
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles

from core.analytics import (
    build_dashboard_payload,
    build_debtor_matched_dataset,
    calculate_fibex_summary,
    calculate_global_tramos_summary,
    calculate_totals,
    enrich_and_merge_datasets,
    get_metadata,
    ingest_uploaded_file,
)
from core.config import (
    BASE_DIR,
    DATA_STORE,
    DEFAULT_CONTACTS_PATH,
    DEFAULT_MORA_PATH,
    FilterRequest,
    SetCarteraActivaRequest,
)
from core.database import load_contacts_from_db
from core.parsers import (
    parse_bytes_to_dataframe,
    safe_read_bytes,
    standardize_cartera_dataframe,
    standardize_contacts_dataframe,
)
from core.reports import (
    generate_campanas_matrix_excel,
    generate_campanas_matrix_pdf,
    generate_contacts_excel,
    generate_fibex_excel,
    generate_fibex_pdf,
    generate_filtered_excel,
    generate_filtered_pdf,
    generate_tramos_excel,
    generate_tramos_pdf,
)

# -------------------------------------------------------------
# FASTAPI APP & STATIC MOUNT
# -------------------------------------------------------------

app = FastAPI(title="Motor de Reestructuracion, Segmentacion y Cruce de Cartera")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC_DIR = os.path.join(BASE_DIR, "static")
if os.path.isdir(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Cargar automáticamente la base de datos maestra al iniciar el servidor
try:
    load_contacts_from_db()
except Exception as _db_err:
    print("Aviso al iniciar base de datos SQLite:", _db_err)


# -------------------------------------------------------------
# RUTAS DE LA API PRINCIPAL
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
    payload = build_dashboard_payload(active_df)

    return {
        "loaded": cartera_loaded,
        "contacts_loaded": contacts_loaded,
        "cartera_filename": DATA_STORE["cartera_filename"],
        "contacts_filename": DATA_STORE["contacts_filename"],
        **payload,
        "sample_available": sample_available,
        "contacts_sample_available": contacts_sample_available
    }


@app.post("/api/load-sample")
def load_sample():
    """Carga el archivo de mora de ejemplo y lo cruza instantáneamente con la Base de Datos Maestra SQLite."""
    if not os.path.exists(DEFAULT_MORA_PATH):
        raise HTTPException(status_code=404, detail="Archivo mora.csv no encontrado en el servidor.")

    try:
        cartera_bytes = safe_read_bytes(DEFAULT_MORA_PATH)
        df_cartera = parse_bytes_to_dataframe(cartera_bytes, "mora.csv")
        df_cartera = standardize_cartera_dataframe(df_cartera)
        DATA_STORE["df_cartera_raw"] = df_cartera
        DATA_STORE["cartera_filename"] = "mora.csv (Cartera Principal)"
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No se pudo leer mora.csv: {e}")

    if DATA_STORE["df_contacts_raw"] is None or DATA_STORE["df_contacts_raw"].empty:
        load_contacts_from_db()

    total_db_contacts = len(DATA_STORE["df_contacts_raw"]) if DATA_STORE["df_contacts_raw"] is not None else 0
    DATA_STORE["contacts_filename"] = f"Base de Datos Activa ({total_db_contacts:,} clientes)"
    contacts_msg = f" + {total_db_contacts:,} clientes en Base de Datos."

    enrich_and_merge_datasets()
    payload = build_dashboard_payload(DATA_STORE["df_original"])

    return {
        "message": f"Cartera cargada exitosamente: {len(df_cartera):,} cuotas{contacts_msg}",
        **payload
    }


@app.post("/api/upload")
async def upload_cartera(file: UploadFile = File(...)):
    """Sube el archivo de cartera (CSV o Excel). Si se detecta que es contactos o campañas, actualiza la Base de Datos SQLite."""
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
    try:
        return ingest_uploaded_file(contents, file.filename, preferred_role="cartera")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error al procesar archivo: {str(e)}")


@app.post("/api/contacts/upload")
async def upload_contacts(file: UploadFile = File(...)):
    """Sube un archivo de contactos o campañas para sumar/actualizar la Base de Datos SQLite."""
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
    try:
        return ingest_uploaded_file(contents, file.filename, preferred_role="contacts")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error al procesar archivo de contactos: {str(e)}")


@app.post("/api/process")
def process_data(req: FilterRequest):
    """Aplica filtros dinamicos de tramos, campanas, tipos y agrupa los resultados. Solo requiere el archivo de Cartera."""
    if DATA_STORE["df_cartera_raw"] is None:
        raise HTTPException(
            status_code=400,
            detail="Falta archivo obligatorio: Archivo de Cartera / Mora (cuotas y saldos). Debes subir el archivo de mora para poder procesar."
        )

    if DATA_STORE["df_contacts_raw"] is None or DATA_STORE["df_contacts_raw"].empty:
        load_contacts_from_db()
    if DATA_STORE["df_original"] is None:
        enrich_and_merge_datasets()

    df_master = DATA_STORE["df_original"].copy()

    if req.cartera_activa is not None:
        DATA_STORE["cartera_activa_manual"] = req.cartera_activa

    payload = build_dashboard_payload(df_master, req.cartera_activa)

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
        r["wa_link"] = f"https://wa.me/{tel1}" if tel1 and len(tel1) >= 10 else ""
        preview.append(r)

    return {
        "kpis": kpis,
        "resumen_global_tramos": payload["resumen_global_tramos"],
        "resumen_fibex": payload["resumen_fibex"],
        "matriz_campanas_tramos": payload["matriz_campanas_tramos"],
        "agrupacion": payload["resumen_global_tramos"]["tramos"],
        "group_by_used": "Tramo Oficial de Mora",
        "preview_count": len(preview),
        "total_filtered": len(df),
        "preview": preview
    }


@app.post("/api/tramos/set-cartera-activa")
def set_cartera_activa_endpoint(req: SetCarteraActivaRequest):
    """Guarda o actualiza el monto manual de Cartera Activa y recalcula el resumen de tramos y fibex."""
    DATA_STORE["cartera_activa_manual"] = req.cartera_activa
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is not None and not active_df.empty:
        return {
            "status": "ok",
            "cartera_activa": req.cartera_activa,
            "summary": calculate_global_tramos_summary(active_df, req.cartera_activa),
            "fibex_summary": calculate_fibex_summary(active_df, req.cartera_activa)
        }
    return {
        "status": "ok",
        "cartera_activa": req.cartera_activa,
        "summary": None,
        "fibex_summary": None
    }


# -------------------------------------------------------------
# RUTAS DE EXPORTACIÓN (EXCEL Y PDF)
# -------------------------------------------------------------

@app.get("/api/download/excel")
def download_excel():
    return generate_filtered_excel()


@app.get("/api/download/pdf")
def download_pdf():
    return generate_filtered_pdf()


@app.get("/api/tramos/download/excel")
def download_tramos_excel(cartera_activa: Optional[float] = None):
    return generate_tramos_excel(cartera_activa)


@app.get("/api/tramos/download/pdf")
def download_tramos_pdf(cartera_activa: Optional[float] = None):
    return generate_tramos_pdf(cartera_activa)


@app.get("/api/fibex/download/excel")
def download_fibex_excel(cartera_activa: Optional[float] = None):
    return generate_fibex_excel(cartera_activa)


@app.get("/api/fibex/download/pdf")
def download_fibex_pdf(cartera_activa: Optional[float] = None):
    return generate_fibex_pdf(cartera_activa)


@app.get("/api/campanas-tramos/download/excel")
def download_campanas_tramos_excel():
    return generate_campanas_matrix_excel()


@app.get("/api/campanas-tramos/download/pdf")
def download_campanas_tramos_pdf():
    return generate_campanas_matrix_pdf()


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
    return generate_contacts_excel()


@app.get("/health")
def health_check():
    return {"status": "ok", "service": "cobranzapp"}


@app.get("/", response_class=HTMLResponse)
def get_dashboard():
    with open(os.path.join(BASE_DIR, "templates", "index.html"), "r", encoding="utf-8") as f:
        return f.read()


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")
    uvicorn.run("app:app", host=host, port=port)
