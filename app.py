import os
from typing import Optional
from fastapi import FastAPI, File, HTTPException, Request, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from core.analytics import (
    build_dashboard_payload,
    build_debtor_matched_dataset,
    calculate_3_tramos_summary,
    calculate_campanas_tramos_matrix,
    calculate_fibex_summary,
    parse_cartera_activa_campanas_file,
    calculate_global_tramos_summary,
    calculate_totals,
    detect_cartera_module,
    enrich_and_merge_datasets,
    get_metadata,
    get_modules_summary,
    ingest_uploaded_file,
)
from core.config import (
    BASE_DIR,
    CreateUserRequest,
    DATA_STORE,
    DEFAULT_CONTACTS_PATH,
    DEFAULT_MORA_AVILA_PATH,
    DEFAULT_MORA_PATH,
    FilterRequest,
    LoginRequest,
    SelectModuleRequest,
    SetCarteraActivaCampanasRequest,
    SetCarteraActivaRequest,
    switch_active_module,
    sync_active_module_to_slot,
)
from core.database import (
    authenticate_user,
    create_or_update_user,
    create_session_token,
    delete_history_item,
    ensure_supabase_default_users,
    get_history_detail,
    get_history_list,
    list_users,
    load_cartera_activa_campanas,
    load_contacts_from_db,
    save_cartera_activa_campanas,
    save_history_snapshot,
    verify_session_token,
)
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
    generate_tramos_3_excel,
    generate_tramos_3_pdf,
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

# Cargar automáticamente la base de datos maestra y usuarios al iniciar el servidor
try:
    load_contacts_from_db()
    ensure_supabase_default_users()
except Exception as _db_err:
    print("Aviso al iniciar base de datos:", _db_err)


def _get_current_user(request: Request) -> Optional[dict]:
    token = request.cookies.get("cobranzapp_session")
    return verify_session_token(token)


# -------------------------------------------------------------
# RUTAS DE AUTENTICACIÓN Y USUARIOS (SUPABASE)
# -------------------------------------------------------------

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    user = _get_current_user(request)
    if user:
        return RedirectResponse(url="/", status_code=302)
    with open(os.path.join(BASE_DIR, "templates", "login.html"), "r", encoding="utf-8") as f:
        return f.read()


@app.post("/api/auth/login")
def api_login(req: LoginRequest, response: Response):
    user = authenticate_user(req.email, req.password)
    if not user:
        raise HTTPException(
            status_code=401,
            detail="Correo o contraseña incorrectos. Verifica tus credenciales.",
        )
    token = create_session_token(user)
    response.set_cookie(
        key="cobranzapp_session",
        value=token,
        max_age=86400 * 7,
        httponly=True,
        samesite="lax",
        path="/",
    )
    return {"status": "ok", "user": user}


@app.get("/logout")
def direct_logout():
    resp = RedirectResponse(url="/login", status_code=302)
    resp.delete_cookie(key="cobranzapp_session", path="/")
    return resp


@app.get("/api/auth/logout")
@app.post("/api/auth/logout")
def api_logout(response: Response):
    response.delete_cookie(key="cobranzapp_session", path="/")
    return {"status": "ok"}


@app.get("/api/auth/me")
def api_me(request: Request):
    user = _get_current_user(request)
    if not user:
        return {"authenticated": False, "user": None}
    return {"authenticated": True, "user": user}


@app.get("/api/auth/users")
def api_list_users(request: Request):
    user = _get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="No autorizado")
    return {"users": list_users()}


@app.post("/api/auth/users")
def api_create_user(req: CreateUserRequest, request: Request):
    user = _get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="No autorizado")
    if len(req.password or "") < 6:
        raise HTTPException(status_code=400, detail="La contraseña debe tener al menos 6 caracteres.")
    created = create_or_update_user(req.email, req.nombre, req.password, req.rol)
    return {"status": "ok", "user": created, "users": list_users()}


# -------------------------------------------------------------
# RUTAS DE HISTÓRICO DE MORA (SUPABASE public.historial_cargas)
# -------------------------------------------------------------

@app.get("/api/history")
def api_get_history(modulo: Optional[str] = None, limit: int = 100):
    mod_filter = modulo if modulo in ("general", "avila") else None
    return get_history_list(modulo=mod_filter, limit=limit)


@app.get("/api/history/{item_id}")
def api_get_history_detail(item_id: int):
    detail = get_history_detail(item_id)
    if not detail:
        raise HTTPException(status_code=404, detail="Registro histórico no encontrado.")
    return detail


@app.post("/api/history/snapshot")
def api_save_manual_snapshot(request: Request, overwrite: bool = True):
    """Permite guardar manualmente en Supabase una instantánea de la cartera actualmente cargada."""
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is None or active_df.empty:
        raise HTTPException(status_code=400, detail="No hay ninguna cartera cargada en memoria para guardar en el historial.")

    user = _get_current_user(request)
    user_email = user.get("email") if user else "admin@credix.com"
    fname = DATA_STORE.get("cartera_filename") or "corte_manual.csv"
    active_mod = detect_cartera_module(active_df, fname)
    ca_manual = DATA_STORE.get("cartera_activa_manual")
    payload = build_dashboard_payload(active_df, ca_manual)

    saved = save_history_snapshot(
        modulo=active_mod,
        nombre_archivo=fname,
        origen="web_manual",
        usuario_email=user_email,
        cartera_activa_usd=ca_manual,
        kpis=payload["totals"],
        resumen_tramos=payload["resumen_global_tramos"],
        resumen_fibex=payload["resumen_fibex"],
        filtros=DATA_STORE.get("filter_info") or {},
        overwrite_existing=overwrite,
    )
    return {"status": "ok", "snapshot": saved}


@app.delete("/api/history/{item_id}")
def api_delete_history(item_id: int):
    ok = delete_history_item(item_id)
    if not ok:
        raise HTTPException(status_code=404, detail="No se pudo eliminar el registro.")
    return {"status": "ok", "deleted_id": item_id}


# -------------------------------------------------------------
# RUTAS DE LA API PRINCIPAL
# -------------------------------------------------------------

@app.get("/api/status")
def get_status(request: Request):
    user = _get_current_user(request)
    cartera_loaded = DATA_STORE["df_cartera_raw"] is not None
    contacts_loaded = DATA_STORE["df_contacts_raw"] is not None
    sample_available = os.path.exists(DEFAULT_MORA_PATH)
    avila_sample_available = os.path.exists(DEFAULT_MORA_AVILA_PATH)
    contacts_sample_available = os.path.exists(DEFAULT_CONTACTS_PATH)

    if not cartera_loaded and not contacts_loaded:
        return {
            "user": user,
            "loaded": False,
            "contacts_loaded": False,
            "active_module": DATA_STORE.get("active_module", "general"),
            "enable_mora_avila": bool(DATA_STORE.get("enable_mora_avila", True)),
            "modules_summary": get_modules_summary(),
            "sample_available": sample_available,
            "avila_sample_available": avila_sample_available,
            "contacts_sample_available": contacts_sample_available
        }

    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    payload = build_dashboard_payload(active_df)

    return {
        "user": user,
        "loaded": cartera_loaded,
        "contacts_loaded": contacts_loaded,
        "cartera_filename": DATA_STORE["cartera_filename"],
        "contacts_filename": DATA_STORE["contacts_filename"],
        **payload,
        "sample_available": sample_available,
        "avila_sample_available": avila_sample_available,
        "contacts_sample_available": contacts_sample_available
    }


@app.post("/api/module/select")
def select_module_endpoint(req: SelectModuleRequest):
    """Cambia el módulo activo ('general' o 'avila') conservando en memoria los datos de ambos."""
    mod_id = switch_active_module(req.module_id)
    cartera_loaded = DATA_STORE["df_cartera_raw"] is not None
    contacts_loaded = DATA_STORE["df_contacts_raw"] is not None
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    payload = build_dashboard_payload(active_df)
    mod_label = "Mora Ávila" if mod_id == "avila" else "Mora Ubi"
    return {
        "status": "ok",
        "message": f"Módulo activo cambiado a: {mod_label}",
        "loaded": cartera_loaded,
        "contacts_loaded": contacts_loaded,
        "cartera_filename": DATA_STORE["cartera_filename"],
        "contacts_filename": DATA_STORE["contacts_filename"],
        **payload,
    }


@app.post("/api/load-sample")
def load_sample(module_id: Optional[str] = None):
    """Carga el archivo de mora de ejemplo del módulo activo ('general' o 'avila') y lo cruza con la Base de Datos."""
    if module_id in ("general", "avila"):
        switch_active_module(module_id)

    active_mod = DATA_STORE.get("active_module", "general")
    if active_mod == "avila":
        sample_path = DEFAULT_MORA_AVILA_PATH
        sample_filename = "reporte_mora_2026-10-28.csv"
        display_name = "reporte_mora_2026-10-28.csv (Mora Ávila)"
    else:
        sample_path = DEFAULT_MORA_PATH
        sample_filename = "mora.csv"
        display_name = "mora.csv (Mora Ubi)"

    if not os.path.exists(sample_path):
        raise HTTPException(status_code=404, detail=f"Archivo de muestra {sample_filename} no encontrado en el servidor.")

    try:
        cartera_bytes = safe_read_bytes(sample_path)
        df_cartera = parse_bytes_to_dataframe(cartera_bytes, sample_filename)
        df_cartera = standardize_cartera_dataframe(df_cartera)
        DATA_STORE["df_cartera_raw"] = df_cartera
        DATA_STORE["cartera_filename"] = display_name
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No se pudo leer {sample_filename}: {e}")

    if DATA_STORE["df_contacts_raw"] is None or DATA_STORE["df_contacts_raw"].empty:
        load_contacts_from_db()

    total_db_contacts = len(DATA_STORE["df_contacts_raw"]) if DATA_STORE["df_contacts_raw"] is not None else 0
    DATA_STORE["contacts_filename"] = f"Base Supabase Activa ({total_db_contacts:,} clientes)"
    contacts_msg = f" + {total_db_contacts:,} clientes en Base Supabase."

    enrich_and_merge_datasets()
    payload = build_dashboard_payload(DATA_STORE["df_original"])
    mod_title = "Mora Ávila" if active_mod == "avila" else "Mora Ubi"

    return {
        "message": f"{mod_title} cargada exitosamente: {len(df_cartera):,} cuotas{contacts_msg}",
        **payload
    }


@app.post("/api/upload")
async def upload_cartera(file: UploadFile = File(...), module_id: Optional[str] = None):
    """Sube el archivo de cartera (CSV o Excel). Si se detecta que es contactos o campañas, actualiza la Base de Datos."""
    if module_id in ("general", "avila"):
        switch_active_module(module_id)
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
    """Sube un archivo de contactos o campañas para sumar/actualizar la Base de Datos en Supabase y SQLite."""
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
def process_data(req: FilterRequest, request: Request):
    """Aplica filtros dinamicos de tramos, campanas, tipos y agrupa los resultados. Registra automáticamente en el historial de Supabase."""
    if req.module_id in ("general", "avila"):
        switch_active_module(req.module_id)
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
    fname = DATA_STORE.get("cartera_filename") or "mora.csv"
    detected_mod = detect_cartera_module(df_master, fname)
    df_master["_Modulo"] = detected_mod

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
    sync_active_module_to_slot()
    kpis = calculate_totals(df)

    # Registrar automáticamente en el Histórico de Mora en Supabase (si hay conflicto de fecha, avisar sin duplicar)
    cutoff_conflict = None
    try:
        user = _get_current_user(request)
        origen_efectivo = req.origen or ("web" if user else "n8n")
        email_efectivo = user.get("email") if user else "automatizacion@n8n.credix"
        saved_res = save_history_snapshot(
            modulo=detected_mod,
            nombre_archivo=fname,
            origen=origen_efectivo,
            usuario_email=email_efectivo,
            cartera_activa_usd=DATA_STORE.get("cartera_activa_manual"),
            kpis=payload["totals"],
            resumen_tramos=payload["resumen_global_tramos"],
            resumen_fibex=payload["resumen_fibex"],
            filtros=req.dict(),
            overwrite_existing=False,
        )
        if isinstance(saved_res, dict) and saved_res.get("status") == "conflict":
            cutoff_conflict = saved_res.get("conflict")
    except Exception as _hist_err:
        print("Aviso al registrar instantánea en historial:", _hist_err)

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
        "active_module": DATA_STORE.get("active_module", "general"),
        "modules_summary": get_modules_summary(),
        "kpis": kpis,
        "resumen_global_tramos": payload["resumen_global_tramos"],
        "resumen_3_tramos": payload["resumen_3_tramos"],
        "resumen_fibex": payload["resumen_fibex"],
        "matriz_campanas_tramos": payload["matriz_campanas_tramos"],
        "agrupacion": payload["resumen_global_tramos"]["tramos"],
        "group_by_used": "Tramo Oficial de Mora",
        "preview_count": len(preview),
        "total_filtered": len(df),
        "preview": preview,
        "cutoff_conflict": cutoff_conflict,
    }


@app.post("/api/tramos/set-cartera-activa")
def set_cartera_activa_endpoint(req: SetCarteraActivaRequest, request: Request):
    """Guarda o actualiza el monto manual de Cartera Activa y recalcula el resumen de tramos y fibex."""
    DATA_STORE["cartera_activa_manual"] = req.cartera_activa
    try:
        active_mod = DATA_STORE.get("active_module", "general")
        curr_cfg = DATA_STORE.get("cartera_activa_campanas") or load_cartera_activa_campanas(active_mod)
        curr_cfg["global_cartera_activa"] = req.cartera_activa
        DATA_STORE["cartera_activa_campanas"] = save_cartera_activa_campanas(active_mod, curr_cfg)
    except Exception:
        pass
    sync_active_module_to_slot()
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is not None and not active_df.empty:
        tramos_sum = calculate_global_tramos_summary(active_df, req.cartera_activa)
        tramos_3_sum = calculate_3_tramos_summary(active_df, req.cartera_activa)
        fibex_sum = calculate_fibex_summary(active_df, req.cartera_activa)
        try:
            user = _get_current_user(request)
            save_history_snapshot(
                modulo=DATA_STORE.get("active_module", "general"),
                nombre_archivo=DATA_STORE.get("cartera_filename") or "mora.csv",
                origen="web" if user else "n8n",
                usuario_email=user.get("email") if user else "admin@credix.com",
                cartera_activa_usd=req.cartera_activa,
                kpis=calculate_totals(active_df),
                resumen_tramos=tramos_sum,
                resumen_fibex=fibex_sum,
                filtros=DATA_STORE.get("filter_info") or {},
            )
        except Exception:
            pass
        return {
            "status": "ok",
            "cartera_activa": req.cartera_activa,
            "summary": tramos_sum,
            "tramos_3_summary": tramos_3_sum,
            "fibex_summary": fibex_sum
        }
    return {
        "status": "ok",
        "cartera_activa": req.cartera_activa,
        "summary": None,
        "tramos_3_summary": None,
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


@app.get("/api/tramos-3/download/excel")
def download_tramos_3_excel(cartera_activa: Optional[float] = None):
    return generate_tramos_3_excel(cartera_activa)


@app.get("/api/tramos-3/download/pdf")
def download_tramos_3_pdf(cartera_activa: Optional[float] = None):
    return generate_tramos_3_pdf(cartera_activa)


@app.get("/api/fibex/download/excel")
def download_fibex_excel(cartera_activa: Optional[float] = None):
    return generate_fibex_excel(cartera_activa)


@app.get("/api/fibex/download/pdf")
def download_fibex_pdf(cartera_activa: Optional[float] = None):
    return generate_fibex_pdf(cartera_activa)


@app.post("/api/campanas-tramos/set-cartera-activa")
def set_cartera_activa_campanas_endpoint(req: SetCarteraActivaCampanasRequest):
    """Guarda los montos de Cartera Activa de Fin de Mes y del Día Anterior por campaña en Supabase y SQLite."""
    active_mod = req.module_id if req.module_id in ("general", "avila") else DATA_STORE.get("active_module", "general")
    if active_mod != DATA_STORE.get("active_module"):
        switch_active_module(active_mod)

    saved_cfg = save_cartera_activa_campanas(
        active_mod,
        {
            "label_fin_mes": req.label_fin_mes,
            "label_dia_anterior": req.label_dia_anterior,
            "campanas": req.campanas,
        },
    )
    DATA_STORE["cartera_activa_campanas"] = saved_cfg

    if req.sync_global_cartera:
        tot_da = round(sum(float(v.get("dia_anterior") or 0.0) for v in saved_cfg.get("campanas", {}).values() if isinstance(v, dict)), 2)
        if tot_da > 0:
            DATA_STORE["cartera_activa_manual"] = tot_da

    sync_active_module_to_slot()
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    matrix_data = calculate_campanas_tramos_matrix(active_df) if active_df is not None else None
    tramos_sum = calculate_global_tramos_summary(active_df, DATA_STORE.get("cartera_activa_manual")) if active_df is not None else None
    tramos_3_sum = calculate_3_tramos_summary(active_df, DATA_STORE.get("cartera_activa_manual")) if active_df is not None else None
    fibex_sum = calculate_fibex_summary(active_df, DATA_STORE.get("cartera_activa_manual")) if active_df is not None else None

    return {
        "status": "ok",
        "cartera_activa_campanas": saved_cfg,
        "cartera_activa_global": DATA_STORE.get("cartera_activa_manual"),
        "matriz_campanas_tramos": matrix_data,
        "resumen_global_tramos": tramos_sum,
        "resumen_3_tramos": tramos_3_sum,
        "resumen_fibex": fibex_sum,
    }


@app.post("/api/campanas-tramos/upload-cartera-activa")
async def upload_cartera_activa_campanas_endpoint(file: UploadFile = File(...), target_col: str = "auto"):
    """Permite subir el archivo de Contabilidad (Excel o CSV) para poblar automáticamente la Cartera Activa por Campaña."""
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="El archivo enviado está vacío.")
    parsed_map = parse_cartera_activa_campanas_file(contents, file.filename, target_col=target_col)
    if not parsed_map:
        raise HTTPException(status_code=400, detail="No se detectaron columnas numéricas de Cartera Activa por campaña en el archivo.")

    active_mod = DATA_STORE.get("active_module", "general")
    curr_cfg = DATA_STORE.get("cartera_activa_campanas") or load_cartera_activa_campanas(active_mod)
    merged_camps = dict(curr_cfg.get("campanas") or {})
    for camp_name, vals in parsed_map.items():
        prev = dict(merged_camps.get(camp_name) or {"fin_mes": 0.0, "dia_anterior": 0.0})
        if "fin_mes" in vals:
            prev["fin_mes"] = vals["fin_mes"]
        if "dia_anterior" in vals:
            prev["dia_anterior"] = vals["dia_anterior"]
        merged_camps[camp_name] = prev

    curr_cfg["campanas"] = merged_camps
    saved_cfg = save_cartera_activa_campanas(active_mod, curr_cfg)
    DATA_STORE["cartera_activa_campanas"] = saved_cfg
    sync_active_module_to_slot()

    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    matrix_data = calculate_campanas_tramos_matrix(active_df) if active_df is not None else None
    return {
        "status": "ok",
        "matched_campaigns": len(parsed_map),
        "cartera_activa_campanas": saved_cfg,
        "matriz_campanas_tramos": matrix_data,
    }


@app.get("/api/campanas-tramos/download/excel")
def download_campanas_tramos_excel(mode: Optional[str] = "5tramos"):
    return generate_campanas_matrix_excel(mode=mode or "5tramos")


@app.get("/api/campanas-tramos/download/pdf")
def download_campanas_tramos_pdf(mode: Optional[str] = "5tramos"):
    return generate_campanas_matrix_pdf(mode=mode or "5tramos")


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
def get_dashboard(request: Request):
    user = _get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    with open(os.path.join(BASE_DIR, "templates", "index.html"), "r", encoding="utf-8") as f:
        return f.read()


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")
    uvicorn.run("app:app", host=host, port=port)
