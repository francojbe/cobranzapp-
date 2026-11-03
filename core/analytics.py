from datetime import datetime
from typing import Optional
import pandas as pd
from core.config import DATA_STORE, OFFICIAL_3_TRAMOS_CONFIG, OFFICIAL_TRAMOS_CONFIG, switch_active_module, sync_active_module_to_slot
from core.database import check_existing_cutoff_snapshot, load_cartera_activa_campanas, load_contacts_from_db, save_history_snapshot, upsert_contacts_to_db
from core.parsers import (
    assign_tramo_3_mora,
    assign_tramo_mora,
    detect_file_type,
    parse_bytes_to_dataframe,
    standardize_cartera_dataframe,
    standardize_contacts_dataframe,
)


def detect_cartera_module(df: pd.DataFrame, filename: str = "") -> str:
    """Detecta automáticamente si un DataFrame de cartera pertenece a Mora Ubi ('general') o Mora Ávila ('avila')."""
    fname = (filename or "").lower()
    if "avila" in fname or "ávila" in fname:
        return "avila"
    if "ubi" in fname:
        return "general"

    if df is not None and not df.empty:
        if "Programa" in df.columns:
            progs = " ".join(df["Programa"].dropna().astype(str).unique()).upper()
            if "AVILA" in progs or "ÁVILA" in progs:
                return "avila"
            if "CREDIX" in progs or "FIBEX" in progs:
                return "general"

        tarifa_sum = float(df["Pendiente Tarifa"].sum()) if "Pendiente Tarifa" in df.columns else 0.0
        cuota_sum = float(df["Pendiente Cuota"].sum()) if "Pendiente Cuota" in df.columns else 0.0

        if cuota_sum > tarifa_sum and cuota_sum > 0:
            return "avila"
        if tarifa_sum > cuota_sum and tarifa_sum > 0:
            return "general"

    if "reporte_mora" in fname:
        return "avila"

    return DATA_STORE.get("active_module", "general")


def resolve_active_debt_df(df: pd.DataFrame) -> pd.DataFrame:
    """Resuelve la columna monetaria activa (_Tot) y consolida Tramo y Campaña por cliente único (Cedula_Key):
    - En Mora Ávila (cuando Pendiente Cuota supera a Pendiente Tarifa), utiliza la suma total de la deuda
      (Pendiente Tarifa + Pendiente Cuota + Pendiente Mora).
    - En Mora Ubi (Cartera General), utiliza exclusivamente Pendiente Tarifa y filtra filas con deuda > 0."""
    df = df.copy()

    tarifa_sum = float(df["Pendiente Tarifa"].sum()) if "Pendiente Tarifa" in df.columns else 0.0
    cuota_sum = float(df["Pendiente Cuota"].sum()) if "Pendiente Cuota" in df.columns else 0.0

    # Si Pendiente Tarifa supera ampliamente a Pendiente Cuota (firma financiera de Mora Ubi),
    # jamás debe calcularse como Mora Ávila aunque el usuario haya subido el archivo en la pestaña equivocada.
    if tarifa_sum > cuota_sum * 1.5 and tarifa_sum > 0:
        is_avila = False
    elif cuota_sum > tarifa_sum and cuota_sum > 0:
        is_avila = True
    elif "_Modulo" in df.columns:
        is_avila = bool((df["_Modulo"] == "avila").any())
    else:
        is_avila = (DATA_STORE.get("active_module") == "avila")

    if is_avila and all(c in df.columns for c in ["Pendiente Tarifa", "Pendiente Cuota", "Pendiente Mora"]):
        df["_Tot"] = df["Pendiente Tarifa"] + df["Pendiente Cuota"] + df["Pendiente Mora"]
    elif "Pendiente Tarifa" in df.columns and tarifa_sum > 0:
        df["_Tot"] = df["Pendiente Tarifa"]
    elif "Total" in df.columns:
        df["_Tot"] = df["Total"]
    elif "Pendiente Cuota" in df.columns and "Pendiente Mora" in df.columns and (cuota_sum + float(df["Pendiente Mora"].sum())) > 0:
        df["_Tot"] = df["Pendiente Cuota"] + df["Pendiente Mora"]
    else:
        df["_Tot"] = 0.0

    df_act = df[df["_Tot"] > 0].copy() if (df["_Tot"] > 0).any() else df.copy()

    # Consolidar Tramo_Mora y Campaña a nivel de Cédula Única sobre las cuotas activas,
    # garantizando que ningún cliente se duplique entre distintos tramos ni entre distintas campañas.
    if not df_act.empty and "Cedula_Key" in df_act.columns:
        if "Dias de Mora" in df_act.columns:
            max_dias_cliente = df_act.groupby("Cedula_Key")["Dias de Mora"].transform("max")
            df_act["Tramo_Mora"] = max_dias_cliente.apply(assign_tramo_mora)
            df_act["Tramo_3_Mora"] = max_dias_cliente.apply(assign_tramo_3_mora)
        if "Tipo_Campana" in df_act.columns:
            df_act["Tipo_Campana"] = df_act.groupby("Cedula_Key")["Tipo_Campana"].transform("first")
        if "Programa" in df_act.columns:
            df_act["Programa"] = df_act.groupby("Cedula_Key")["Programa"].transform("first")

    return df_act


def enrich_and_merge_datasets():
    """Cruza y unifica la cartera con la base de contactos de forma instantanea."""
    if DATA_STORE["df_cartera_raw"] is None:
        sync_active_module_to_slot()
        return

    df_mora = DATA_STORE["df_cartera_raw"].copy()
    df_mora["_Modulo"] = DATA_STORE.get("active_module", "general")
    df_contacts = DATA_STORE["df_contacts_raw"]

    if df_contacts is not None and not df_contacts.empty:
        cols_to_merge = [
            "Cedula_Key", "Tipo_Campana", "Programa_Contacto",
            "Telefono_1", "Telefono_2", "Telefono_3", "Correo", "Nombre_Contacto"
        ]
        avail_cols = [c for c in cols_to_merge if c in df_contacts.columns]
        contacts_dedup = df_contacts[avail_cols].drop_duplicates(subset=["Cedula_Key"], keep="last")

        merged = pd.merge(
            df_mora,
            contacts_dedup,
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

    if not merged.empty and "Cedula_Key" in merged.columns:
        merged["Tipo_Campana"] = merged.groupby("Cedula_Key")["Tipo_Campana"].transform("first")
        merged["Programa"] = merged.groupby("Cedula_Key")["Programa"].transform("first")
        max_dias = merged.groupby("Cedula_Key")["Dias de Mora"].transform("max")
        merged["Tramo_Mora"] = max_dias.apply(assign_tramo_mora)
        merged["Tramo_3_Mora"] = max_dias.apply(assign_tramo_3_mora)

    DATA_STORE["df_original"] = merged
    DATA_STORE["df_filtered"] = merged
    build_debtor_matched_dataset(merged)
    sync_active_module_to_slot()


def build_debtor_matched_dataset(df: pd.DataFrame):
    """Consolida la cartera a nivel deudor (1 fila por deudor) para el modulo Call Center."""
    if df is None or df.empty:
        DATA_STORE["df_matched"] = None
        return

    df_active = resolve_active_debt_df(df)

    debtors = df_active.groupby("Cedula_Key", as_index=False).agg(
        Cedula=("Cedula Cliente", "first"),
        Nombre_Cliente=("Nombre Cliente", "first"),
        Cuotas_Vencidas=("Cedula Cliente", "count"),
        Monto_Total_USD=("_Tot", "sum"),
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

    df_active = resolve_active_debt_df(df)

    total_cuotas = len(df_active)
    total_clientes = int(df_active["Cedula Cliente"].nunique()) if "Cedula Cliente" in df_active.columns else total_cuotas
    total_creditos = int(df_active["Identificador de credito"].nunique()) if "Identificador de credito" in df_active.columns else 0
    p_tarifa = float(df_active["Pendiente Tarifa"].sum()) if "Pendiente Tarifa" in df_active.columns else 0.0
    p_cuota = float(df_active["Pendiente Cuota"].sum()) if "Pendiente Cuota" in df_active.columns else 0.0
    p_mora = float(df_active["Pendiente Mora"].sum()) if "Pendiente Mora" in df_active.columns else 0.0
    monto_total = float(df_active["_Tot"].sum())

    promedio = monto_total / total_clientes if total_clientes > 0 else 0.0

    # Metricas de contactabilidad sobre la cartera activa
    has_tel = (df_active["Telefono_1"].fillna("") != "") | (df_active["Telefono_2"].fillna("") != "") | (df_active["Telefono_3"].fillna("") != "")
    has_mail = df_active["Correo"].fillna("") != ""
    has_contact = has_tel | has_mail

    clientes_con_tel = df_active[has_tel]["Cedula Cliente"].nunique() if "Cedula Cliente" in df_active.columns else int(has_tel.sum())
    clientes_con_mail = df_active[has_mail]["Cedula Cliente"].nunique() if "Cedula Cliente" in df_active.columns else int(has_mail.sum())

    monto_contactable = float(df_active.loc[has_contact, "_Tot"].sum())

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
            s_vals = df["Fecha Vencimiento"].dropna().astype(str).str.strip()
            s_vals = s_vals[s_vals != ""]
            if not s_vals.empty:
                sample_val = s_vals.iloc[0]
                if len(sample_val) >= 10 and sample_val[4:5] == "-":
                    dates = pd.to_datetime(s_vals, format="%Y-%m-%d", errors="coerce").dropna()
                else:
                    dates = pd.to_datetime(s_vals, dayfirst=True, errors="coerce").dropna()
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

    df_active = resolve_active_debt_df(df)

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


def calculate_3_tramos_summary(df: pd.DataFrame, cartera_activa: Optional[float] = None):
    """Calcula el Resumen Ejecutivo en 3 Categorías Contables (Irregular 1-15d, Vencido Gestión 16-30d, Vencido contable +30d)."""
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

    df_active = resolve_active_debt_df(df)

    total_cartera_usd = round(float(df_active["_Tot"].sum()), 2)
    total_cuotas = len(df_active)
    total_clientes = int(df_active["Cedula Cliente"].nunique()) if "Cedula Cliente" in df_active.columns else total_cuotas

    target_ca = cartera_activa if (cartera_activa is not None and cartera_activa > 0) else DATA_STORE.get("cartera_activa_manual")
    denom = float(target_ca) if (target_ca is not None and target_ca > 0) else None
    total_pct = round((total_cartera_usd / denom * 100), 2) if denom else None

    tramos = []
    for t_cfg in OFFICIAL_3_TRAMOS_CONFIG:
        sub = df_active[df_active["Tramo_3_Mora"] == t_cfg["grupo"]]
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
            "pct_badge_class": t_cfg["pct_badge_class"],
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


def calculate_fibex_summary(df: pd.DataFrame, cartera_activa: Optional[float] = None):
    """Calcula el Resumen Comparativo de Clientes Web NO Fibex vs. Fibex (y Resto)
    sobre la cartera activa y la mora total."""
    if df is None or df.empty:
        return {
            "cutoff_date": datetime.now().strftime("%d-%m-%Y"),
            "total_cartera_usd": 0.0,
            "total_cuotas": 0,
            "total_clientes": 0,
            "cartera_activa_usd": None,
            "total_porcentaje": None,
            "segments": []
        }

    df_active = resolve_active_debt_df(df)

    total_cartera_usd = round(float(df_active["_Tot"].sum()), 2)
    total_cuotas = len(df_active)
    total_clientes = int(df_active["Cedula Cliente"].nunique()) if "Cedula Cliente" in df_active.columns else total_cuotas

    target_ca = cartera_activa if (cartera_activa is not None and cartera_activa > 0) else DATA_STORE.get("cartera_activa_manual")
    denom = float(target_ca) if (target_ca is not None and target_ca > 0) else None
    total_pct = round((total_cartera_usd / denom * 100), 2) if denom else None

    col_camp = "Tipo_Campana" if "Tipo_Campana" in df_active.columns else ("Programa" if "Programa" in df_active.columns else None)
    if col_camp:
        is_nofibex_mask = df_active[col_camp].fillna("").astype(str).str.upper().str.contains("NO FIBEX")
    else:
        is_nofibex_mask = pd.Series(False, index=df_active.index)

    sub_nofibex = df_active[is_nofibex_mask]
    sub_fibex = df_active[~is_nofibex_mask]

    def _build_segment(sub_df, seg_id, etiqueta, subtitulo, color, badge_class):
        monto = round(float(sub_df["_Tot"].sum()), 2)
        cuotas = len(sub_df)
        clientes = int(sub_df["Cedula Cliente"].nunique()) if "Cedula Cliente" in sub_df.columns else cuotas
        pct_mora = round((monto / total_cartera_usd * 100), 2) if total_cartera_usd > 0 else 0.0
        pct_activa = round((monto / denom * 100), 2) if denom else None

        tramos_dict = {}
        for t_cfg in OFFICIAL_TRAMOS_CONFIG:
            sub_t = sub_df[sub_df["Tramo_Mora"] == t_cfg["grupo"]]
            tramos_dict[t_cfg["id"]] = {
                "monto": round(float(sub_t["_Tot"].sum()), 2),
                "clientes": int(sub_t["Cedula Cliente"].nunique()) if "Cedula Cliente" in sub_t.columns else len(sub_t),
                "cuotas": len(sub_t)
            }

        return {
            "id": seg_id,
            "etiqueta": etiqueta,
            "subtitulo": subtitulo,
            "color": color,
            "badge_class": badge_class,
            "cuotas": cuotas,
            "clientes": clientes,
            "total_monto": monto,
            "porcentaje_mora": pct_mora,
            "porcentaje": pct_activa,
            "tramos": tramos_dict
        }

    seg_nofibex = _build_segment(
        sub_nofibex, "no_fibex", "Clientes Web NO Fibex",
        "22 sub-campañas regionales sin suscripción Fibex", "#f59e0b",
        "inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-50 text-amber-800 border border-amber-300 shadow-2xs"
    )
    seg_fibex = _build_segment(
        sub_fibex, "fibex_resto", "Clientes Fibex y Resto",
        "Clientes con suscripción Fibex, Pre-aprobados, Aliados y otros programas", "#3b82f6",
        "inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-blue-50 text-blue-800 border border-blue-300 shadow-2xs"
    )

    return {
        "cutoff_date": get_cutoff_date(df),
        "total_cartera_usd": total_cartera_usd,
        "total_cuotas": total_cuotas,
        "total_clientes": total_clientes,
        "cartera_activa_usd": denom,
        "total_porcentaje": total_pct,
        "segments": [seg_nofibex, seg_fibex]
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
    """Calcula la matriz bidimensional de Campañas vs 5 Tramos Oficiales y 3 Tramos Contables de Mora,
    incluyendo el % de morosidad contra Cartera Activa de Fin de Mes y del Día Anterior por campaña."""
    active_mod = DATA_STORE.get("active_module", "general")
    ca_cfg = DATA_STORE.get("cartera_activa_campanas")
    if not ca_cfg or not isinstance(ca_cfg, dict):
        ca_cfg = load_cartera_activa_campanas(active_mod)
        DATA_STORE["cartera_activa_campanas"] = ca_cfg

    label_fin_mes = str(ca_cfg.get("label_fin_mes") or "Fin de Mes")
    label_dia_anterior = str(ca_cfg.get("label_dia_anterior") or "Día Anterior")
    raw_camp_map = ca_cfg.get("campanas") or {}

    # Normalizar claves del mapa de cartera activa por campaña para cruce exacto e insensible a mayúsculas
    norm_camp_map = {}
    for k, v in raw_camp_map.items():
        if isinstance(v, dict):
            nk = normalize_campaign_name(k)
            norm_camp_map[nk.lower()] = v
            norm_camp_map[str(k).strip().lower()] = v

    if df is None or df.empty:
        return {
            "cutoff_date": datetime.now().strftime("%d-%m-%Y"),
            "total_cartera_usd": 0.0,
            "total_clientes": 0,
            "total_cuotas": 0,
            "label_fin_mes": label_fin_mes,
            "label_dia_anterior": label_dia_anterior,
            "total_ca_fin_mes": 0.0,
            "total_ca_dia_anterior": 0.0,
            "total_pct_fin_mes": None,
            "total_pct_dia_anterior": None,
            "total_variacion_pct": None,
            "matrix": [],
            "totals_by_tramo": {t["id"]: {"monto": 0.0, "clientes": 0, "cuotas": 0} for t in OFFICIAL_TRAMOS_CONFIG},
            "totals_by_tramo_3": {t["id"]: {"monto": 0.0, "clientes": 0, "cuotas": 0} for t in OFFICIAL_3_TRAMOS_CONFIG},
        }

    df_active = resolve_active_debt_df(df)

    total_cartera = float(df_active["_Tot"].sum())
    total_clientes_global = int(df_active["Cedula Cliente"].nunique()) if "Cedula Cliente" in df_active.columns else len(df_active)
    total_cuotas_global = len(df_active)

    col_campana = "Tipo_Campana" if "Tipo_Campana" in df_active.columns else "Programa"
    df_active["_Macro_Campana"] = df_active[col_campana].apply(normalize_campaign_name)

    matrix = []
    totals_by_tramo = {t["id"]: {"monto": 0.0, "clientes": 0, "cuotas": 0} for t in OFFICIAL_TRAMOS_CONFIG}
    totals_by_tramo_3 = {t["id"]: {"monto": 0.0, "clientes": 0, "cuotas": 0} for t in OFFICIAL_3_TRAMOS_CONFIG}

    # Totales globales por 5 tramos operativos
    for t_cfg in OFFICIAL_TRAMOS_CONFIG:
        sub_t = df_active[df_active["Tramo_Mora"] == t_cfg["grupo"]]
        totals_by_tramo[t_cfg["id"]]["monto"] = round(float(sub_t["_Tot"].sum()), 2)
        totals_by_tramo[t_cfg["id"]]["cuotas"] = len(sub_t)
        totals_by_tramo[t_cfg["id"]]["clientes"] = int(sub_t["Cedula Cliente"].nunique()) if "Cedula Cliente" in sub_t.columns else len(sub_t)

    # Totales globales por 3 tramos contables
    for t3_cfg in OFFICIAL_3_TRAMOS_CONFIG:
        sub_t3 = df_active[df_active["Tramo_3_Mora"] == t3_cfg["grupo"]]
        totals_by_tramo_3[t3_cfg["id"]]["monto"] = round(float(sub_t3["_Tot"].sum()), 2)
        totals_by_tramo_3[t3_cfg["id"]]["cuotas"] = len(sub_t3)
        totals_by_tramo_3[t3_cfg["id"]]["clientes"] = int(sub_t3["Cedula Cliente"].nunique()) if "Cedula Cliente" in sub_t3.columns else len(sub_t3)

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

        tramos_3_dict = {}
        for t3_cfg in OFFICIAL_3_TRAMOS_CONFIG:
            sub3 = grp[grp["Tramo_3_Mora"] == t3_cfg["grupo"]]
            m3 = float(sub3["_Tot"].sum())
            cl3 = int(sub3["Cedula Cliente"].nunique()) if "Cedula Cliente" in sub3.columns else len(sub3)
            cu3 = len(sub3)
            tramos_3_dict[t3_cfg["id"]] = {
                "monto": round(m3, 2),
                "clientes": cl3,
                "cuotas": cu3
            }

        ca_entry = norm_camp_map.get(camp.lower(), {})
        ca_fin_mes = round(float(ca_entry.get("fin_mes") or 0.0), 2)
        ca_dia_anterior = round(float(ca_entry.get("dia_anterior") or 0.0), 2)

        pct_fin_mes = round((camp_tot / ca_fin_mes) * 100, 2) if ca_fin_mes > 0 else None
        pct_dia_anterior = round((camp_tot / ca_dia_anterior) * 100, 2) if ca_dia_anterior > 0 else None
        variacion_pct = round(pct_dia_anterior - pct_fin_mes, 2) if (pct_fin_mes is not None and pct_dia_anterior is not None) else None

        pct = round(camp_tot / total_cartera * 100, 2) if total_cartera > 0 else 0.0
        matrix.append({
            "campana": camp,
            "total_monto": round(camp_tot, 2),
            "total_clientes": camp_cl,
            "total_cuotas": camp_cuotas,
            "porcentaje": pct,
            "ca_fin_mes": ca_fin_mes,
            "ca_dia_anterior": ca_dia_anterior,
            "pct_fin_mes": pct_fin_mes,
            "pct_dia_anterior": pct_dia_anterior,
            "variacion_pct": variacion_pct,
            "tramos": tramos_dict,
            "tramos_3": tramos_3_dict,
        })

    # Ordenar por deuda descendente (las campañas más críticas primero)
    matrix.sort(key=lambda x: x["total_monto"], reverse=True)

    total_ca_fin_mes = round(sum(float(v.get("fin_mes") or 0.0) for v in raw_camp_map.values() if isinstance(v, dict)), 2)
    total_ca_dia_anterior = round(sum(float(v.get("dia_anterior") or 0.0) for v in raw_camp_map.values() if isinstance(v, dict)), 2)
    if total_ca_dia_anterior <= 0 and DATA_STORE.get("cartera_activa_manual") and float(DATA_STORE["cartera_activa_manual"]) > 0:
        total_ca_dia_anterior = round(float(DATA_STORE["cartera_activa_manual"]), 2)

    total_pct_fin_mes = round((total_cartera / total_ca_fin_mes) * 100, 2) if total_ca_fin_mes > 0 else None
    total_pct_dia_anterior = round((total_cartera / total_ca_dia_anterior) * 100, 2) if total_ca_dia_anterior > 0 else None
    total_variacion_pct = (
        round(total_pct_dia_anterior - total_pct_fin_mes, 2)
        if (total_pct_fin_mes is not None and total_pct_dia_anterior is not None)
        else None
    )

    return {
        "cutoff_date": get_cutoff_date(df),
        "total_cartera_usd": round(total_cartera, 2),
        "total_clientes": total_clientes_global,
        "total_cuotas": total_cuotas_global,
        "label_fin_mes": label_fin_mes,
        "label_dia_anterior": label_dia_anterior,
        "total_ca_fin_mes": total_ca_fin_mes,
        "total_ca_dia_anterior": total_ca_dia_anterior,
        "total_pct_fin_mes": total_pct_fin_mes,
        "total_pct_dia_anterior": total_pct_dia_anterior,
        "total_variacion_pct": total_variacion_pct,
        "cartera_activa_campanas": ca_cfg,
        "matrix": matrix,
        "totals_by_tramo": totals_by_tramo,
        "totals_by_tramo_3": totals_by_tramo_3,
    }


def parse_cartera_activa_campanas_file(contents: bytes, filename: str, target_col: str = "auto") -> dict:
    """Extrae los montos de Cartera Activa por Campaña desde un archivo Excel o CSV de Contabilidad."""
    from core.parsers import clean_currency_series
    raw_df = parse_bytes_to_dataframe(contents, filename)
    if raw_df is None or raw_df.empty:
        return {}

    cols = list(raw_df.columns)
    camp_col = None
    for c in cols:
        cl = str(c).lower()
        if any(k in cl for k in ["campa", "segmento", "programa", "categoria", "categoría"]):
            camp_col = c
            break
    if camp_col is None:
        camp_col = cols[0]

    # Detectar columnas numéricas / monetarias candidatas
    num_cols = []
    for c in cols:
        if c == camp_col:
            continue
        cl = str(c).lower()
        if any(k in cl for k in ["cedula", "cédula", "id", "telefono", "teléfono", "dias", "días", "cuotas", "clientes", "cantidad"]):
            continue
        s_num = clean_currency_series(raw_df[c])
        if float(s_num.sum()) > 0:
            num_cols.append((c, s_num))

    if not num_cols:
        return {}

    df_work = pd.DataFrame()
    df_work["_Camp"] = raw_df[camp_col].astype(str).apply(normalize_campaign_name)

    result = {}
    if len(num_cols) >= 2 and target_col == "auto" and len(raw_df) <= 150:
        # El archivo trae ambas columnas (Fin de Mes y Día Anterior) resumidas por campaña
        df_work["_Val1"] = num_cols[0][1]
        df_work["_Val2"] = num_cols[1][1]
        grouped = df_work.groupby("_Camp", as_index=False).agg(_Val1=("_Val1", "sum"), _Val2=("_Val2", "sum"))
        for _, r in grouped.iterrows():
            c_name = str(r["_Camp"]).strip()
            if c_name and c_name.lower() not in ("total", "totales", "general"):
                result[c_name] = {
                    "fin_mes": round(float(r["_Val1"]), 2),
                    "dia_anterior": round(float(r["_Val2"]), 2),
                }
    else:
        # El archivo trae 1 columna de Cartera Activa (ej. archivo diario de Contabilidad o ESTATUS_CREDITOS)
        chosen_series = num_cols[0][1]
        for c_name, s_series in num_cols:
            if any(k in str(c_name).lower() for k in ["solicitado", "activa", "monto", "saldo", "cartera", "usd"]):
                chosen_series = s_series
                break
        df_work["_Val"] = chosen_series
        grouped = df_work.groupby("_Camp", as_index=False)["_Val"].sum()
        slot_key = "fin_mes" if target_col == "fin_mes" else "dia_anterior"
        for _, r in grouped.iterrows():
            c_name = str(r["_Camp"]).strip()
            if c_name and c_name.lower() not in ("total", "totales", "general"):
                result[c_name] = {slot_key: round(float(r["_Val"]), 2)}

    return result


def get_modules_summary() -> dict:
    """Retorna el estado ejecutivo resumido de cada módulo de cartera (Cartera General y Mora Ávila)."""
    sync_active_module_to_slot()
    summary = {}
    for mod_id, slot in DATA_STORE.get("modules", {}).items():
        df_mod = slot.get("df_original") if slot.get("df_original") is not None else slot.get("df_cartera_raw")
        if df_mod is not None and not df_mod.empty:
            df_act = resolve_active_debt_df(df_mod)
            monto = round(float(df_act["_Tot"].sum()), 2)
            cuotas = len(df_act)
            clientes = int(df_act["Cedula Cliente"].nunique()) if "Cedula Cliente" in df_act.columns else cuotas
            summary[mod_id] = {
                "id": mod_id,
                "label": slot.get("label", mod_id),
                "loaded": True,
                "filename": slot.get("cartera_filename", ""),
                "total_cuotas": cuotas,
                "total_clientes": clientes,
                "monto_total": monto,
                "cartera_activa": slot.get("cartera_activa_manual"),
            }
        else:
            summary[mod_id] = {
                "id": mod_id,
                "label": slot.get("label", mod_id),
                "loaded": False,
                "filename": "",
                "total_cuotas": 0,
                "total_clientes": 0,
                "monto_total": 0.0,
                "cartera_activa": slot.get("cartera_activa_manual"),
            }
    return summary


def get_metadata(df: pd.DataFrame):
    has_contacts = (DATA_STORE["df_contacts_raw"] is not None and not DATA_STORE["df_contacts_raw"].empty)
    total_contacts = len(DATA_STORE["df_contacts_raw"]) if has_contacts else 0
    active_mod = DATA_STORE.get("active_module", "general")
    enable_avila = bool(DATA_STORE.get("enable_mora_avila", True))

    if df is None or df.empty:
        return {
            "active_module": active_mod,
            "enable_mora_avila": enable_avila,
            "total_rows": 0, "total_clientes": 0, "programas": [],
            "tipos_campana": [], "tipos": [], "canales": [], "tramos": {},
            "min_dias": 0, "max_dias": 0, "has_contacts": has_contacts,
            "total_contacts": total_contacts,
            "cartera_filename": DATA_STORE["cartera_filename"],
            "contacts_filename": DATA_STORE["contacts_filename"]
        }

    df_act = resolve_active_debt_df(df)
    programas = sorted([p for p in df_act["Programa"].dropna().unique().tolist() if str(p).strip()]) if "Programa" in df_act.columns else []
    tipos_campana = sorted([t for t in df_act["Tipo_Campana"].dropna().unique().tolist() if str(t).strip()]) if "Tipo_Campana" in df_act.columns else []
    tipos = sorted([t for t in df_act["Tipo"].dropna().unique().tolist() if str(t).strip()]) if "Tipo" in df_act.columns else []
    canales = sorted([c for c in df_act["Canal"].dropna().unique().tolist() if str(c).strip()]) if "Canal" in df_act.columns else []

    tramos_counts = df_act["Tramo_Mora"].value_counts().to_dict() if "Tramo_Mora" in df_act.columns else {}

    min_dias = int(df_act["Dias de Mora"].min()) if "Dias de Mora" in df_act.columns and not df_act.empty else 0
    max_dias = int(df_act["Dias de Mora"].max()) if "Dias de Mora" in df_act.columns and not df_act.empty else 0

    return {
        "active_module": active_mod,
        "enable_mora_avila": enable_avila,
        "total_rows": len(df_act),
        "total_clientes": int(df_act["Cedula Cliente"].nunique()) if "Cedula Cliente" in df_act.columns else 0,
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


def build_dashboard_payload(df: Optional[pd.DataFrame], cartera_activa: Optional[float] = None) -> dict:
    """Construye el paquete estándar de indicadores del dashboard para evitar duplicación en los endpoints."""
    return {
        "active_module": DATA_STORE.get("active_module", "general"),
        "enable_mora_avila": bool(DATA_STORE.get("enable_mora_avila", True)),
        "modules_summary": get_modules_summary(),
        "metadata": get_metadata(df),
        "totals": calculate_totals(df),
        "resumen_global_tramos": calculate_global_tramos_summary(df, cartera_activa),
        "resumen_3_tramos": calculate_3_tramos_summary(df, cartera_activa),
        "resumen_fibex": calculate_fibex_summary(df, cartera_activa),
        "matriz_campanas_tramos": calculate_campanas_tramos_matrix(df)
    }


def ingest_uploaded_file(contents: bytes, filename: str, preferred_role: str = "cartera") -> dict:
    """Procesa un archivo subido (Cartera o Contactos/Campañas) con detección automática de tipo y de módulo (Ubi vs. Ávila)."""
    raw_df = parse_bytes_to_dataframe(contents, filename)
    detected = detect_file_type(raw_df)

    # Resolver rol efectivo según detección inteligente
    effective_role = detected if detected in ("cartera", "contacts") else preferred_role

    if effective_role == "contacts":
        df_contacts = standardize_contacts_dataframe(raw_df)
        db_stats = upsert_contacts_to_db(df_contacts)
        DATA_STORE["contacts_filename"] = f"Base de Datos Activa ({db_stats['total_in_db']:,} clientes)"

        enrich_and_merge_datasets()

        active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
        payload = build_dashboard_payload(active_df)

        res = {
            "message": (
                f"Base de Datos actualizada con '{filename}': "
                f"{db_stats['processed']:,} registros procesados ({db_stats['added']:,} clientes nuevos sumados). "
                f"Total en BD: {db_stats['total_in_db']:,} clientes."
            ),
            "total_contacts": db_stats["total_in_db"],
            "added_contacts": db_stats["added"],
            "filename": DATA_STORE["contacts_filename"],
            **payload
        }
        if preferred_role == "cartera":
            res["auto_routed"] = "contacts"
        return res

    # Rol efectivo: cartera
    df = standardize_cartera_dataframe(raw_df)

    # Auto-detectar si el archivo corresponde a Mora Ubi ('general') o Mora Ávila ('avila')
    # para cambiar automáticamente al módulo correcto y no mezclar fórmulas de cálculo
    detected_mod = detect_cartera_module(df, filename)
    if DATA_STORE.get("active_module") != detected_mod:
        switch_active_module(detected_mod)

    DATA_STORE["df_cartera_raw"] = df
    DATA_STORE["cartera_filename"] = filename

    if DATA_STORE["df_contacts_raw"] is None or DATA_STORE["df_contacts_raw"].empty:
        load_contacts_from_db()
    enrich_and_merge_datasets()

    active_df = DATA_STORE["df_original"]
    payload = build_dashboard_payload(active_df)
    mod_label = "Mora Ávila" if detected_mod == "avila" else "Mora Ubi"

    try:
        saved_res = save_history_snapshot(
            modulo=detected_mod,
            nombre_archivo=filename,
            origen="web",
            usuario_email="admin@credix.com",
            cartera_activa_usd=DATA_STORE.get("cartera_activa_manual"),
            kpis=payload.get("totals"),
            resumen_tramos=payload.get("resumen_global_tramos"),
            resumen_fibex=payload.get("resumen_fibex"),
            filtros=DATA_STORE.get("filter_info") or {},
            overwrite_existing=False,
        )
        if isinstance(saved_res, dict) and saved_res.get("status") == "conflict":
            payload["cutoff_conflict"] = saved_res.get("conflict")
    except Exception as _hist_err:
        print("Aviso al registrar instantánea al subir archivo:", _hist_err)

    if preferred_role == "contacts":
        return {
            "auto_routed": "cartera",
            "message": f"Detección inteligente: '{filename}' es el Archivo de Cartera ({mod_label}, {len(df):,} cuotas) y fue procesado automáticamente.",
            **payload
        }

    return {
        "message": f"Archivo de cartera '{filename}' ({mod_label}) procesado y cruzado con la Base de Datos con éxito.",
        **payload
    }
