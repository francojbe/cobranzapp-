import os
from typing import List, Optional
from pydantic import BaseModel

# Directorio raíz del proyecto (un nivel arriba de /core)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DEFAULT_MORA_PATH = os.path.join(BASE_DIR, "mora.csv")
DEFAULT_MORA_AVILA_PATH = os.path.join(BASE_DIR, "reporte_mora_2026-10-28.csv")
DEFAULT_CONTACTS_PATH = os.path.join(BASE_DIR, "Base de contactos.csv")
DEFAULT_CAMPANAS_PATH = os.path.join(BASE_DIR, "campañas.xlsx")
DB_PATH = os.path.join(BASE_DIR, "cobranzapp.db")

# Interruptor de licencia para el segundo módulo (Mora Ávila)
# En la rama 'version-basica', viene desactivado por defecto (Mora Ubi única)
ENABLE_MORA_AVILA = os.getenv("ENABLE_MORA_AVILA", "false").strip().lower() in ("true", "1", "yes", "si")

def _empty_module_slot(name: str, label: str) -> dict:
    return {
        "id": name,
        "label": label,
        "df_cartera_raw": None,
        "cartera_filename": "",
        "df_original": None,
        "df_filtered": None,
        "filter_info": {},
        "df_matched": None,
        "cartera_activa_manual": None,
    }

# Almacén en memoria del estado activo de la sesión (soporta múltiples módulos de mora)
DATA_STORE = {
    "active_module": "general",   # "general" | "avila"
    "enable_mora_avila": ENABLE_MORA_AVILA,
    "df_cartera_raw": None,
    "cartera_filename": "",
    "df_contacts_raw": None,
    "contacts_filename": "",
    "df_original": None,          # Enriquecida y unificada del módulo activo
    "df_filtered": None,          # Filtrada según criterios operativos del módulo activo
    "filter_info": {},
    "df_matched": None,           # Agregada a nivel deudor para Call Center del módulo activo
    "cartera_activa_manual": None,# Monto manual de cartera activa del módulo activo
    "modules": {
        "general": _empty_module_slot("general", "Mora Ubi"),
        "avila": _empty_module_slot("avila", "Mora Ávila"),
    }
}

_MODULE_KEYS = [
    "df_cartera_raw",
    "cartera_filename",
    "df_original",
    "df_filtered",
    "filter_info",
    "df_matched",
    "cartera_activa_manual",
]


def sync_active_module_to_slot():
    """Guarda el estado actual del módulo activo en su ranura independiente en memoria."""
    mod_id = DATA_STORE.get("active_module", "general")
    slot = DATA_STORE["modules"].get(mod_id)
    if slot is not None:
        for k in _MODULE_KEYS:
            slot[k] = DATA_STORE.get(k)


def switch_active_module(target_module: str) -> str:
    """Cambia el módulo activo ('general' o 'avila') preservando en memoria los datos de ambos."""
    if target_module not in ("general", "avila"):
        target_module = "general"
    if target_module == "avila" and not DATA_STORE.get("enable_mora_avila", True):
        target_module = "general"

    # 1. Guardar estado del módulo saliente
    sync_active_module_to_slot()

    # 2. Cargar estado del módulo entrante
    DATA_STORE["active_module"] = target_module
    target_slot = DATA_STORE["modules"][target_module]
    for k in _MODULE_KEYS:
        DATA_STORE[k] = target_slot.get(k)

    return target_module

# Configuración fija de los 5 Tramos Oficiales de Mora
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
    ],
    "Fecha Vencimiento": [
        r"^fechavencimiento$", r"^fechadevencimiento$", r"^vencimiento$"
    ],
    "Cuota Numero": [
        r"^cuotanumero$", r"^nrocuota$", r"^numerocuota$", r"^numerodecuota$"
    ]
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


class SelectModuleRequest(BaseModel):
    module_id: str = "general"
