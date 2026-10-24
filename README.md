# 📊 CobranzApp - Motor de Reestructuración y Segmentación de Cartera

Sistema integral para el análisis, segmentación inteligente y reestructuración de carteras de crédito y cobranza vencida en USD. Permite procesar de forma instantánea archivos masivos de mora (`.csv`, `.xlsx`, `.xls`), aplicar tramos de gestión estratégica y exportar reportes ejecutivos en PDF y Excel para equipos operativos.

---

## 🚀 Características Principales

- **Ingesta Inteligente de Datos:**
  - Carga y procesamiento tolerante a fallos para archivos CSV y Excel de gran tamaño (más de 47.000 cuotas por archivo).
  - Normalización automática de nombres de columnas y detección de formatos numéricos (soporta `$`, `Bs.`, comas y puntos decimales latinos/anglosajones).
  - Selector inteligente con puntuación de pestañas para libros Excel con múltiples hojas (`mora general`, `consolidado`, `call center`, etc.).
- **Segmentación por Tramos Estratégicos:**
  - **1 a 30 días:** Cobranza preventiva / temprana (29.8% de la cartera).
  - **31 a 60 días:** Cobranza administrativa (13.5% de la cartera).
  - **61 a 90 días:** Reestructuración y negociación diferida (10.6% de la cartera).
  - **91 a 180 días:** Pre-jurídica y condonación de mora (35.0% de la cartera - tramo crítico).
  - **181+ días:** Castigo y cartera pesada (11.2% de la cartera).
- **Filtros Multidimensionales:**
  - Filtrado cruzado por Campaña / Programa (`CREDIX`, `FIBEX`, `WEB`, etc.).
  - Filtrado por Tipo de Deuda (`cuota`, `cuota/mora`, `cuota/tarifa/mora`, etc.).
- **Reportes y Exportaciones:**
  - **Excel Operativo:** Archivo enriquecido con todas las cuotas filtradas listo para dialers o gestores de cobranza.
  - **PDF Ejecutivo Formal:** Informe gerencial con indicadores clave (KPIs), promedios por cliente y tablas de distribución porcentual.
- **Guía Visual Integrada:**
  - Manual de uso paso a paso accesible en la interfaz web con recomendaciones tácticas por cada tramo de atraso.

---

## 🛠️ Tecnologías Utilizadas

- **Backend:** Python 3.12+, FastAPI, Uvicorn, Pandas, OpenPyXL, Xlrd, ReportLab.
- **Frontend:** HTML5, Tailwind CSS, FontAwesome 6, Chart.js.

---

## 📦 Instalación y Puesta en Marcha

### 1. Clonar el repositorio
```bash
git clone https://github.com/francojbe/cobranzapp-.git
cd cobranzapp-
```

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 3. Iniciar el servidor
Puedes iniciar el servidor mediante el archivo por lotes en Windows:
```bash
iniciar_servidor.bat
```
O directamente con Python:
```bash
python app.py
```

El servidor estará disponible en [http://127.0.0.1:8000](http://127.0.0.1:8000).

---

## 📂 Estructura del Proyecto

```text
cobranzapp-/
│
├── app.py                  # API principal (FastAPI), lógica de cálculo y reportes
├── iniciar_servidor.bat    # Script de arranque rápido para Windows
├── requirements.txt        # Dependencias de Python
├── mora.csv               # Archivo base de prueba de cartera (47.083 filas)
├── test_pipeline.py        # Suite de pruebas automatizadas del flujo de datos
├── templates/
│   └── index.html          # Interfaz web interactiva con Tailwind y Chart.js
└── README.md               # Documentación del proyecto
```

---

## 📄 Licencia

Uso interno y desarrollo empresarial por **ProFrank**.
