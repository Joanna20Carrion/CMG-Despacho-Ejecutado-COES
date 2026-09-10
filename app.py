from __future__ import annotations

from datetime import datetime, date, timedelta
import zipfile
import requests
import pandas as pd
import streamlit as st
from zoneinfo import ZoneInfo

# ------------------ Configuración ------------------
MES_TXT_TITLE = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"]
MES_TXT = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"]

base_ieod = ("https://www.coes.org.pe/portal/browser/download?"
             "url=Post%20Operaci%C3%B3n%2FReportes%2FIEOD%2F"
             "{y}%2F{m}_{M}%2F{d}%2FAnexoA_{ddmm}.xlsx")
base_cmg = ("https://www.coes.org.pe/portal/browser/download?url=Post%20Operaci%C3%B3n%2FReportes%2FIEOD%2F"
            "{y}%2F{m}_{M}%2F{d}%2FCMg{y}{m}{d}.zip")

# -----------------------------------------------------------------------------
# ------------------------------- PANTALLA ------------------------------------
# -----------------------------------------------------------------------------
def render_graficos_en_pantalla(ini: date, fin: date):    
    # ==== Pestañas ====
    tab1, tab2  = st.tabs(["CMG","Despacho Ejecutado"])      
        
    with tab1:
    # =========================================================
    # ======================== CMG ============================
    # =========================================================
        try:
            st.markdown("### COSTOS MARGINALES")

            def hacer_columnas_unicas(cols):
                conteo = {}
                nuevas = []

                for col in cols:
                    if col not in conteo:
                        conteo[col] = 0
                        nuevas.append(col)
                    else:
                        conteo[col] += 1
                        nuevas.append(f"{col}_{conteo[col]}")

                return nuevas

            def leer_cmg_dia(y, m, d, M):
                from io import BytesIO

                url = base_cmg.format(
                    y=y,
                    m=m,
                    d=d,
                    M=M
                )

                try:
                    r = requests.get(
                        url,
                        verify=False,
                        timeout=30
                    )
                    r.raise_for_status()
                except Exception:
                    return None

                try:
                    z = zipfile.ZipFile(BytesIO(r.content))

                    # Buscar el Excel dentro del ZIP
                    nombre_excel = None

                    for f in z.namelist():
                        if f.endswith(".xlsx") and "CMgCP" in f:
                            nombre_excel = f
                            break

                    if nombre_excel is None:
                        return None

                    with z.open(nombre_excel) as f:
                        df_raw = pd.read_excel(
                            f,
                            sheet_name="Cmg_Barra",
                            header=None
                        )

                    # Encabezado: fila 3 → índice 2
                    header_row = df_raw.iloc[2]

                    # Detectar columnas desde B hasta el fin real
                    col_validas = []

                    for i in range(1, len(header_row)):
                        if pd.isna(header_row[i]):
                            break

                        col_validas.append(i)

                    headers = [
                        str(header_row[i])
                        for i in col_validas
                    ]

                    # Evitar columnas duplicadas
                    headers = hacer_columnas_unicas(headers)

                    # Datos: filas 4 a 51
                    df = df_raw.iloc[
                        3:51,
                        col_validas
                    ].copy()

                    df.columns = headers

                    df = (
                        df
                        .dropna(how="all")
                        .reset_index(drop=True)
                    )

                    return df

                except Exception:
                    return None

            # ========== RECORRER RANGO DE FECHAS ==========
            resultados = []
            fecha_actual = ini

            with st.spinner("Procesando CMG histórico…"):

                while fecha_actual <= fin:

                    y = fecha_actual.year
                    m = f"{fecha_actual.month:02d}"
                    d = f"{fecha_actual.day:02d}"
                    M = MES_TXT_TITLE[
                        int(m) - 1
                    ]

                    df_dia = leer_cmg_dia(
                        y,
                        m,
                        d,
                        M
                    )

                    if (
                        df_dia is not None
                        and not df_dia.empty
                    ):
                        df_dia["FECHA"] = fecha_actual
                        resultados.append(df_dia)

                    fecha_actual += timedelta(days=1)

            # ============ MOSTRAR TABLA ============
            if not resultados:

                st.info(
                    "No hay datos CMG para el rango seleccionado."
                )

            else:

                df_final = pd.concat(
                    resultados,
                    ignore_index=True
                )

                # FECHA primero
                cols = (
                    ["FECHA"]
                    + [
                        c
                        for c in df_final.columns
                        if c != "FECHA"
                    ]
                )

                df_final = df_final[cols]

                # Tabla interactiva
                st.dataframe(
                    df_final
                )

        except Exception as e:
            st.error(
                f"Error en CMG Histórico: {e}"
            )
            
    with tab2:
        # =========================================================
        # ================== DESPACHO EJECUTADO ===================
        # =========================================================
        st.markdown("### DESPACHO EJECUTADO")
        
        # ===== función para evitar columnas duplicadas =====
        def hacer_columnas_unicas(cols):
            conteo = {}
            nuevas = []
    
            for col in cols:
                if col not in conteo:
                    conteo[col] = 0
                    nuevas.append(col)
                else:
                    conteo[col] += 1
                    nuevas.append(f"{col}_{conteo[col]}")
    
            return nuevas
    
        # ===== función para leer IEOD por día =====
        def leer_ieod_dia(y, m, d, M):
            from io import BytesIO
    
            ddmm = f"{d}{m}"
            url = base_ieod.format(y=y, m=m, d=d, M=M, ddmm=ddmm)
    
            try:
                r = requests.get(url, verify=False, timeout=30)
                r.raise_for_status()
            except:
                return None
    
            try:
                xls = pd.ExcelFile(BytesIO(r.content))
    
                if "DESPACHO_EJECUTADO" not in xls.sheet_names:
                    return None
    
                df_raw = pd.read_excel(
                    xls,
                    sheet_name="DESPACHO_EJECUTADO",
                    header=None
                )
    
                # ===== TODAS las columnas =====
                todas_cols = list(range(df_raw.shape[1]))
    
                # ===== encabezados (filas 5, 9, 10) =====
                headers = []
                for col in todas_cols:
                    h1 = str(df_raw.iloc[4, col]) if not pd.isna(df_raw.iloc[4, col]) else ""
                    h2 = str(df_raw.iloc[8, col]) if not pd.isna(df_raw.iloc[8, col]) else ""
                    h3 = str(df_raw.iloc[9, col]) if not pd.isna(df_raw.iloc[9, col]) else ""
    
                    header = " | ".join([h for h in [h1, h2, h3] if h.strip() != ""])
                    headers.append(header if header else f"COL_{col}")
    
                # ===== RECORTE por columna "MW" =====
                col_validas = []
                for i, h in enumerate(headers):
                    col_validas.append(i)
                    if h.strip() == "MW":
                        break
    
                headers = [headers[i] for i in col_validas]
    
                # ===== hacer encabezados únicos =====
                headers = hacer_columnas_unicas(headers)
    
                # ===== data =====
                df = df_raw.iloc[10:, col_validas].copy()
                df.columns = headers
    
                # eliminar filas vacías
                df = df.dropna(how="all").reset_index(drop=True)
    
                return df
    
            except:
                return None
    
        # ===== recorrer rango =====
        resultados = []
        fecha_actual = ini
    
        with st.spinner("Procesando IEOD (Despacho Ejecutado)…"):
            while fecha_actual <= fin:
                y = fecha_actual.year
                m = f"{fecha_actual.month:02d}"
                d = f"{fecha_actual.day:02d}"
                M = MES_TXT[int(m) - 1]
    
                df_dia = leer_ieod_dia(y, m, d, M)
    
                if df_dia is not None and not df_dia.empty:
                    df_dia["FECHA"] = fecha_actual
                    resultados.append(df_dia)
    
                fecha_actual += timedelta(days=1)
    
        # ===== mostrar =====
        if not resultados:
            st.info("No hay datos disponibles para el rango seleccionado.")
        else:
            df_final = pd.concat(resultados, ignore_index=True)
    
            cols = ["FECHA"] + [c for c in df_final.columns if c != "FECHA"]
            df_final = df_final[cols]
    
            st.dataframe(df_final)
            
# -----------------------------------------------------------------------------
# ------------------------------------ PDF ------------------------------------
# -----------------------------------------------------------------------------          
st.set_page_config(
    page_title="Reporte CMG y Despacho Ejecutado",
    layout="wide"
)

st.markdown("""
<style>
div[data-baseweb="tab-list"] {
    overflow-x: auto;
    white-space: nowrap;
    scrollbar-width: thin;
}

div[data-baseweb="tab-list"] button {
    flex: 0 0 auto;
}
</style>
""", unsafe_allow_html=True)

st.sidebar.header("Parámetros")
ini = st.sidebar.date_input("Inicio del rango", value=date.today(), format="DD/MM/YYYY")
fecha_sel = st.sidebar.date_input("Fecha del reporte", value=ini, format="DD/MM/YYYY")
fin = fecha_sel
gen_generar = st.sidebar.button("Generar", type="primary")

st.title("Reporte CMG y Despacho Ejecutado")
btn_cols = st.columns([1, 8])

if gen_generar:
    
    fecha_hum = fecha_sel.strftime("%d/%m/%Y")
    
    ahora_pe = datetime.now(ZoneInfo("America/Lima"))
    now_str   = ahora_pe.strftime("%H:%M")
    
    y, m, d = fecha_sel.year, f"{fecha_sel.month:02d}", f"{fecha_sel.day:02d}"
    M = MES_TXT[int(m) - 1]
    fecha_str = f"{y}{m}{d}"
    ddmm = f"{d}{m}"
    
    st.subheader(f"Reporte del {fecha_hum}")
    st.caption(f"Actualizado a las {now_str} horas")

    with st.spinner("Renderizando en pantalla…"):
        render_graficos_en_pantalla(ini=ini, fin=fin)

st.caption("© Reporte CMG y Despacho Ejecutado - USGE")