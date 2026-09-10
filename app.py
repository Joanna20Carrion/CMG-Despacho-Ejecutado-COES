from __future__ import annotations

from datetime import datetime, date, timedelta
from io import BytesIO
import zipfile
import requests
import pandas as pd
import streamlit as st
from zoneinfo import ZoneInfo

# ------------------ Configuración ------------------
MES_TXT_TITLE = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"]
MES_TXT       = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"]

base_ieod = ("https://www.coes.org.pe/portal/browser/download?"
             "url=Post%20Operaci%C3%B3n%2FReportes%2FIEOD%2F"
             "{y}%2F{m}_{M}%2F{d}%2FAnexoA_{ddmm}.xlsx")
base_cmg  = ("https://www.coes.org.pe/portal/browser/download?url=Post%20Operaci%C3%B3n%2FReportes%2FIEOD%2F"
             "{y}%2F{m}_{M}%2F{d}%2FCMg{y}{m}{d}.zip")

# ------------------ Orden deseado de columnas (Despacho Ejecutado) ------------------
ORDEN_CENTRALES = [
    "MANTARO","RESTITUCION","TUMBES MAK 1","TUMBES MAK 2","HUINCO","MATUCANA",
    "CALLAHUANCA","MOYOPAMPA","HUAMPANI","CH_HER1","STA ROSA WEST TG7","STAROSA TG8",
    "UTI_5","UTI_6","VENTANILLA 3","VENTANILLA 4","VENTANILLA TV","RUBI","CLEMESI",
    "WAYRA-I","CHIMAY","YANANGO","CHEVES","YAUPI","MALPASO","OROYA","PACHACHACA",
    "CAHUA","ARCATA","PARIAC","GALLITO CIEGO","SHOUGESA TV1","SHOUGESA TV2",
    "SHOUGESA TV3","SHOUGESA CUMMINS","MALACAS1 TG6","MALACAS2 TG4","MALACAS3 TG 5",
    "CAÑON DEL PATO","CARHUAQUERO","CARHUAQUERO4","CAÑA BRAVA","AGUAYTIA TG1",
    "AGUAYTIA TG2","MACHUPICCHU","ARICOTA 1","ARICOTA 2","INDEPENDENCIA","YUNCAN",
    "ILO2 TV CARB1","RF ILO2 TG1","RF ILO2 TG2","RF ILO2 TG3","CTNEPI_TG41",
    "CTNEPI_TG42","CTNEPI_TG43","CHILCA1 TG1","CHILCA1 TG2","CHILCA1 TG3","CHILCA-TV",
    "CHILCA2 TG41","CHILCA2-TV","INTIPAMPA","PUNTA LOMITAS-BL1","PUNTA LOMITAS-BL2",
    "PUN LOMITAS_EXP-BL1","PUN LOMITAS_EXP-BL2","CHARCANI I,II,III","CHARCANI IV",
    "CHARCANI V","CHARCANI VI","CHILINA TG","MOLLENDO 1, 2, 3","CHILINA (SULZ 1,2)",
    "SAN GABAN II","MCH TUPURI","KALLPA TG1","KALLPA TG2","KALLPA TG3","KALLPA TV",
    "LFLORES TG1","LFLORES TV","CERRO DEL AGUILA","CERRODELAGUILAG4",
    "FENIX GT11","FENIX GT12","FENIX TV10","OLLEROS TG1","OLLEROS TV",
    "RF ETEN TG1","RF ETEN TG2","RECKA","PTO_BRVO TG1","PTO_BRVO TG2","PTO_BRVO TG3",
    "PTO_BRVO TG4","PLATANAL","HUANZA","HUANCHOR","STA TERESA","CHAGLLA GP1","CHAGLLA",
    "RF PUCALLPA","RF PTO MALDONADO","OQUENDO TG1","OQUENDO TV1","PARAMONGA","MAPLE",
    "HUAYCOLORO","LAGRINGAV","DONA_CATALINA","CALLAO","STACRUZ_1_2","HUASAHUASI",
    "RUNATULLO_II","RUNATULLO_III","POECHOS2","CHANCAY","YANAPAMPA","PIZARRAS",
    "RONCADOR","PURMACANA","LA JOYA","ANGEL_I","ANGEL_II","ANGEL_III","NIMPERIAL",
    "CANCHAYLLO","RUCUY","POTRERO","MARAÑON","YARUCAYA","RENOVANDESH1","LA VIRGEN",
    "CARHUAC","SAN_JACINTO","CANA BRAVA","8AGOSTO","EL CARMEN","MANTA I",
    "SANTA ROSA 1","SANTA ROSA 2","MAJES","REPARTICION","TACNA-SOLAR",
    "PANAMERICANA-SOLAR","MOQUEGUA-SOLAR","CS YARUCAYA","PQE-EOLICO-MARCONA",
    "PQE-EOLICO-3-HERMANAS","HUAMBOS","DUNA","PQE-EOLICO-TALARA",
    "PQE-EOLICO-CUPISNIQUE","FLUJO AL ECUADOR","SAN JUAN","WAYRA EXTENSION","QUITARACSA",
    "CARHUAQUERO (CS)","REFINERIA TALARA TV1","REFINERÍA TALARA TV2","HUALLIN",
    "MATARANI","SAN GABAN III","SAN MARTIN SOLAR","ANASHIRONI","AGROLMOS","SUNNY",
    "CASA GRANDE","EXPANSIÓN INTIPAMPA","COENERGY","SUNNY EXPANSION","WAYRA SOLAR","SAN JOSE",
]

# Alias: nombre en la lista → nombre que puede aparecer en el export (distinto)
ALIAS_CENTRALES = {
    "MCH TUPURI": "TUPURI",
    "8AGOSTO":    "8AGOSTO",   # el export ya lo tiene así
}


def _norm(s: str) -> str:
    """Normaliza para comparar: mayúsculas, sin tabs, sin espacios dobles."""
    return " ".join(str(s).replace("\t", " ").split()).upper()


def reordenar_despacho(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe el df con columnas 'EMPRESA | CENTRAL' y FECHA,
    devuelve un df con FECHA + columnas en ORDEN_CENTRALES.
    Las columnas del orden que no existan quedan como NaN.
    Las columnas del export que no estén en el orden van al final.
    """
    # Mapeo norm(central) → nombre_columna_real_en_df
    central_map: dict[str, str] = {}
    for col in df.columns:
        central = col.split("|", 1)[1].strip() if "|" in col else col
        # normalizar quitando tab
        central_norm = _norm(central)
        central_map[central_norm] = col

    # FENIX: columnas rotas por tab → mapear por nombre limpio
    # "FENIX POWER PERÚ | FENIX \tGT11" se lee como "FENIX POWER PERÚ | FENIX " + col suelta "GT11"
    # Reasignamos manualmente buscando la col que termina en "| FENIX " (sin GT)
    for col in df.columns:
        if col.endswith("| FENIX "):
            central_map[_norm("FENIX GT11")] = col
        if col.endswith("| FENIX .1"):
            central_map[_norm("FENIX GT12")] = col

    # Columnas ya usadas (para no repetirlas en el "resto")
    usadas: set[str] = set()

    ordered_cols: list[str] = ["FECHA"]
    usadas.add("FECHA")

    for desired in ORDEN_CENTRALES:
        key = _norm(desired)
        # probar alias
        alias_key = _norm(ALIAS_CENTRALES.get(desired, desired))

        col_found = central_map.get(key) or central_map.get(alias_key)

        if col_found and col_found not in usadas:
            ordered_cols.append(col_found)
            usadas.add(col_found)
        else:
            # columna faltante → insertar columna vacía con el nombre deseado
            ordered_cols.append(f"__MISSING__{desired}")

    # Columnas del export que no estaban en el orden → al final
    for col in df.columns:
        if col not in usadas and col != "FECHA":
            ordered_cols.append(col)

    # Construir df final
    result_frames = []
    for col in ordered_cols:
        if col.startswith("__MISSING__"):
            label = col[len("__MISSING__"):]
            result_frames.append(pd.Series([None] * len(df), name=label))
        else:
            s = df[col].copy()
            # Renombrar: solo la parte de la central (sin empresa)
            if "|" in col:
                label = col.split("|", 1)[1].strip()
                # Limpiar tab del nombre FENIX
                label = " ".join(label.replace("\t", " ").split())
            else:
                label = col
            s.name = label
            result_frames.append(s)

    df_out = pd.concat(result_frames, axis=1)
    return df_out


def df_a_xlsx(df: pd.DataFrame) -> bytes:
    """Convierte un DataFrame a bytes xlsx (sin fórmulas, listo para download)."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Despacho Ejecutado")
        # Ajustar ancho de columnas automáticamente
        ws = writer.sheets["Despacho Ejecutado"]
        for col_cells in ws.columns:
            max_len = max(
                len(str(c.value)) if c.value is not None else 0
                for c in col_cells
            )
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 30)
    return buf.getvalue()


# -----------------------------------------------------------------------------
# ------------------------------- PANTALLA ------------------------------------
# -----------------------------------------------------------------------------
def render_graficos_en_pantalla(ini: date, fin: date):
    tab1, tab2 = st.tabs(["CMG", "Despacho Ejecutado"])

    # =========================================================
    # ======================== CMG ============================
    # =========================================================
    with tab1:
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
                url = base_cmg.format(y=y, m=m, d=d, M=M)
                try:
                    r = requests.get(url, verify=False, timeout=30)
                    r.raise_for_status()
                except Exception:
                    return None
                try:
                    z = zipfile.ZipFile(BytesIO(r.content))
                    nombre_excel = next(
                        (f for f in z.namelist() if f.endswith(".xlsx") and "CMgCP" in f),
                        None,
                    )
                    if nombre_excel is None:
                        return None
                    with z.open(nombre_excel) as f:
                        df_raw = pd.read_excel(f, sheet_name="Cmg_Barra", header=None)
                    header_row = df_raw.iloc[2]
                    col_validas = []
                    for i in range(1, len(header_row)):
                        if pd.isna(header_row[i]):
                            break
                        col_validas.append(i)
                    headers = hacer_columnas_unicas(
                        [str(header_row[i]) for i in col_validas]
                    )
                    df = df_raw.iloc[3:51, col_validas].copy()
                    df.columns = headers
                    return df.dropna(how="all").reset_index(drop=True)
                except Exception:
                    return None

            resultados = []
            fecha_actual = ini
            with st.spinner("Procesando CMG histórico…"):
                while fecha_actual <= fin:
                    y = fecha_actual.year
                    m = f"{fecha_actual.month:02d}"
                    d = f"{fecha_actual.day:02d}"
                    M = MES_TXT_TITLE[int(m) - 1]
                    df_dia = leer_cmg_dia(y, m, d, M)
                    if df_dia is not None and not df_dia.empty:
                        df_dia["FECHA"] = fecha_actual
                        resultados.append(df_dia)
                    fecha_actual += timedelta(days=1)

            if not resultados:
                st.info("No hay datos CMG para el rango seleccionado.")
            else:
                df_final = pd.concat(resultados, ignore_index=True)
                cols = ["FECHA"] + [c for c in df_final.columns if c != "FECHA"]
                df_final = df_final[cols]

                st.dataframe(df_final, use_container_width=True)

                # ── Descarga CMG ──
                c1, c2 = st.columns([1, 1])
                with c1:
                    st.download_button(
                        "⬇️ Descargar CSV",
                        data=df_final.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"CMG_{ini}_{fin}.csv",
                        mime="text/csv",
                    )
                with c2:
                    st.download_button(
                        "⬇️ Descargar Excel",
                        data=df_a_xlsx(df_final),
                        file_name=f"CMG_{ini}_{fin}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        except Exception as e:
            st.error(f"Error en CMG Histórico: {e}")

    # =========================================================
    # ================== DESPACHO EJECUTADO ===================
    # =========================================================
    with tab2:
        st.markdown("### DESPACHO EJECUTADO")

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

        def leer_ieod_dia(y, m, d, M):
            ddmm = f"{d}{m}"
            url = base_ieod.format(y=y, m=m, d=d, M=M, ddmm=ddmm)
            try:
                r = requests.get(url, verify=False, timeout=30)
                r.raise_for_status()
            except Exception:
                return None
            try:
                xls = pd.ExcelFile(BytesIO(r.content))
                if "DESPACHO_EJECUTADO" not in xls.sheet_names:
                    return None
                df_raw = pd.read_excel(xls, sheet_name="DESPACHO_EJECUTADO", header=None)
                todas_cols = list(range(df_raw.shape[1]))
                headers = []
                for col in todas_cols:
                    h1 = str(df_raw.iloc[4, col]) if not pd.isna(df_raw.iloc[4, col]) else ""
                    h2 = str(df_raw.iloc[8, col]) if not pd.isna(df_raw.iloc[8, col]) else ""
                    h3 = str(df_raw.iloc[9, col]) if not pd.isna(df_raw.iloc[9, col]) else ""
                    header = " | ".join([h for h in [h1, h2, h3] if h.strip()])
                    headers.append(header if header else f"COL_{col}")
                col_validas = []
                for i, h in enumerate(headers):
                    col_validas.append(i)
                    if h.strip() == "MW":
                        break
                headers = hacer_columnas_unicas([headers[i] for i in col_validas])
                df = df_raw.iloc[10:, col_validas].copy()
                df.columns = headers
                return df.dropna(how="all").reset_index(drop=True)
            except Exception:
                return None

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

        if not resultados:
            st.info("No hay datos disponibles para el rango seleccionado.")
        else:
            df_raw_final = pd.concat(resultados, ignore_index=True)

            # ── Aplicar orden deseado ──
            df_final = reordenar_despacho(df_raw_final)

            st.dataframe(df_final, use_container_width=True)

            # ── Descarga Despacho ──
            c1, c2 = st.columns([1, 1])
            with c1:
                st.download_button(
                    "⬇️ Descargar CSV",
                    data=df_final.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"Despacho_{ini}_{fin}.csv",
                    mime="text/csv",
                )
            with c2:
                st.download_button(
                    "⬇️ Descargar Excel",
                    data=df_a_xlsx(df_final),
                    file_name=f"Despacho_{ini}_{fin}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# -----------------------------------------------------------------------------
# --------------------------------- CONFIG ------------------------------------
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Reporte CMG y Despacho Ejecutado",
    layout="wide",
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
ini       = st.sidebar.date_input("Inicio del rango",    value=date.today(), format="DD/MM/YYYY")
fecha_sel = st.sidebar.date_input("Fecha del reporte",   value=ini,          format="DD/MM/YYYY")
fin       = fecha_sel
gen_generar = st.sidebar.button("Generar", type="primary")

st.title("Reporte CMG y Despacho Ejecutado")

if gen_generar:
    fecha_hum = fecha_sel.strftime("%d/%m/%Y")
    ahora_pe  = datetime.now(ZoneInfo("America/Lima"))
    now_str   = ahora_pe.strftime("%H:%M")

    st.subheader(f"Reporte del {fecha_hum}")
    st.caption(f"Actualizado a las {now_str} horas")

    with st.spinner("Renderizando en pantalla…"):
        render_graficos_en_pantalla(ini=ini, fin=fin)

st.caption("© Reporte CMG y Despacho Ejecutado - USGE")