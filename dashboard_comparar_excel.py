import streamlit as st
import pandas as pd
import io
import os
import requests
import time
from datetime import datetime
import plotly.express as px

# ======================================================
# CONFIGURACIÓN DE LA PÁGINA
# ======================================================
st.set_page_config(page_title="Comparador de Excels", layout="wide")
st.title("Dashboard Comparador de Excels")
st.markdown("""
Esta herramienta permite comparar varios archivos Excel (.xlsx),
detectar coincidencias, encontrar registros exclusivos
y consultar información de revistas en OpenAlex.
""")
st.divider()

# ======================================================
# PANEL LATERAL
# ======================================================
st.sidebar.header("Configuración")

modo = st.sidebar.radio("Selecciona el modo de ejecución:", ["Rápido", "Avanzado"])
usar_openalex = st.sidebar.checkbox("Consultar información en OpenAlex (batch)", value=False)
consultar_solo_uno = st.sidebar.checkbox("Consultar OpenAlex para un solo archivo", value=False)

correo_openalex = st.sidebar.text_input(
    "Correo para identificarte ante OpenAlex (recomendado)",
    placeholder="tucorreo@institucion.cl"
)

archivos = st.sidebar.file_uploader(
    "Sube uno o más archivos Excel (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

# ======================================================
# FUNCIONES AUXILIARES
# ======================================================
@st.cache_data
def leer_excel(archivo):
    try:
        return pd.read_excel(archivo)
    except Exception as e:
        st.error(f"Error al leer {archivo.name}: {e}")
        return pd.DataFrame()

def normalizar_valor(valor):
    """Normaliza ISSN, ISBN, EISSN, etc."""
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = valor.replace("-", "").replace(" ", "").replace(".", "")
    if valor.isdigit() and len(valor) == 8:
        return valor
    return valor

def generar_clave_prioritaria(row, columnas, normalizar=False):
    """Devuelve la primera columna con valor válido, con o sin normalización."""
    for col in columnas:
        valor = row[col]
        if normalizar:
            valor = normalizar_valor(valor)
        if valor and str(valor).lower() != "nan":
            return valor
    return None

@st.cache_data
def consultar_openalex_batch(lista_issn, correo_openalex=None):
    """Consulta OpenAlex en lotes de 50 ISSN."""
    resultados = []
    base_url = "https://api.openalex.org/sources?filter=issn:"
    batch_size = 50

    for i in range(0, len(lista_issn), batch_size):
        lote = lista_issn[i:i + batch_size]
        url = base_url + "|".join(lote)
        if correo_openalex:
            url += f"&mailto={correo_openalex}"

        try:
            r = requests.get(url)
            if r.status_code == 200:
                data = r.json()
                resultados_lote = data.get("results", [])
                for item in resultados_lote:
                    resultados.append({
                        "ISSN": item.get("issn_l"),
                        "Nombre revista": item.get("display_name", ""),
                        "País": item.get("country_code", ""),
                        "Tipo": item.get("type", ""),
                        "Acceso abierto": "Sí" if item.get("is_oa") else "No",
                        "Publisher": item.get("host_organization_name", ""),
                        "Total artículos": item.get("works_count", ""),
                        "Citado por": item.get("cited_by_count", ""),
                        "Última actualización": item.get("updated_date", "")
                    })
            else:
                st.warning(f"Error {r.status_code} al consultar OpenAlex.")
            time.sleep(0.3)
        except Exception as e:
            st.error(f"Error consultando OpenAlex: {e}")

    return pd.DataFrame(resultados)

# ======================================================
# PROCESO PRINCIPAL
# ======================================================
if archivos:
    dfs = [leer_excel(a) for a in archivos]
    nombres = [a.name for a in archivos]

    # CASO 1: CONSULTAR UN SOLO ARCHIVO
    if len(archivos) == 1 and consultar_solo_uno:
        st.subheader("Vista previa del archivo cargado")
        df = dfs[0]
        filas, columnas = df.shape
        st.markdown(f"{nombres[0]} — {filas} filas × {columnas} columnas")
        st.dataframe(df.head(10))

        columna_issn = st.selectbox("Selecciona la columna que contiene ISSN o E-ISSN", df.columns)
        if st.button("Consultar OpenAlex"):
            issn_unicos = df[columna_issn].dropna().astype(str).unique().tolist()
            st.info("Consultando OpenAlex, por favor espera unos segundos...")
            df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)
            if len(df_openalex) > 0:
                st.success(f"Se obtuvieron {len(df_openalex)} resultados desde OpenAlex.")
                st.dataframe(df_openalex)
                output = io.BytesIO()
                df_openalex.to_excel(output, index=False, sheet_name="OpenAlex_Resultados")
                output.seek(0)
                st.download_button(
                    "Descargar resultados de OpenAlex",
                    data=output,
                    file_name=f"OpenAlex_{nombres[0]}_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No se obtuvieron resultados desde OpenAlex. "
                           "Verifica el formato de los ISSN o el correo institucional.")
        st.stop()

    # CASO 2: COMPARACIÓN ENTRE VARIOS ARCHIVOS
    if len(archivos) > 1:
        st.subheader("Vista previa de los archivos cargados")
        for nombre, df in zip(nombres, dfs):
            filas, columnas = df.shape
            st.markdown(f"{nombre} — {filas} filas × {columnas} columnas")
            st.dataframe(df.head(10))
            st.markdown("---")

        columnas_comunes = set(dfs[0].columns)
        for df in dfs[1:]:
            columnas_comunes &= set(df.columns)
        columnas_comunes = list(columnas_comunes)

        if columnas_comunes:
            columnas_clave = st.multiselect(
                "Selecciona las columnas clave para comparar (se usará la primera con datos válidos por fila)",
                columnas_comunes
            )

            if columnas_clave:
                for df in dfs:
                    df[columnas_clave] = df[columnas_clave].fillna("")
                    df["__clave__"] = df.apply(
                        lambda r: generar_clave_prioritaria(
                            r, columnas_clave, normalizar=(modo == "Avanzado")
                        ),
                        axis=1,
                    )
                    df.dropna(subset=["__clave__"], inplace=True)

                claves = pd.concat([df[["__clave__"]] for df in dfs], keys=range(len(dfs)))
                claves["Archivo"] = claves.index.get_level_values(0)
                conteo = claves.groupby("__clave__")["Archivo"].nunique()

                claves_comunes = conteo[conteo > 1].index
                coincidencias_total = pd.concat([
                    df[df["__clave__"].isin(claves_comunes)] for df in dfs
                ]).drop(columns=["__clave__"])

                exclusivos_por_archivo = [
                    df[df["__clave__"].isin(conteo[conteo == 1].index)].drop(columns=["__clave__"])
                    for df in dfs
                ]

                total_exclusivos = sum(len(df) for df in exclusivos_por_archivo)
                st.divider()
                st.subheader("Resumen general")
                c1, c2, c3 = st.columns(3)
                c1.metric("Archivos cargados", len(archivos))
                c2.metric("Coincidencias encontradas", len(coincidencias_total))
                c3.metric("Registros exclusivos", total_exclusivos)

                st.markdown("Visualización de resultados")
                fig1 = px.pie(
                    pd.DataFrame({
                        "Tipo": ["Coincidencias", "Exclusivos"],
                        "Cantidad": [len(coincidencias_total), total_exclusivos]
                    }),
                    names="Tipo", values="Cantidad",
                    title="Distribución general de registros",
                    color="Tipo",
                    color_discrete_map={"Coincidencias": "#2ECC71", "Exclusivos": "#3498DB"}
                )
                fig1.update_traces(textinfo="percent+value")
                st.plotly_chart(fig1, use_container_width=True)

                fig2 = px.bar(
                    pd.DataFrame({
                        "Archivo": nombres,
                        "Exclusivos": [len(df) for df in exclusivos_por_archivo],
                    }),
                    x="Archivo", y="Exclusivos",
                    title="Registros exclusivos por archivo",
                    text_auto=True,
                    color="Archivo"
                )
                st.plotly_chart(fig2, use_container_width=True)

                df_openalex = pd.DataFrame()
                if usar_openalex:
                    st.info("Consultando OpenAlex para coincidencias...")
                    if "ISSN" in coincidencias_total.columns:
                        issn_unicos = coincidencias_total["ISSN"].dropna().astype(str).unique().tolist()
                        df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)
                        if len(df_openalex) > 0:
                            st.success(f"Se obtuvieron {len(df_openalex)} registros desde OpenAlex.")
                            st.dataframe(df_openalex.head(20))
                        else:
                            st.warning("No se obtuvieron resultados desde OpenAlex. "
                                       "Verifica el formato de los ISSN o el correo ingresado.")
                    else:
                        st.warning("No se encontró columna 'ISSN' en los archivos para consultar OpenAlex.")

                # === GENERAR ARCHIVO SOLO CUANDO SE PRESIONE DESCARGAR ===
                st.divider()
                st.markdown("Generar archivo Excel con los resultados")

                if st.button("Generar y preparar archivo para descarga"):
                    with st.spinner("Generando archivo Excel... por favor espera unos segundos."):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            workbook = writer.book

                            resumen = pd.DataFrame({
                                "Parámetro": [
                                    "Fecha de generación",
                                    "Modo de ejecución",
                                    "Archivos comparados",
                                    "Columnas clave",
                                    "Coincidencias encontradas",
                                    "Registros exclusivos totales"
                                ],
                                "Valor": [
                                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    modo,
                                    ", ".join(nombres),
                                    ", ".join(columnas_clave),
                                    len(coincidencias_total),
                                    total_exclusivos
                                ]
                            })
                            resumen.to_excel(writer, sheet_name="Resumen", index=False)
                            coincidencias_total.to_excel(writer, sheet_name="Coincidencias", index=False)

                            for i, exclusivos in enumerate(exclusivos_por_archivo):
                                nombre_limpio = os.path.splitext(nombres[i])[0]
                                nombre_limpio = "".join(c for c in nombre_limpio if c.isalnum() or c in (" ", "_", "-"))
                                nombre_hoja = f"Exclusivos_{nombre_limpio}"[:31]
                                exclusivos.to_excel(writer, sheet_name=nombre_hoja, index=False)

                            if 'df_openalex' in locals() and not df_openalex.empty:
                                df_openalex.to_excel(writer, sheet_name="OpenAlex_Resultados", index=False)

                        output.seek(0)
                        st.session_state["excel_resultado"] = output.getvalue()

                    st.success("Archivo Excel generado correctamente. Ahora puedes descargarlo.")

                if "excel_resultado" in st.session_state:
                    st.download_button(
                        label="Descargar archivo Excel con resultados",
                        data=st.session_state["excel_resultado"],
                        file_name=f"resultado_comparacion_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.warning("Selecciona al menos una columna clave para realizar la comparación.")
else:
    st.info("Sube al menos un archivo Excel para comenzar la comparación.")
