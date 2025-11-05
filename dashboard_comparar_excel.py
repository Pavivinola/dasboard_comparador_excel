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
st.set_page_config(page_title="Compareitor", layout="wide")
st.title("Compareitor")
st.markdown("""
Esta herramienta permite comparar varios archivos Excel (.xlsx),
detectar coincidencias, encontrar registros exclusivos
y consultar información de revistas en OpenAlex.
""")
st.divider()

# ======================================================
# PANEL LATERAL
# ======================================================
st.sidebar.header("Configuración") # Este st. es para el panel lateral

modo = st.sidebar.radio("Selecciona el modo de ejecución:", ["Rápido", "Avanzado"])
usar_openalex = st.sidebar.checkbox("Consultar información en OpenAlex (batch)", value=False)
consultar_solo_uno = st.sidebar.checkbox("Consultar OpenAlex para un solo archivo", value=False)

correo_openalex = st.sidebar.text_input(
    "Correo para identificarte ante OpenAlex (recomendado)",
    placeholder="tucorreo@institucion.cl"
)

archivos = st.sidebar.file_uploader(
    "Sube uno o más archivos Excel (.xlsx)",
    type="xlsx", # Con esto solo se aceptan archivos .xlsx, en caso de que se agreguen otros tipos de archivo, salta un error
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
    """Normaliza ISSN, ISBN, EISSN, etc. Conserva el guion en posición 4 si aplica."""
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper().replace(" ", "").replace(".", "")
    valor = valor.replace("-", "")
    # Si parece un ISSN válido (8 caracteres), volver a agregar guion estándar
    if len(valor) == 8 and valor.isalnum():
        return valor[:4] + "-" + valor[4:]
    return valor

def obtener_issn_validos(df, columna):
    """
    Limpia y valida los ISSN de una columna.
    Devuelve una lista de ISSN únicos válidos (8 caracteres alfanuméricos, con guion).
    """
    if columna not in df.columns:
        st.warning(f"La columna '{columna}' no existe en el archivo.")
        return []

    # Convertir a texto, limpiar espacios, aplicar normalización
    df[columna] = df[columna].astype(str).fillna("").str.strip()
    df[columna] = df[columna].apply(normalizar_valor)

    # Filtrar válidos
    issn_unicos = [i for i in df[columna].unique() if len(i) == 9 and "-" in i]

    # Mostrar diagnóstico
    st.write(f"ISSN detectados en '{columna}':", issn_unicos[:15])
    st.write(f"Total ISSN válidos: {len(issn_unicos)}")

    return issn_unicos

def generar_clave_prioritaria(row, columnas, normalizar=False):
    """Devuelve la primera columna con valor válido, con o sin normalización."""
    for col in columnas:
        valor = row[col]
        if normalizar:
            valor = normalizar_valor(valor)
        if valor and str(valor).lower() != "nan":
            return valor
    return None

# ======================================================
# CONSULTA A OPENALEX
# ======================================================
def consultar_openalex_batch(issn_list, correo):
    """
    Consulta la API de OpenAlex en lotes de 30 ISSN (formato XXXX-XXXX, separados por coma).
    """
    resultados = []
    lote_size = 30
    base_url = "https://api.openalex.org/sources"

    if not issn_list:
        st.warning("No se encontraron ISSN válidos para consultar en OpenAlex.")
        return pd.DataFrame()

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_lotes = (len(issn_list) + lote_size - 1) // lote_size

    headers = {"User-Agent": f"CompareitorDashboard/1.0 (mailto:{correo})"}

    for i in range(0, len(issn_list), lote_size):
        lote = issn_list[i:i + lote_size]
        filtro = ",".join(lote)
        url = f"{base_url}?filter=issn:{filtro}"
        if correo:
            url += f"&mailto={correo}"

        status_text.text(f"Consultando lote {i//lote_size + 1} de {total_lotes}...")
        progress_bar.progress(min((i + lote_size) / len(issn_list), 1.0))

        try:
            response = requests.get(url, headers=headers, timeout=30)
            print(f"[{i//lote_size + 1}] HTTP {response.status_code} → {url}")
            if response.status_code == 200:
                data = response.json()
                lote_resultados = data.get("results", [])
                resultados.extend(lote_resultados)
            else:
                st.warning(f"Error HTTP {response.status_code} en el lote {i//lote_size + 1}.")
        except Exception as e:
            st.error(f"Error en lote {i//lote_size + 1}: {e}")

        time.sleep(1.2)

    progress_bar.empty()
    status_text.text("Consulta finalizada.")

    if resultados:
        df_openalex = pd.json_normalize(resultados)
        st.success(f"Se obtuvieron {len(df_openalex)} resultados desde OpenAlex.")
        st.dataframe(df_openalex.head())
        return df_openalex
    else:
        st.warning("No se obtuvieron resultados desde OpenAlex. "
                   "Verifica el formato de los ISSN o el correo institucional.")
        return pd.DataFrame()

# ======================================================
# PROCESO PRINCIPAL
# ======================================================
if archivos:
    dfs = [leer_excel(a) for a in archivos]
    nombres = [a.name for a in archivos]

    # === MODO: UN SOLO ARCHIVO ===
    if len(archivos) == 1 and consultar_solo_uno:
        st.subheader("Vista previa del archivo cargado")
        df = dfs[0]
        filas, columnas = df.shape
        st.markdown(f"{nombres[0]} — {filas} filas × {columnas} columnas")
        st.dataframe(df.head(10))

        columna_issn = st.selectbox("Selecciona la columna que contiene ISSN o E-ISSN", df.columns)
        if st.button("Consultar OpenAlex"):
            issn_unicos = obtener_issn_validos(df, columna_issn)
            if len(issn_unicos) == 0:
                st.warning("No se encontraron ISSN válidos para consultar en OpenAlex.")
            else:
                st.info("Consultando OpenAlex, por favor espera unos segundos...")
                df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)

                if len(df_openalex) > 0:
                    output = io.BytesIO()
                    df_openalex.to_excel(output, index=False, sheet_name="OpenAlex_Resultados")
                    output.seek(0)
                    st.download_button(
                        "Descargar resultados de OpenAlex",
                        data=output,
                        file_name=f"OpenAlex_{nombres[0]}_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        st.stop()

    # === MODO: COMPARAR VARIOS ARCHIVOS ===
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
                        lambda r: generar_clave_prioritaria( #El lambda usa la función definida arriba
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

                fig1 = px.pie(
                    pd.DataFrame({
                        "Tipo": ["Coincidencias", "Exclusivos"],
                        "Cantidad": [len(coincidencias_total), total_exclusivos]
                    }),
                    names="Tipo", values="Cantidad",
                    title="Distribución general de registros"
                )
                st.plotly_chart(fig1, use_container_width=True)

                df_openalex = pd.DataFrame()
                if usar_openalex:
                    st.info("Consultando OpenAlex para coincidencias...")
                    if "ISSN" in coincidencias_total.columns:
                        issn_unicos = obtener_issn_validos(coincidencias_total, "ISSN")

                        # Si no hay válidos, intentar con todos los archivos
                        if len(issn_unicos) == 0:
                            st.warning("No se encontraron ISSN válidos en coincidencias. Se intentará con todos los archivos.")
                            df_combinado = pd.concat(dfs, ignore_index=True)
                            issn_unicos = obtener_issn_validos(df_combinado, "ISSN")

                        if len(issn_unicos) == 0:
                            st.warning("No se encontraron ISSN válidos para consultar en OpenAlex.")
                        else:
                            df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)
                    else:
                        st.warning("No se encontró columna 'ISSN' en los archivos para consultar OpenAlex.")

                st.divider()
                st.markdown("Generar archivo Excel con los resultados")

                if st.button("Generar y preparar archivo para descarga"):
                    with st.spinner("Generando archivo Excel..."):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
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

                            if not df_openalex.empty:
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
