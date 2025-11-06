import streamlit as st
import pandas as pd
import io
import os
import requests
import time
import socket
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
# FUNCIÓN PARA DETECTAR ENTORNO (LOCAL O CLOUD)
# ======================================================
def es_entorno_local():
    try:
        ip_local = socket.gethostbyname(socket.gethostname())
        return ip_local.startswith("127.") or ip_local.startswith("192.") or ip_local.startswith("10.")
    except:
        return True

ENTORNO_LOCAL = es_entorno_local()

# ======================================================
# FUNCIONES AUXILIARES
# ======================================================
@st.cache_data
def leer_excel(archivo):
    """Lee un archivo Excel y elimina filas vacías."""
    try:
        df = pd.read_excel(archivo)
        df = df.dropna(how="all")
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Error al leer {archivo.name}: {e}")
        return pd.DataFrame()

def normalizar_valor(valor):
    """Normaliza ISSN, ISBN, etc."""
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper().replace(" ", "").replace(".", "")
    valor = valor.replace("_", "").replace("–", "-").replace("—", "-").replace("−", "-")
    valor = valor.replace("-", "")
    if len(valor) == 8 and valor.isalnum():
        return valor[:4] + "-" + valor[4:]
    return valor

def generar_clave_combinada(row, columnas, normalizar=False):
    """Genera una clave combinada con las columnas seleccionadas (OR lógico)."""
    valores = []
    for col in columnas:
        val = str(row[col]).strip()
        if val and val.lower() != "nan":
            val = normalizar_valor(val) if normalizar else val
            valores.append(val)
    if valores:
        return "|".join(sorted(set(valores)))
    return None

def obtener_issn_validos(df, columna):
    """Obtiene ISSN válidos de una columna."""
    if columna not in df.columns:
        return []
    df[columna] = df[columna].astype(str).fillna("").str.strip().apply(normalizar_valor)
    return [i for i in df[columna].unique() if len(i) == 9 and "-" in i]

# ======================================================
# FUNCIÓN DE CONSULTA OPENALEX (COMPATIBLE 2025)
# ======================================================
def consultar_openalex_batch(issn_list, correo_openalex=None):
    """
    Versión 2025: compatible con la API y entorno local.
    Usa headers completos y maneja bloqueos 403.
    """
    resultados = []
    base_url = "https://api.openalex.org/sources"
    batch_size = 25

    if not issn_list:
        st.warning("No se encontraron ISSN válidos para consultar en OpenAlex.")
        return pd.DataFrame()

    if not correo_openalex or "@" not in correo_openalex:
        st.error("Por favor ingresa un correo institucional válido para usar la API de OpenAlex.")
        return pd.DataFrame()

    headers = {
        "User-Agent": f"CompareitorDashboard/1.0 (contact: {correo_openalex})",
        "From": correo_openalex,
        "Accept": "application/json"
    }

    progreso = st.progress(0)
    inicio = time.time()

    for i in range(0, len(issn_list), batch_size):
        lote = issn_list[i:i + batch_size]
        filtro = ",".join(lote)
        url = f"{base_url}?filter=issn:{filtro}&mailto={correo_openalex}"

        try:
            r = requests.get(url, headers=headers, timeout=30)
            if r.status_code == 200:
                data = r.json()
                for item in data.get("results", []):
                    resultados.append({
                        "Título": item.get("display_name", ""),
                        "ISSN": item.get("issn_l", ""),
                        "Acceso abierto": "✅ Sí" if item.get("is_oa") else "❌ No"
                    })
            elif r.status_code == 403:
                st.warning(
                    f"OpenAlex devolvió 403 en el lote {i//batch_size + 1}. "
                    "Esto ocurre cuando se ejecuta desde localhost o IP privada.\n"
                    "👉 Prueba subir el dashboard a Streamlit Cloud, donde la consulta sí funcionará."
                )
                break
            else:
                st.warning(f"Error {r.status_code} al consultar OpenAlex (lote {i//batch_size + 1}).")
            time.sleep(0.6)
        except Exception as e:
            st.error(f"Error en lote {i//batch_size + 1}: {e}")

        progreso.progress(min((i + batch_size) / len(issn_list), 1.0))

    progreso.empty()
    duracion = time.time() - inicio

    if resultados:
        st.success(f"✅ Consulta finalizada: {len(resultados)} resultados obtenidos en {duracion:.1f} s.")
    else:
        st.warning(f"⚠️ No se obtuvieron resultados desde OpenAlex (duración {duracion:.1f} s).")

    return pd.DataFrame(resultados)

# ======================================================
# PROCESO PRINCIPAL
# ======================================================
if archivos:
    dfs = [leer_excel(a) for a in archivos]
    nombres = [a.name for a in archivos]

    # === CONSULTA ÚNICA A OPENALEX ===
    if len(archivos) == 1 and consultar_solo_uno:
        st.subheader("Vista previa del archivo cargado")
        df = dfs[0]
        filas, columnas = df.shape
        st.markdown(f"{nombres[0]} — {filas} filas × {columnas} columnas")
        st.dataframe(df.head(10))
        columna_issn = st.selectbox("Selecciona la columna que contiene ISSN o E-ISSN", df.columns)
        if st.button("Consultar OpenAlex"):
            issn_unicos = obtener_issn_validos(df, columna_issn)
            if len(issn_unicos) > 0:
                st.info("Consultando OpenAlex, por favor espera...")
                df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)
                if not df_openalex.empty:
                    st.dataframe(df_openalex.head(10))
                    output = io.BytesIO()
                    df_openalex.to_excel(output, index=False, sheet_name="OpenAccess")
                    output.seek(0)
                    st.download_button(
                        "Descargar resultados de OpenAlex",
                        data=output,
                        file_name=f"OpenAlex_{nombres[0]}_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        st.stop()

    # === COMPARACIÓN ENTRE VARIOS ARCHIVOS ===
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
                "Selecciona las columnas clave para comparar:",
                columnas_comunes
            )

            if columnas_clave:
                for i in range(len(dfs)):
                    df = dfs[i].copy()
                    df[columnas_clave] = df[columnas_clave].fillna("")
                    mascara_valida = df[columnas_clave].apply(
                        lambda r: any(str(x).strip() not in ["", "nan", "None"] for x in r),
                        axis=1
                    )
                    df = df[mascara_valida]
                    df["__clave__"] = df.apply(
                        lambda r: generar_clave_combinada(r, columnas_clave, normalizar=True),
                        axis=1
                    )
                    df = df.dropna(subset=["__clave__"])
                    df = df[df["__clave__"] != ""]
                    dfs[i] = df.reset_index(drop=True)

                filas_totales = sum(len(df) for df in dfs)
                claves = pd.concat([df[["__clave__"]] for df in dfs], keys=range(len(dfs)))
                claves["Archivo"] = claves.index.get_level_values(0)
                conteo = claves.groupby("__clave__")["Archivo"].nunique()

                claves_comunes = conteo[conteo > 1].index
                coincidencias_total = pd.concat([
                    df[df["__clave__"].isin(claves_comunes)] for df in dfs
                ])
                exclusivos_por_archivo = [
                    df[df["__clave__"].isin(conteo[conteo == 1].index)] for df in dfs
                ]

                total_exclusivos = sum(len(df) for df in exclusivos_por_archivo)

                # === RESUMEN ===
                st.divider()
                st.subheader("Resumen general")
                c1, c2, c3 = st.columns(3)
                c1.metric("Archivos cargados", len(archivos))
                c2.metric("Coincidencias encontradas", len(coincidencias_total))
                c3.metric("Registros exclusivos", total_exclusivos)
                st.caption(f"Filas analizadas (válidas): {filas_totales}")

                # === GRÁFICOS ===
                fig1 = px.pie(
                    pd.DataFrame({
                        "Tipo": ["Coincidencias", "Exclusivos"],
                        "Cantidad": [len(coincidencias_total), total_exclusivos]
                    }),
                    names="Tipo", values="Cantidad", title="Distribución general"
                )
                st.plotly_chart(fig1, use_container_width=True)

                resumen_exclusivos = pd.DataFrame({
                    "Archivo": nombres,
                    "Registros exclusivos": [len(df) for df in exclusivos_por_archivo]
                })
                fig2 = px.bar(
                    resumen_exclusivos,
                    x="Archivo",
                    y="Registros exclusivos",
                    title="Registros exclusivos por archivo",
                    text_auto=True,
                    color="Archivo"
                )
                st.plotly_chart(fig2, use_container_width=True)

                # === CONSULTA OPENALEX ===
                df_openalex = pd.DataFrame()
                if usar_openalex:
                    st.info("Consultando OpenAlex para coincidencias...")
                    if "ISSN" in coincidencias_total.columns:
                        issn_unicos = obtener_issn_validos(coincidencias_total, "ISSN")
                        if len(issn_unicos) > 0:
                            df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)

                # === GENERACIÓN DE EXCEL ===
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
                                    "Columnas clave utilizadas",
                                    "Coincidencias encontradas",
                                    "Registros exclusivos totales",
                                    "Filas analizadas"
                                ],
                                "Valor": [
                                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    modo,
                                    ", ".join(nombres),
                                    ", ".join(columnas_clave),
                                    len(coincidencias_total),
                                    total_exclusivos,
                                    filas_totales
                                ]
                            })
                            resumen.to_excel(writer, sheet_name="Resumen", index=False)

                            coincidencias_unicas = coincidencias_total.drop_duplicates(subset="__clave__", keep="first")
                            columnas_salida = [col for col in coincidencias_unicas.columns if col in ("Titulo", "__clave__")]
                            coincidencias_salida = coincidencias_unicas[columnas_salida].rename(
                                columns={"__clave__": "Clave usada"}
                            )

                            if not df_openalex.empty:
                                coincidencias_salida = coincidencias_salida.merge(
                                    df_openalex,
                                    how="left",
                                    left_on="Clave usada",
                                    right_on="ISSN"
                                ).drop(columns=["ISSN"])
                            coincidencias_salida.to_excel(writer, sheet_name="Coincidencias", index=False)

                            for i, exclusivos in enumerate(exclusivos_por_archivo):
                                nombre_limpio = os.path.splitext(nombres[i])[0]
                                nombre_limpio = "".join(c for c in nombre_limpio if c.isalnum() or c in (" ", "_", "-"))
                                nombre_hoja = f"Exclusivos_{nombre_limpio}"[:31]
                                exclusivos.to_excel(writer, sheet_name=nombre_hoja, index=False)

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
    st.info("Sube al menos un archivo Excel para comenzar la comparación.")
