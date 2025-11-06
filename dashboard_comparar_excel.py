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
Esta herramienta permite comparar varios archivos Excel (.xlsx o .xls),
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
    "Sube uno o más archivos Excel (.xlsx o .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# ======================================================
# FUNCIONES AUXILIARES
# ======================================================
@st.cache_data
def leer_excel(archivo):
    """Lee un archivo Excel (.xlsx o .xls) y elimina filas vacías."""
    try:
        nombre = archivo.name.lower()
        if nombre.endswith(".xls"):
            df = pd.read_excel(archivo, engine="xlrd")
        else:
            df = pd.read_excel(archivo, engine="openpyxl")
        df = df.dropna(how="all")
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Error al leer {archivo.name}: {e}")
        return pd.DataFrame()


def normalizar_valor(valor):
    """Normaliza ISSN, ISBN, EISSN, etc."""
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = valor.replace(" ", "").replace(".", "")
    # mantener el guión si está presente
    if len(valor) == 9 and valor[4] == "-":
        return valor
    # insertar guión si tiene 8 dígitos
    if valor.isdigit() and len(valor) == 8:
        return valor[:4] + "-" + valor[4:]
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


def consultar_openalex_batch(issn_list, correo_openalex=None):
    """Consulta OpenAlex en lotes de 25 ISSN válidos."""
    resultados = []
    base_url = "https://api.openalex.org/sources"
    batch_size = 25

    issn_validos = [i for i in issn_list if len(i) == 9 and "-" in i]
    if not issn_validos:
        st.warning("No se encontraron ISSN válidos para consultar en OpenAlex.")
        return pd.DataFrame()

    if not correo_openalex or "@" not in correo_openalex:
        st.error("Por favor ingresa un correo institucional válido para usar la API de OpenAlex.")
        return pd.DataFrame()

    headers = {
        "User-Agent": f"Compareitor/1.0 (Jorge Andrés Moreno Quintanilla; mailto:{correo_openalex})",
        "From": correo_openalex,
        "Accept": "application/json",
        "Accept-Language": "es-CL,es;q=0.9"
    }

    progreso = st.progress(0)
    inicio = time.time()

    for i in range(0, len(issn_validos), batch_size):
        lote = issn_validos[i:i + batch_size]
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
                st.error("❌ Error 403: OpenAlex bloqueó la solicitud. Prueba desde Streamlit Cloud o con otro correo institucional.")
                break
            else:
                st.warning(f"Error {r.status_code} en lote {i//batch_size + 1}.")
            time.sleep(0.6)
        except Exception as e:
            st.error(f"Error consultando lote {i//batch_size + 1}: {e}")

        progreso.progress(min((i + batch_size) / len(issn_validos), 1.0))

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

    # CASO 1: CONSULTAR UN SOLO ARCHIVO
    if len(archivos) == 1 and consultar_solo_uno:
        st.subheader("Vista previa del archivo cargado")
        df = dfs[0]
        filas, columnas = df.shape
        st.markdown(f"{nombres[0]} — {filas} filas × {columnas} columnas")
        st.dataframe(df.head(10))

        columna_issn = st.selectbox("Selecciona la columna que contiene ISSN o E-ISSN", df.columns)
        if st.button("Consultar OpenAlex"):
            issn_unicos = (
                df[columna_issn]
                .dropna()
                .astype(str)
                .apply(normalizar_valor)
                .unique()
                .tolist()
            )
            st.info("Consultando OpenAlex, por favor espera unos segundos...")
            df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)
            if not df_openalex.empty:
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
                            r, columnas_clave, normalizar=True
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

                if usar_openalex:
                    st.info("Consultando OpenAlex para coincidencias...")
                    if "ISSN" in coincidencias_total.columns:
                        issn_unicos = (
                            coincidencias_total["ISSN"]
                            .dropna()
                            .astype(str)
                            .apply(normalizar_valor)
                            .unique()
                            .tolist()
                        )
                        df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)
                        if not df_openalex.empty:
                            st.dataframe(df_openalex.head(20))
                    else:
                        st.warning("No se encontró columna 'ISSN' en los archivos para consultar OpenAlex.")

                # === DESCARGA FINAL ===
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    coincidencias_total.to_excel(writer, sheet_name="Coincidencias", index=False)
                    for i, exclusivos in enumerate(exclusivos_por_archivo):
                        nombre_limpio = os.path.splitext(nombres[i])[0][:25]
                        exclusivos.to_excel(writer, sheet_name=f"Exclusivos_{nombre_limpio}", index=False)
                    if usar_openalex and 'df_openalex' in locals() and not df_openalex.empty:
                        df_openalex.to_excel(writer, sheet_name="OpenAlex_Resultados", index=False)
                output.seek(0)

                st.download_button(
                    label="Descargar archivo Excel con resultados",
                    data=output,
                    file_name=f"resultado_comparacion_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
