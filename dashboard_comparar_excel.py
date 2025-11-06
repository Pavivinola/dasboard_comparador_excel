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
comparar_fechas = st.sidebar.checkbox("Comparar fechas", value=False)

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
    if len(valor) == 9 and valor[4] == "-":
        return valor
    if valor.isdigit() and len(valor) == 8:
        return valor[:4] + "-" + valor[4:]
    return valor


def formatear_issn_para_api(issn):
    """Formatea ISSN para la API de OpenAlex (formato XXXX-XXXX)."""
    issn_limpio = str(issn).replace("-", "").replace(" ", "").strip()
    if len(issn_limpio) == 8 and issn_limpio.isdigit():
        return f"{issn_limpio[:4]}-{issn_limpio[4:]}"
    # Si ya tiene guión en posición correcta
    if len(issn) == 9 and issn[4] == "-":
        return issn
    return None


def generar_clave_prioritaria(row, columnas, normalizar=False):
    """Devuelve la primera columna con valor válido, con o sin normalización."""
    for col in columnas:
        valor = row[col]
        if normalizar:
            valor = normalizar_valor(valor)
        if valor and str(valor).lower() != "nan":
            return valor
    return None


def obtener_issn_de_dataframe(df):
    """Extrae todos los ISSN válidos de un DataFrame buscando en columnas relevantes."""
    issn_list = []
    
    # Buscar columnas que contengan ISSN
    columnas_issn = [col for col in df.columns if 'ISSN' in col.upper() or 'E-ISSN' in col.upper()]
    
    st.write(f"🔍 Columnas ISSN detectadas: {columnas_issn}")
    
    for col in columnas_issn:
        valores = df[col].dropna().astype(str).unique()
        for val in valores:
            issn_formateado = formatear_issn_para_api(val)
            if issn_formateado:
                issn_list.append(issn_formateado)
    
    # Eliminar duplicados
    issn_list = list(set(issn_list))
    st.write(f"📋 Total ISSN válidos encontrados: {len(issn_list)}")
    if issn_list:
        st.write(f"📝 Primeros ISSN: {issn_list[:10]}")
    
    return issn_list


def consultar_openalex_batch(issn_list, correo_openalex=None):
    """Consulta OpenAlex en lotes de 50 ISSN válidos (versión corregida)."""
    resultados = []
    base_url = "https://api.openalex.org/sources"
    batch_size = 50

    if not issn_list:
        st.warning("⚠️ No se encontraron ISSN válidos para consultar en OpenAlex.")
        return pd.DataFrame()

    if not correo_openalex or "@" not in correo_openalex:
        st.error("❌ Por favor ingresa un correo institucional válido para usar la API de OpenAlex.")
        return pd.DataFrame()

    headers = {
        "User-Agent": f"Compareitor/1.0 (mailto:{correo_openalex})",
        "From": correo_openalex,
        "Accept": "application/json",
        "Connection": "keep-alive"
    }

    progreso = st.progress(0)
    status_text = st.empty()
    inicio = time.time()
    
    total_lotes = (len(issn_list) + batch_size - 1) // batch_size
    st.info(f"🔄 Consultando {len(issn_list)} ISSN en {total_lotes} lotes...")

    for i in range(0, len(issn_list), batch_size):
        lote = issn_list[i:i + batch_size]
        # CRÍTICO: usar pipe | para separar múltiples valores del mismo filtro
        filtro = "|".join(lote)
        url = f"{base_url}?filter=issn:{filtro}&mailto={correo_openalex}&per_page=200"

        status_text.text(f"Consultando lote {i//batch_size + 1} de {total_lotes} ({len(lote)} ISSN)...")
        
        try:
            r = requests.get(url, headers=headers, timeout=60)
            st.write(f"🌐 Lote {i//batch_size + 1}: HTTP {r.status_code}")
            
            if r.status_code == 200:
                data = r.json()
                items = data.get("results", [])
                st.write(f"✅ Obtenidos {len(items)} resultados en este lote")
                
                for item in items:
                    resultados.append({
                        "Título": item.get("display_name", ""),
                        "ISSN": item.get("issn_l", ""),
                        "ISSN_Alternos": ", ".join(item.get("issn", [])),
                        "Acceso abierto": "✅ Sí" if item.get("is_oa") else "❌ No",
                        "Editorial": item.get("host_organization_name", ""),
                        "País": item.get("country_code", ""),
                        "Tipo": item.get("type", ""),
                        "Works_Count": item.get("works_count", 0),
                        "Cited_By_Count": item.get("cited_by_count", 0),
                        "OpenAlex_ID": item.get("id", "")
                    })
            elif r.status_code == 403:
                st.error("❌ OpenAlex devolvió 403: verifica tu correo institucional.")
                st.write(f"URL intentada: {url[:150]}...")
                break
            elif r.status_code == 429:
                st.warning("⚠️ Límite de tasa excedido. Esperando 5 segundos...")
                time.sleep(5)
                continue
            else:
                st.warning(f"⚠️ Error {r.status_code} al consultar lote {i//batch_size + 1}")
                st.write(f"Respuesta: {r.text[:300]}")
                
            time.sleep(1)  # Respetar límites de API
            
        except Exception as e:
            st.error(f"❌ Error consultando OpenAlex: {e}")

        progreso.progress(min((i + batch_size) / len(issn_list), 1.0))

    progreso.empty()
    status_text.empty()
    duracion = time.time() - inicio
    
    if resultados:
        st.success(f"✅ Consulta finalizada: {len(resultados)} resultados obtenidos en {duracion:.1f} s.")
    else:
        st.warning(f"⚠️ No se obtuvieron resultados desde OpenAlex (duración {duracion:.1f} s).")

    return pd.DataFrame(resultados)


# ======================================================
# FUNCIÓN: PROCESAMIENTO DE FECHAS (VERSIÓN FINAL)
# ======================================================
def procesar_fechas(df):
    """
    Genera la columna 'Rango Calculado' según las reglas definidas:
    - Si existe 'Fecha Rango', se copia directamente.
    - Si existen 'Fecha Inicio', 'Fecha Termino' y 'Retraso':
        * Se toma solo el año de las fechas (formato datetime o texto mm/dd/aaaa).
        * Si 'Fecha Termino' está vacía, se usa el año actual.
        * Si 'Retraso' tiene valor (en meses), se convierte a años y se resta al año final.
        * Se genera un rango tipo "AAAA_inicio - AAAA_final".
    """
    año_actual = datetime.now().year

    # Caso 1: si hay una columna "Fecha Rango"
    if "Fecha Rango" in df.columns:
        df["Rango Calculado"] = df["Fecha Rango"]
        return df

    # Caso 2: si existen las tres columnas requeridas
    if all(c in df.columns for c in ["Fecha Inicio", "Fecha Termino", "Retraso"]):
        import re

        def obtener_año(valor):
            """Devuelve el año de una celda, sin importar si es datetime o texto."""
            if pd.isna(valor):
                return None
            if isinstance(valor, (datetime, pd.Timestamp)):
                return valor.year
            valor_str = str(valor)
            match = re.search(r"(19|20)\d{2}", valor_str)
            if match:
                return int(match.group(0))
            return None

        def calcular_rango(row):
            año_inicio = obtener_año(row["Fecha Inicio"])
            año_final = obtener_año(row["Fecha Termino"]) or año_actual

            # Si hay retraso, ajustarlo en años
            retraso = 0
            try:
                if pd.notna(row["Retraso"]) and str(row["Retraso"]).strip() != "":
                    retraso = int(float(row["Retraso"])) // 12
            except Exception:
                retraso = 0

            año_final_ajustado = año_final - retraso if año_final else año_actual

            if año_inicio:
                return f"{año_inicio} - {año_final_ajustado}"
            return None

        df["Rango Calculado"] = df.apply(calcular_rango, axis=1)

    return df


# ======================================================
# PROCESO PRINCIPAL
# ======================================================
if archivos:
    dfs = [leer_excel(a) for a in archivos]
    nombres = [a.name for a in archivos]

    # === CASO 1: UN SOLO ARCHIVO ===
    if len(archivos) == 1 and consultar_solo_uno:
        st.subheader("Vista previa del archivo cargado")
        df = dfs[0]
        st.dataframe(df.head(10))
        
        columna_issn = st.selectbox("Selecciona la columna ISSN o E-ISSN", df.columns)
        
        if st.button("Consultar OpenAlex"):
            issn_unicos = []
            valores = df[columna_issn].dropna().astype(str).unique()
            
            for val in valores:
                issn_fmt = formatear_issn_para_api(val)
                if issn_fmt:
                    issn_unicos.append(issn_fmt)
            
            if not issn_unicos:
                st.warning("⚠️ No se encontraron ISSN válidos en la columna seleccionada.")
            else:
                st.info("🔄 Consultando OpenAlex, por favor espera...")
                df_openalex = consultar_openalex_batch(issn_unicos, correo_openalex)
                
                if not df_openalex.empty:
                    st.subheader("📊 Resultados de OpenAlex")
                    st.dataframe(df_openalex)
                    
                    # Botón de descarga
                    output = io.BytesIO()
                    df_openalex.to_excel(output, index=False, sheet_name="OpenAlex")
                    output.seek(0)
                    
                    st.download_button(
                        "📥 Descargar resultados OpenAlex",
                        data=output,
                        file_name=f"OpenAlex_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        st.stop()

    # === CASO 2: MÚLTIPLES ARCHIVOS ===
    if len(archivos) > 1:
        st.subheader("Vista previa de los archivos cargados")
        for nombre, df in zip(nombres, dfs):
            st.markdown(f"**{nombre}** — {df.shape[0]} filas × {df.shape[1]} columnas")
            st.dataframe(df.head(10))
            st.markdown("---")

        columnas_comunes = list(set.intersection(*(set(df.columns) for df in dfs)))

        if columnas_comunes:
            columnas_clave = st.multiselect("Selecciona las columnas clave para comparar", columnas_comunes)

            if columnas_clave:
                for df in dfs:
                    df[columnas_clave] = df[columnas_clave].fillna("")
                    df["__clave__"] = df.apply(
                        lambda r: generar_clave_prioritaria(r, columnas_clave, normalizar=True),
                        axis=1,
                    )
                    df.dropna(subset=["__clave__"], inplace=True)

                claves = pd.concat([df[["__clave__"]] for df in dfs], keys=range(len(dfs)))
                claves["Archivo"] = claves.index.get_level_values(0)
                conteo = claves.groupby("__clave__")["Archivo"].nunique()

                claves_comunes = conteo[conteo > 1].index
                coincidencias_total = pd.concat(
                    [df[df["__clave__"].isin(claves_comunes)] for df in dfs]
                ).drop(columns=["__clave__"])

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

                # === GRÁFICOS ===
                fig1 = px.pie(
                    pd.DataFrame({"Tipo": ["Coincidencias", "Exclusivos"],
                                  "Cantidad": [len(coincidencias_total), total_exclusivos]}),
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

                # === COMPARACIÓN DE FECHAS ===
                if comparar_fechas:
                    st.info("🗓️ Procesando columnas de fechas en coincidencias...")
                    coincidencias_total = procesar_fechas(coincidencias_total)
                    st.success("✅ Procesamiento de fechas completado.")
                    st.dataframe(coincidencias_total.head(10))

                # === CONSULTA OPENALEX (CORREGIDO) ===
                df_openalex = pd.DataFrame()
                if usar_openalex:
                    st.divider()
                    st.subheader("🔍 Consultando OpenAlex")
                    st.info("Extrayendo ISSN de las coincidencias...")
                    
                    issn_list = obtener_issn_de_dataframe(coincidencias_total)
                    
                    if issn_list:
                        df_openalex = consultar_openalex_batch(issn_list, correo_openalex)
                        
                        if not df_openalex.empty:
                            st.subheader("📊 Resultados de OpenAlex")
                            st.dataframe(df_openalex)
                    else:
                        st.warning("⚠️ No se encontraron ISSN válidos en las coincidencias.")

                # === DESCARGA FINAL ===
                st.divider()
                st.subheader("📥 Descargar resultados")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    # Resumen
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
                    
                    # Coincidencias
                    coincidencias_total.to_excel(writer, sheet_name="Coincidencias", index=False)
                    
                    # Exclusivos por archivo
                    for i, exclusivos in enumerate(exclusivos_por_archivo):
                        nombre_limpio = os.path.splitext(nombres[i])[0]
                        nombre_limpio = "".join(c for c in nombre_limpio if c.isalnum() or c in (" ", "_", "-"))
                        nombre_hoja = f"Exclusivos_{nombre_limpio}"[:31]
                        exclusivos.to_excel(writer, sheet_name=nombre_hoja, index=False)
                    
                    # OpenAlex si existe
                    if not df_openalex.empty:
                        df_openalex.to_excel(writer, sheet_name="OpenAlex", index=False)
                        
                output.seek(0)

                st.download_button(
                    "📥 Descargar archivo Excel con resultados",
                    data=output,
                    file_name=f"resultado_comparacion_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ No se encontraron columnas comunes entre los archivos.")
else:
    st.info("📂 Sube al menos un archivo Excel para comenzar.")