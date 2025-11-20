import streamlit as st
import pandas as pd
import io
import os
import requests
import time
from datetime import datetime
import plotly.express as px
import re
from io import BytesIO

# ======================================================
# CONFIGURACIÓN DE LA PÁGINA
# ======================================================
st.set_page_config(page_title="Comparador de Excels", layout="wide")
st.title("Compareitor")
st.markdown("""
Esta herramienta permite comparar varios archivos Excel (.xlsx o .xls),
detectar coincidencias, encontrar registros exclusivos,
analizar cobertura temporal y consultar información en OpenAlex.
""")
st.divider()

# ======================================================
# PANEL LATERAL
# ======================================================
st.sidebar.header("Configuración")

modo = st.sidebar.radio(
    "Selecciona el modo de ejecución:",
    ["Rápido", "Avanzado"],
    help="**Rápido**: Análisis básico y rápido\n**Avanzado**: Todas las opciones de análisis disponibles"
)

# Mostrar descripción del modo seleccionado
if modo == "Rápido":
    st.sidebar.info("**Modo Rápido**: Comparación básica, visualizaciones esenciales")
else:
    st.sidebar.success("**Modo Avanzado**: Análisis completo con todas las opciones")

st.sidebar.markdown("---")

# Opciones según el modo
if modo == "Avanzado":
    st.sidebar.subheader("Análisis sobre coincidencias")
    comparar_fechas = st.sidebar.checkbox("Análisis temporal y referenciales", value=False)
    usar_openalex = st.sidebar.checkbox("Consultar OpenAlex (batch)", value=False)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Análisis archivo individual")
    consultar_solo_uno = st.sidebar.checkbox("Consultar OpenAlex para un archivo", value=False)
    analizar_tiempo_individual = st.sidebar.checkbox("Análisis temporal y referencial para un archivo", value=False)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Opciones avanzadas")
    normalizar_datos = st.sidebar.checkbox("Normalizar ISSN/ISBN automáticamente", value=True)
    mostrar_metricas_detalladas = st.sidebar.checkbox("Mostrar métricas detalladas", value=True)
else:
    # Modo Rápido: valores predeterminados
    comparar_fechas = False
    usar_openalex = False
    consultar_solo_uno = False
    analizar_tiempo_individual = False
    normalizar_datos = True
    mostrar_metricas_detalladas = False
    umbral_similitud = 100

correo_openalex = st.sidebar.text_input(
    "Correo para OpenAlex (recomendado)",
    placeholder="tucorreo@institucion.cl",
    help="Necesario para usar la API de OpenAlex"
)

archivos = st.sidebar.file_uploader(
    "Sube uno o más archivos Excel (.xlsx o .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# ======================================================
# FUNCIONES AUXILIARES
# ======================================================
def crear_excel_descargable(dataframes_dict):
    """
    Crea un archivo Excel con múltiples hojas a partir de un diccionario de DataFrames.
    
    Args:
        dataframes_dict: Diccionario con formato {nombre_hoja: dataframe}
    
    Returns:
        BytesIO: Objeto en memoria con el archivo Excel
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for nombre_hoja, df in dataframes_dict.items():
            if df is not None and not df.empty:
                # Limitar nombre de hoja a 31 caracteres (límite de Excel)
                nombre_hoja_limpio = str(nombre_hoja)[:31]
                df.to_excel(writer, sheet_name=nombre_hoja_limpio, index=False)
    output.seek(0)
    return output


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
    """Extrae todos los ISSN válidos de un DataFrame."""
    issn_list = []
    columnas_issn = [col for col in df.columns if 'ISSN' in col.upper() or 'E-ISSN' in col.upper()]
    for col in columnas_issn:
        valores = df[col].dropna().astype(str).unique()
        for val in valores:
            issn_formateado = formatear_issn_para_api(val)
            if issn_formateado:
                issn_list.append(issn_formateado)
    return list(set(issn_list))


def consultar_openalex_batch(issn_list, correo_openalex=None):
    """Consulta OpenAlex en lotes de 50 ISSN válidos."""
    resultados = []
    base_url = "https://api.openalex.org/sources"
    batch_size = 50

    if not issn_list:
        st.warning("⚠ No se encontraron ISSN válidos para consultar en OpenAlex.")
        return pd.DataFrame()

    if not correo_openalex or "@" not in correo_openalex:
        st.error("⚠ Por favor ingresa un correo institucional válido para usar la API de OpenAlex.")
        return pd.DataFrame()

    headers = {"User-Agent": f"Compareitor/1.0 (mailto:{correo_openalex})"}
    progreso = st.progress(0)
    status_text = st.empty()
    inicio = time.time()

    total_lotes = (len(issn_list) + batch_size - 1) // batch_size
    for i in range(0, len(issn_list), batch_size):
        lote = issn_list[i:i + batch_size]
        filtro = "|".join(lote)
        url = f"{base_url}?filter=issn:{filtro}&mailto={correo_openalex}&per_page=200"

        status_text.text(f"🔄 Lote {i//batch_size + 1} de {total_lotes} ({len(lote)} ISSN)...")
        try:
            r = requests.get(url, headers=headers, timeout=60)
            if r.status_code == 200:
                data = r.json()
                for item in data.get("results", []):
                    resultados.append({
                        "Título": item.get("display_name", ""),
                        "ISSN": item.get("issn_l", ""),
                        "Acceso abierto": "✅ Sí" if item.get("is_oa") else "❌ No",
                        "Editorial": item.get("host_organization_name", ""),
                        "País": item.get("country_code", ""),
                        "Tipo": item.get("type", ""),
                        "Works_Count": item.get("works_count", 0),
                        "Cited_By_Count": item.get("cited_by_count", 0),
                        "OpenAlex_ID": item.get("id", "")
                    })
            elif r.status_code == 429:
                time.sleep(5)
                continue
            time.sleep(1)
        except Exception as e:
            st.error(f"❌ Error consultando OpenAlex: {e}")
        progreso.progress(min((i + batch_size) / len(issn_list), 1.0))

    progreso.empty()
    status_text.empty()
    st.success(f"✅ Consulta finalizada ({len(resultados)} resultados obtenidos en {time.time()-inicio:.1f}s)")
    return pd.DataFrame(resultados)


# ======================================================
# FUNCIÓN DE FECHAS
# ======================================================
def procesar_fechas(df):
    """Genera la columna 'Rango Calculado' según las reglas definidas."""
    año_actual = datetime.now().year

    if "Fecha Rango" in df.columns:
        df["Rango Calculado"] = df["Fecha Rango"]
        return df

    if all(c in df.columns for c in ["Fecha Inicio", "Fecha Termino", "Retraso"]):
        def obtener_año(valor):
            if pd.isna(valor):
                return None
            if isinstance(valor, (datetime, pd.Timestamp)):
                return valor.year
            valor_str = str(valor)
            match = re.search(r"(19|20)\d{2}", valor_str)
            return int(match.group(0)) if match else None

        def calcular_rango(row):
            año_inicio = obtener_año(row["Fecha Inicio"])
            año_final = obtener_año(row["Fecha Termino"]) or año_actual
            retraso = 0
            try:
                if pd.notna(row["Retraso"]) and str(row["Retraso"]).strip() != "":
                    retraso = int(float(row["Retraso"])) // 12
            except Exception:
                retraso = 0
            año_final_ajustado = año_final - retraso
            return f"{año_inicio} - {año_final_ajustado}" if año_inicio else None

        df["Rango Calculado"] = df.apply(calcular_rango, axis=1)

    return df


def tiene_fecha_valida(valor):
    """Verifica si un valor tiene una fecha válida (año de 4 dígitos)."""
    if pd.isna(valor) or str(valor).strip() == "":
        return False
    try:
        texto = str(valor)
        return bool(re.search(r"(19|20)\d{2}", texto))
    except Exception:
        return False


# ======================================================
# ANÁLISIS DE COINCIDENCIAS - FECHAS Y REFERENCIALES
# ======================================================
def analizar_fechas_coincidencias(coincidencias_df, modo_avanzado=False, resultados_dict=None):
    """Analiza fechas y detecta referenciales EN LAS COINCIDENCIAS."""
    st.divider()
    st.subheader(" Análisis temporal y detección de registros referenciales")
    st.caption("Este análisis se realiza SOLO sobre las coincidencias encontradas")
    
    # Procesar fechas
    coincidencias_df = procesar_fechas(coincidencias_df.copy())
    
    # Detectar referenciales (sin Fecha Inicio válida)
    if "Fecha Inicio" in coincidencias_df.columns:
        coincidencias_df["Es Referencial"] = ~coincidencias_df["Fecha Inicio"].apply(tiene_fecha_valida)
    else:
        st.warning("⚠ No se encontró la columna 'Fecha Inicio'. No se puede detectar referenciales.")
        return coincidencias_df
    
    # ---- 1) Análisis de referenciales por archivo ----
    st.markdown("###  Recursos referenciales por archivo")
    st.caption("Registros sin Fecha Inicio válida (recursos de referencia continua)")
    
    total_por_archivo = coincidencias_df["Archivo"].value_counts()
    referenciales_por_archivo = coincidencias_df.groupby("Archivo")["Es Referencial"].sum()
    
    df_referenciales = pd.DataFrame({
        "Archivo": total_por_archivo.index,
        "Total Coincidencias": total_por_archivo.values,
        "Referenciales": referenciales_por_archivo.reindex(total_por_archivo.index, fill_value=0).values
    })
    df_referenciales["% Referenciales"] = (
        df_referenciales["Referenciales"] / df_referenciales["Total Coincidencias"] * 100
    ).round(1)
    
    # Guardar en resultados
    if resultados_dict is not None:
        resultados_dict["Análisis_Referenciales"] = df_referenciales
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.dataframe(df_referenciales, use_container_width=True)
    with col2:
        total_ref = df_referenciales["Referenciales"].sum()
        total_coincidencias = df_referenciales["Total Coincidencias"].sum()
        st.metric("Total Referenciales", f"{total_ref} ({(total_ref/total_coincidencias*100):.1f}%)")
        st.metric("Con Fechas", f"{total_coincidencias - total_ref}")
    
    # Gráfico de referenciales (más detallado en modo avanzado)
    if modo_avanzado:
        fig_ref = px.bar(
            df_referenciales,
            x="Archivo",
            y=["Referenciales", "Total Coincidencias"],
            title="Distribución de registros referenciales vs totales",
            barmode="group",
            color_discrete_map={"Referenciales": "#E74C3C", "Total Coincidencias": "#3498DB"}
        )
        st.plotly_chart(fig_ref, use_container_width=True)
    else:
        # Gráfico simple en modo rápido
        fig_ref_simple = px.bar(
            df_referenciales,
            x="Archivo",
            y="Referenciales",
            title="Registros referenciales por archivo",
            color="Referenciales",
            color_continuous_scale="Reds"
        )
        st.plotly_chart(fig_ref_simple, use_container_width=True)
    
    # ---- 2) Análisis de cobertura temporal ----
    coincidencias_temporales = coincidencias_df[coincidencias_df["Es Referencial"] == False].copy()
    
    if coincidencias_temporales.empty:
        st.warning("⚠ No hay registros con Fecha Inicio válida para calcular cobertura.")
        return coincidencias_df
    
    st.markdown("###  Análisis de cobertura temporal")
    st.caption(f"Análisis sobre {len(coincidencias_temporales)} registros con fechas válidas")
    
    if "Rango Calculado" in coincidencias_temporales.columns:
        coincidencias_temporales["Año Inicio"] = (
            coincidencias_temporales["Rango Calculado"]
            .astype(str)
            .str.extract(r"(\d{4})", expand=False)
            .astype(float)
        )
        coincidencias_temporales["Año Fin"] = (
            coincidencias_temporales["Rango Calculado"]
            .astype(str)
            .str.extract(r"-\s*(\d{4})", expand=False)
            .astype(float)
        )
        coincidencias_temporales["Duración (años)"] = (
            coincidencias_temporales["Año Fin"] - coincidencias_temporales["Año Inicio"]
        )
        
        coincidencias_temporales = coincidencias_temporales.dropna(subset=["Año Inicio", "Año Fin"])
        
        if not coincidencias_temporales.empty:
            df_cobertura = coincidencias_temporales.groupby("Archivo", dropna=False).agg({
                "Duración (años)": ["mean", "min", "max"] if modo_avanzado else "mean",
                "Rango Calculado": "count"
            }).reset_index()
            
            if modo_avanzado:
                df_cobertura.columns = ["Archivo", "Promedio duración (años)", "Min duración", "Max duración", "Registros analizados"]
            else:
                df_cobertura.columns = ["Archivo", "Promedio duración (años)", "Registros analizados"]
            
            # Índice de cobertura (ponderado)
            df_cobertura["Índice Cobertura"] = (
                df_cobertura["Promedio duración (años)"].rank(pct=True) * 0.6 +
                df_cobertura["Registros analizados"].rank(pct=True) * 0.4
            ).round(2)
            
            # Guardar en resultados
            if resultados_dict is not None:
                resultados_dict["Cobertura_Temporal"] = df_cobertura
            
            st.dataframe(df_cobertura.style.format({
                "Promedio duración (años)": "{:.1f}",
                "Min duración": "{:.1f}" if modo_avanzado else None,
                "Max duración": "{:.1f}" if modo_avanzado else None,
                "Índice Cobertura": "{:.2f}"
            }), use_container_width=True)
            
            fig_cobertura = px.bar(
                df_cobertura,
                x="Archivo",
                y="Índice Cobertura",
                text_auto=True,
                color="Archivo",
                title="Índice de cobertura por archivo (mayor es mejor)",
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            st.plotly_chart(fig_cobertura, use_container_width=True)
    
    return coincidencias_df


# ======================================================
# ANÁLISIS DE COINCIDENCIAS - OPENALEX
# ======================================================
def analizar_openalex_coincidencias(coincidencias_df, correo, modo_avanzado=False, resultados_dict=None):
    """Consulta OpenAlex SOLO para las coincidencias."""
    st.divider()
    st.subheader("🔍 Consulta OpenAlex sobre coincidencias")
    st.caption("Consultando información de las revistas/recursos encontrados en las coincidencias")
    
    # Extraer ISSN de las coincidencias
    issn_list = obtener_issn_de_dataframe(coincidencias_df)
    
    if not issn_list:
        st.warning("⚠ No se encontraron ISSN válidos en las coincidencias para consultar OpenAlex.")
        return
    
    st.info(f" Se encontraron {len(issn_list)} ISSN únicos en las coincidencias")
    
    # Consultar OpenAlex
    df_openalex = consultar_openalex_batch(issn_list, correo)
    
    if df_openalex.empty:
        st.warning("⚠ No se obtuvieron resultados de OpenAlex.")
        return
    
    # Guardar en resultados
    if resultados_dict is not None:
        resultados_dict["OpenAlex_Coincidencias"] = df_openalex
    
    # Mostrar resultados
    st.success(f"✅ Se obtuvieron {len(df_openalex)} resultados de OpenAlex")
    
    # Estadísticas rápidas
    col1, col2, col3 = st.columns(3)
    with col1:
        total_oa = (df_openalex["Acceso abierto"] == "✅ Sí").sum()
        st.metric("Acceso Abierto", f"{total_oa} ({total_oa/len(df_openalex)*100:.1f}%)")
    with col2:
        promedio_works = df_openalex["Works_Count"].mean()
        st.metric("Promedio Works", f"{promedio_works:.0f}")
    with col3:
        promedio_citas = df_openalex["Cited_By_Count"].mean()
        st.metric("Promedio Citas", f"{promedio_citas:.0f}")
    
    # Gráficos
    fig_oa = px.pie(
        df_openalex,
        names="Acceso abierto",
        title="Distribución de Acceso Abierto",
        color_discrete_map={"✅ Sí": "#2ECC71", "❌ No": "#E74C3C"}
    )
    st.plotly_chart(fig_oa, use_container_width=True)
    
    # Modo avanzado: más visualizaciones
    if modo_avanzado:
        # Top 10 por citas
        top_citadas = df_openalex.nlargest(10, "Cited_By_Count")
        fig_top = px.bar(
            top_citadas,
            x="Cited_By_Count",
            y="Título",
            orientation="h",
            title="Top 10 revistas más citadas",
            color="Cited_By_Count",
            color_continuous_scale="Blues"
        )
        st.plotly_chart(fig_top, use_container_width=True)
        
        # Distribución por país
        if not df_openalex["País"].isna().all():
            pais_count = df_openalex["País"].value_counts().head(10)
            fig_pais = px.bar(
                x=pais_count.values,
                y=pais_count.index,
                orientation="h",
                title="Top 10 países por número de revistas",
                labels={"x": "Cantidad", "y": "País"}
            )
            st.plotly_chart(fig_pais, use_container_width=True)
    
    # Tabla completa
    st.markdown("###  Resultados completos de OpenAlex")
    if modo_avanzado:
        st.dataframe(df_openalex, use_container_width=True)
    else:
        # Modo rápido: solo primeras 20 filas
        st.dataframe(df_openalex.head(20), use_container_width=True)
        if len(df_openalex) > 20:
            st.info(f"Mostrando 20 de {len(df_openalex)} resultados. Descarga el Excel para ver todos.")
    
    # Descargar resultados
    excel_buffer = crear_excel_descargable({"Resultados_OpenAlex": df_openalex})
    st.download_button(
        label=" Descargar resultados OpenAlex (Excel)",
        data=excel_buffer,
        file_name="openalex_coincidencias.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ======================================================
# ANÁLISIS ARCHIVO INDIVIDUAL - OPENALEX
# ======================================================
def analizar_openalex_individual(archivos, nombres, correo):
    """Consulta OpenAlex para un archivo individual seleccionado."""
    st.divider()
    st.subheader(" Consulta OpenAlex - Archivo Individual")
    
    archivo_seleccionado = st.selectbox(
        "Selecciona el archivo a consultar:",
        nombres
    )
    
    idx = nombres.index(archivo_seleccionado)
    df_seleccionado = leer_excel(archivos[idx])
    
    st.info(f" Archivo seleccionado: **{archivo_seleccionado}** ({len(df_seleccionado)} registros)")
    
    if st.button("🔍 Consultar OpenAlex", type="primary"):
        issn_list = obtener_issn_de_dataframe(df_seleccionado)
        
        if not issn_list:
            st.warning("⚠ No se encontraron ISSN válidos en este archivo.")
            return
        
        st.info(f"📋 Se encontraron {len(issn_list)} ISSN únicos")
        
        df_openalex = consultar_openalex_batch(issn_list, correo)
        
        if not df_openalex.empty:
            st.success(f"✅ Se obtuvieron {len(df_openalex)} resultados")
            st.dataframe(df_openalex, use_container_width=True)
            
            excel_buffer = crear_excel_descargable({"Resultados_OpenAlex": df_openalex})
            st.download_button(
                label=" Descargar resultados (Excel)",
                data=excel_buffer,
                file_name=f"openalex_{archivo_seleccionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ======================================================
# ANÁLISIS ARCHIVO INDIVIDUAL - FECHAS
# ======================================================
def analizar_fechas_individual(archivos, nombres, resultados=None):
    """Aplica el análisis temporal/referencial a un solo archivo."""
    st.divider()
    st.subheader(" Análisis temporal y referenciales - Archivo individual")
    
    archivo_seleccionado = st.selectbox(
        "Selecciona el archivo a analizar:",
        nombres,
        key="sel_arch_tiempo"
    )
    
    idx = nombres.index(archivo_seleccionado)
    df_sel = leer_excel(archivos[idx])
    
    st.info(f" Archivo analizado: **{archivo_seleccionado}** ({len(df_sel)} registros)")
    
    # Validar columnas necesarias
    columnas_necesarias = ["Fecha Inicio", "Fecha Termino", "Retraso"]
    columnas_faltantes = [col for col in columnas_necesarias if col not in df_sel.columns]
    
    if columnas_faltantes:
        st.warning(f"⚠️ El archivo no tiene las columnas necesarias: {', '.join(columnas_faltantes)}")
        st.info(" Las columnas deben llamarse exactamente: 'Fecha Inicio', 'Fecha Termino', 'Retraso'")
        return
    
    # Añadimos una columna 'Archivo' para reutilizar la lógica existente
    df_sel = df_sel.copy()
    df_sel["Archivo"] = archivo_seleccionado
    
    # Crear diccionario temporal para no sobrescribir resultados de comparación múltiple
    resultados_temporal = {}
    
    analizar_fechas_coincidencias(
        df_sel,
        modo_avanzado=True,
        resultados_dict=resultados_temporal
    )
    
    # Guardar con prefijo para diferenciarlo del análisis múltiple
    if resultados is not None and resultados_temporal:
        nombre_limpio = os.path.splitext(archivo_seleccionado)[0]
        for key, value in resultados_temporal.items():
            resultados[f"{key}_{nombre_limpio}"] = value
    
    # Botón de descarga individual
    if resultados_temporal:
        st.divider()
        nombre_limpio = os.path.splitext(archivo_seleccionado)[0]
        excel_individual = crear_excel_descargable(resultados_temporal)
        st.download_button(
            label=f" Descargar análisis temporal de {archivo_seleccionado}",
            data=excel_individual,
            file_name=f"analisis_temporal_{nombre_limpio}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ======================================================
# PROCESO PRINCIPAL
# ======================================================
if archivos:
    dfs = [leer_excel(a) for a in archivos]
    nombres = [a.name for a in archivos]
    
    # Diccionario para almacenar todos los resultados del análisis
    resultados_completos = {}
    
    # ---- ANÁLISIS INDIVIDUAL (solo modo avanzado) ----
    if consultar_solo_uno and len(archivos) > 0:
        analizar_openalex_individual(archivos, nombres, correo_openalex)
    
    if analizar_tiempo_individual and len(archivos) > 0:
        analizar_fechas_individual(archivos, nombres, resultados_completos)
    
    # ---- COMPARACIÓN MÚLTIPLE ----
    if len(archivos) > 1:
        # Vista previa según el modo
        if modo == "Avanzado":
            st.subheader(" Vista previa de los archivos cargados")
            for nombre, df in zip(nombres, dfs):
                with st.expander(f"**{nombre}** — {df.shape[0]} filas × {df.shape[1]} columnas"):
                    st.dataframe(df.head(10))
        else:
            # Modo rápido: solo resumen
            st.subheader(" Archivos cargados")
            resumen_archivos = pd.DataFrame({
                "Archivo": nombres,
                "Filas": [df.shape[0] for df in dfs],
                "Columnas": [df.shape[1] for df in dfs],
            })
            st.dataframe(resumen_archivos, use_container_width=True)
        
        columnas_comunes = list(set.intersection(*(set(df.columns) for df in dfs)))
        
        if columnas_comunes:
            columnas_clave = st.multiselect(
                " Selecciona las columnas clave para comparar",
                columnas_comunes,
                help="Selecciona las columnas que se usarán para identificar coincidencias"
            )
            
            if columnas_clave:
                # Generar claves y encontrar coincidencias
                for df in dfs:
                    df[columnas_clave] = df[columnas_clave].fillna("")
                    df["__clave__"] = df.apply(
                        lambda r: generar_clave_prioritaria(r, columnas_clave, normalizar=normalizar_datos),
                        axis=1,
                    )
                    df.dropna(subset=["__clave__"], inplace=True)
                
                claves = pd.concat(
                    [df[["__clave__"]] for df in dfs],
                    keys=range(len(dfs))
                )
                claves = claves.reset_index(level=0).rename(columns={"level_0": "IdxArchivo"})
                conteo = claves.groupby("__clave__")["IdxArchivo"].nunique()
                
                claves_comunes = conteo[conteo > 1].index
                claves_exclusivas = conteo[conteo == 1].index
                
                # Construir coincidencias y exclusivos
                coincidencias_por_archivo = []
                exclusivos_por_archivo = []
                
                for df, nombre in zip(dfs, nombres):
                    temp_coinc = df[df["__clave__"].isin(claves_comunes)].copy()
                    temp_coinc["Archivo"] = nombre
                    coincidencias_por_archivo.append(temp_coinc)
                    
                    temp_excl = df[df["__clave__"].isin(claves_exclusivas)].copy()
                    temp_excl["Archivo"] = nombre
                    temp_excl = temp_excl.drop(columns=["__clave__"])
                    exclusivos_por_archivo.append(temp_excl)
                
                coincidencias_total = pd.concat(coincidencias_por_archivo, ignore_index=True)
                coincidencias_total = coincidencias_total.drop(columns=["__clave__"], errors="ignore")
                
                total_exclusivos = sum(len(df) for df in exclusivos_por_archivo)
                total_registros = sum(len(df) for df in dfs)
                
                # Guardar coincidencias en resultados
                resultados_completos["Coincidencias"] = coincidencias_total
                
                # ---- RESUMEN GENERAL ----
                st.divider()
                st.subheader(" Resumen general")
                
                if modo == "Avanzado" or mostrar_metricas_detalladas:
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Archivos cargados", len(archivos))
                    c2.metric("Coincidencias", len(coincidencias_total))
                    c3.metric("Exclusivos", total_exclusivos)
                    c4.metric("Total registros", total_registros)
                else:
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Archivos cargados", len(archivos))
                    c2.metric("Coincidencias", len(coincidencias_total))
                    c3.metric("Exclusivos", total_exclusivos)
                
                df_resumen = pd.DataFrame(
                    [
                        {"Métrica": "Archivos cargados", "Valor": len(archivos)},
                        {"Métrica": "Coincidencias", "Valor": len(coincidencias_total)},
                        {"Métrica": "Exclusivos", "Valor": total_exclusivos},
                        {"Métrica": "Total registros", "Valor": total_registros},
                    ]
                )
                resultados_completos["Resumen_General"] = df_resumen
                
                fig_general = px.pie(
                    pd.DataFrame({
                        "Tipo": ["Coincidencias", "Exclusivos"],
                        "Cantidad": [len(coincidencias_total), total_exclusivos],
                    }),
                    names="Tipo",
                    values="Cantidad",
                    title="Distribución general de registros",
                    color="Tipo",
                    color_discrete_map={"Coincidencias": "#2ECC71", "Exclusivos": "#3498DB"},
                )
                fig_general.update_traces(textinfo="percent+value")
                st.plotly_chart(fig_general, use_container_width=True)
                
                # Mostrar coincidencias
                if modo == "Avanzado":
                    with st.expander(" Ver tabla de coincidencias completa"):
                        st.dataframe(coincidencias_total, use_container_width=True)
                        
                        excel_coinc = crear_excel_descargable({"Coincidencias": coincidencias_total})
                        st.download_button(
                            label=" Descargar coincidencias (Excel)",
                            data=excel_coinc,
                            file_name="coincidencias.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                else:
                    st.markdown("###  Muestra de coincidencias")
                    st.dataframe(coincidencias_total.head(10), use_container_width=True)
                    if len(coincidencias_total) > 10:
                        st.info(f"Mostrando 10 de {len(coincidencias_total)} coincidencias. Cambia a Modo Avanzado para ver todas.")
                
                # ---- ANÁLISIS SOBRE COINCIDENCIAS ----
                if comparar_fechas:
                    coincidencias_total = analizar_fechas_coincidencias(
                        coincidencias_total,
                        modo_avanzado=(modo == "Avanzado"),
                        resultados_dict=resultados_completos,
                    )
                    resultados_completos["Coincidencias"] = coincidencias_total
                
                if usar_openalex:
                    analizar_openalex_coincidencias(
                        coincidencias_total,
                        correo_openalex,
                        modo_avanzado=(modo == "Avanzado"),
                        resultados_dict=resultados_completos,
                    )
                
                # ---- MOSTRAR EXCLUSIVOS (solo en modo avanzado) ----
                if modo == "Avanzado" and total_exclusivos > 0:
                    st.divider()
                    st.subheader(" Registros exclusivos por archivo")
                    st.caption("Registros que solo aparecen en un archivo")
                    
                    for i, (df_excl, nombre) in enumerate(zip(exclusivos_por_archivo, nombres)):
                        if not df_excl.empty:
                            clave = f"Exclusivos_{os.path.splitext(nombre)[0]}"
                            resultados_completos[clave] = df_excl
                            
                            with st.expander(f"**{nombre}** — {len(df_excl)} exclusivos"):
                                st.dataframe(df_excl.head(20), use_container_width=True)
                                
                                excel_excl = crear_excel_descargable({clave: df_excl})
                                st.download_button(
                                    label=f" Descargar exclusivos de {nombre}",
                                    data=excel_excl,
                                    file_name=f"exclusivos_{os.path.splitext(nombre)[0]}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"btn_excl_{i}",
                                )
                
                # ---- DESCARGA COMPLETA ----
                if resultados_completos:
                    st.divider()
                    st.subheader(" Descargar análisis completo")
                    st.caption("Descarga un único archivo Excel con todas las hojas de análisis disponibles.")
                    excel_full = crear_excel_descargable(resultados_completos)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label=" Descargar análisis completo (.xlsx)",
                        data=excel_full,
                        file_name=f"analisis_completo_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                    )
        else:
            st.error("❌ No se encontraron columnas comunes entre los archivos.")
    elif len(archivos) == 1:
        st.info("ℹ️ Sube al menos 2 archivos para realizar comparaciones.")
        if modo == "Avanzado":
            st.info("Puedes usar las opciones de análisis individual en el panel lateral.")
        st.dataframe(dfs[0].head(20), use_container_width=True)
else:
    st.info(" Sube al menos un archivo Excel en el panel lateral para comenzar.")
