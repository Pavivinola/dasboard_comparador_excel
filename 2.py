import streamlit as st
import pandas as pd
import io
import os
import requests
import time
from datetime import datetime
import plotly.express as px
import re

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ======================================================
# FUNCIONES PARA EXCEL DESCARGABLE
# ======================================================
def sanitizar_nombre_hoja(nombre: str) -> str:
    """Ajusta el nombre de la hoja para cumplir con las restricciones de Excel."""
    if not isinstance(nombre, str):
        nombre = str(nombre)

    # Caracteres no permitidos en nombres de hoja de Excel
    for ch in [":", "\\", "/", "?", "*", "[", "]"]:
        nombre = nombre.replace(ch, " ")

    nombre = nombre.strip()
    if not nombre:
        nombre = "Hoja"

    # Límite de 31 caracteres
    return nombre[:31]


def crear_excel_descargable(hojas_dict: dict, incluir_graficos: bool = False) -> bytes:
    """
    Crea un archivo Excel (.xlsx) en memoria con múltiples hojas.
    - hojas_dict: {nombre_logico_hoja: DataFrame}
    - incluir_graficos: si True, agrega una hoja 'Graficos' basada en los datos disponibles.
    """
    output = io.BytesIO()
    if not hojas_dict:
        return output.getvalue()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        sheet_name_map = {}

        # Orden de hojas cuando se pide el Excel "completo"
        keys = list(hojas_dict.keys())
        if incluir_graficos:
            ordered = []

            # 1) Resumen general (si existe)
            if "Resumen_General" in hojas_dict:
                ordered.append("Resumen_General")

            # 2) Exclusivos (ordenados por nombre lógico)
            exclusivos_keys = sorted(k for k in keys if k.startswith("Exclusivos_"))
            ordered.extend(exclusivos_keys)

            # 3) Coincidencias
            if "Coincidencias" in hojas_dict and "Coincidencias" not in ordered:
                ordered.append("Coincidencias")

            # 4) Análisis temporal
            if "Analisis_Referenciales" in hojas_dict:
                ordered.append("Analisis_Referenciales")
            if "Cobertura_Temporal" in hojas_dict:
                ordered.append("Cobertura_Temporal")

            # 5) OpenAlex
            if "OpenAlex_Coincidencias" in hojas_dict:
                ordered.append("OpenAlex_Coincidencias")

            # 6) Cualquier otra hoja que quede fuera
            for k in keys:
                if k not in ordered:
                    ordered.append(k)
        else:
            ordered = keys

        # Escribir hojas de datos
        for logical_name in ordered:
            df = hojas_dict.get(logical_name)
            if df is None or df.empty:
                continue

            sheet_name = sanitizar_nombre_hoja(logical_name)
            sheet_name_map[logical_name] = sheet_name

            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]

            # Auto-ancho de columnas
            if not df.empty:
                for idx, col in enumerate(df.columns):
                    col_values = df[col].astype(str)
                    max_len = max(col_values.map(len).max(), len(str(col))) + 2
                    worksheet.set_column(idx, idx, max_len)

                # Formato condicional simple para columnas numéricas
                numeric_cols = df.select_dtypes(include=["number"]).columns
                for col in numeric_cols:
                    col_idx = df.columns.get_loc(col)
                    # desde fila 1 (segunda fila, ya que la 0 es encabezado) hasta len(df)
                    worksheet.conditional_format(
                        1,
                        col_idx,
                        len(df),
                        col_idx,
                        {"type": "3_color_scale"},
                    )

        # Hoja de gráficos (solo para el Excel "completo")
        if incluir_graficos and "Resumen_General" in hojas_dict:
            graficos_sheet = workbook.add_worksheet(sanitizar_nombre_hoja("Graficos"))

            # ----- Gráfico 1: Resumen general -----
            df_res = hojas_dict["Resumen_General"]
            if (
                df_res is not None
                and not df_res.empty
                and "Métrica" in df_res.columns
                and "Valor" in df_res.columns
            ):
                chart = workbook.add_chart({"type": "column"})
                sheet_res_name = sheet_name_map.get("Resumen_General")
                n_rows = len(df_res)

                # Categorías = Métrica, Valores = Valor
                chart.add_series(
                    {
                        "name": "Resumen general",
                        "categories": [sheet_res_name, 1, 0, n_rows, 0],
                        "values": [sheet_res_name, 1, 1, n_rows, 1],
                    }
                )
                chart.set_title({"name": "Resumen general"})
                chart.set_x_axis({"name": "Métrica"})
                chart.set_y_axis({"name": "Valor"})

                graficos_sheet.insert_chart(1, 1, chart)

            # ----- Gráfico 2: Índice de cobertura (si existe) -----
            if "Cobertura_Temporal" in hojas_dict:
                df_cov = hojas_dict["Cobertura_Temporal"]
                if (
                    df_cov is not None
                    and not df_cov.empty
                    and "Archivo" in df_cov.columns
                    and "Índice Cobertura" in df_cov.columns
                ):
                    chart2 = workbook.add_chart({"type": "column"})
                    sheet_cov = sheet_name_map.get("Cobertura_Temporal")
                    n2 = len(df_cov)
                    idx_arch = df_cov.columns.get_loc("Archivo")
                    idx_ind = df_cov.columns.get_loc("Índice Cobertura")

                    chart2.add_series(
                        {
                            "name": "Índice de cobertura",
                            "categories": [sheet_cov, 1, idx_arch, n2, idx_arch],
                            "values": [sheet_cov, 1, idx_ind, n2, idx_ind],
                        }
                    )
                    chart2.set_title({"name": "Índice de cobertura por archivo"})
                    chart2.set_x_axis({"name": "Archivo"})
                    chart2.set_y_axis({"name": "Índice"})

                    graficos_sheet.insert_chart(16, 1, chart2)

    output.seek(0)
    return output.getvalue()

# ======================================================
# CONFIGURACIÓN DE LA PÁGINA
# ======================================================
st.set_page_config(page_title="Comparador de Excels", layout="wide")
st.title("Compareitor")
st.markdown("<h3 style='text-align: center;'>Fue desarrollado en la Biblioteca de la Universidad Alberto Hurtado 💙</h3>", unsafe_allow_html=True)
st.divider()
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
consultar_solo_uno = False
analizar_tiempo_individual = False
limpiar_duplicados_individual = False

if modo == "Avanzado":
    st.sidebar.subheader("Análisis sobre coincidencias")
    comparar_fechas = st.sidebar.checkbox("Análisis temporal y referenciales", value=False)
    usar_openalex = st.sidebar.checkbox("Consultar OpenAlex (batch)", value=False)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Análisis archivo individual")
    consultar_solo_uno = st.sidebar.checkbox("Consultar OpenAlex para un archivo", value=False)
    analizar_tiempo_individual = st.sidebar.checkbox("Análisis temporal y referencial para un archivo", value=False)
    limpiar_duplicados_individual = st.sidebar.checkbox(
        "Eliminar duplicados de un archivo",
        value=False,
        help="Permite generar una versión sin duplicados de un archivo seleccionado, usando las columnas clave que elijas."
    )
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Opciones avanzadas")
    normalizar_datos = st.sidebar.checkbox("Normalizar ISSN/ISBN automáticamente", value=True)
    mostrar_metricas_detalladas = st.sidebar.checkbox("Mostrar métricas detalladas", value=True)
else:
    # Modo Rápido: valores predeterminados (pero exponemos la casilla de análisis individual si el usuario la quiere)
    comparar_fechas = False
    usar_openalex = False
    consultar_solo_uno = False
    normalizar_datos = True
    mostrar_metricas_detalladas = False
    umbral_similitud = 100
    # permitir que en modo Rápido el usuario active el análisis temporal por archivo
    analizar_tiempo_individual = st.sidebar.checkbox(
        "Análisis temporal y referencial para un archivo",
        value=False
    )
    limpiar_duplicados_individual = st.sidebar.checkbox(
        "Eliminar duplicados de un archivo",
        value=False
    )

# Casilla para limpiar duplicados en la hoja Coincidencias (disponible en todos los modos)
limpiar_duplicados_final = st.sidebar.checkbox(
    "Eliminar duplicados en 'Coincidencias' (por clave)",
    value=False,
    help="Quita filas duplicadas en la hoja Coincidencias usando la clave seleccionada (mantiene la primera aparición por archivo)."
)

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
        st.warning(" No se encontraron ISSN válidos para consultar en OpenAlex.")
        return pd.DataFrame()

    if not correo_openalex or "@" not in correo_openalex:
        st.error(" Por favor ingresa un correo institucional válido para usar la API de OpenAlex.")
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
def analizar_fechas_coincidencias(coincidencias_df, modo_avanzado=False, resultados=None):
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
        st.warning(" No se encontró la columna 'Fecha Inicio'. No se puede detectar referenciales.")
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
        st.warning(" No hay registros con Fecha Inicio válida para calcular cobertura.")
        if resultados is not None:
            resultados["Analisis_Referenciales"] = df_referenciales
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

            # Guardar en resultados completos
            if resultados is not None:
                resultados["Analisis_Referenciales"] = df_referenciales
                resultados["Cobertura_Temporal"] = df_cobertura
    
    return coincidencias_df


# ======================================================
# ANÁLISIS DE COINCIDENCIAS - OPENALEX
# ======================================================
def analizar_openalex_coincidencias(coincidencias_df, correo, modo_avanzado=False, resultados=None):
    """Consulta OpenAlex SOLO para las coincidencias."""
    st.divider()
    st.subheader(" Consulta OpenAlex sobre coincidencias")
    st.caption("Consultando información de las revistas/recursos encontrados en las coincidencias")
    
    # Extraer ISSN de las coincidencias
    issn_list = obtener_issn_de_dataframe(coincidencias_df)
    
    if not issn_list:
        st.warning(" No se encontraron ISSN válidos en las coincidencias para consultar OpenAlex.")
        return
    
    st.info(f"📋 Se encontraron {len(issn_list)} ISSN únicos en las coincidencias")
    
    # Consultar OpenAlex
    df_openalex = consultar_openalex_batch(issn_list, correo)
    
    if df_openalex.empty:
        st.warning(" No se obtuvieron resultados de OpenAlex.")
        return
    
    # Guardar resultados en el diccionario global si corresponde
    if resultados is not None:
        resultados["OpenAlex_Coincidencias"] = df_openalex

    # Mostrar resultados
    st.success(f" Se obtuvieron {len(df_openalex)} resultados de OpenAlex")
    
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
    
    # Descargar resultados como XLSX
    excel_oa = crear_excel_descargable({"OpenAlex_Coincidencias": df_openalex})
    st.download_button(
        label=" Descargar resultados OpenAlex (XLSX)",
        data=excel_oa,
        file_name="openalex_coincidencias.xlsx",
        mime=EXCEL_MIME
    )


# ======================================================
# ANÁLISIS ARCHIVO INDIVIDUAL - OPENALEX
# ======================================================
def analizar_openalex_individual(archivos, nombres, correo, resultados=None):
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
    
    if st.button(" Consultar OpenAlex", type="primary"):
        issn_list = obtener_issn_de_dataframe(df_seleccionado)
        
        if not issn_list:
            st.warning(" No se encontraron ISSN válidos en este archivo.")
            return
        
        st.info(f" Se encontraron {len(issn_list)} ISSN únicos")
        
        df_openalex = consultar_openalex_batch(issn_list, correo)
        
        if not df_openalex.empty:
            st.success(f"✅ Se obtuvieron {len(df_openalex)} resultados")
            st.dataframe(df_openalex, use_container_width=True)

            # Guardar en resultados completos, si corresponde
            if resultados is not None:
                clave = f"OpenAlex_{os.path.splitext(archivo_seleccionado)[0]}"
                resultados[clave] = df_openalex
            
            excel_oa = crear_excel_descargable(
                {f"OpenAlex_{os.path.splitext(archivo_seleccionado)[0]}": df_openalex}
            )
            st.download_button(
                label=" Descargar resultados (XLSX)",
                data=excel_oa,
                file_name=f"openalex_{os.path.splitext(archivo_seleccionado)[0]}.xlsx",
                mime=EXCEL_MIME
            )


# ======================================================
# ANÁLISIS ARCHIVO INDIVIDUAL - FECHAS
# ======================================================
def analizar_fechas_archivo_individual(archivos, nombres):
    """Analiza fechas y referenciales para un archivo individual."""
    st.divider()
    st.subheader(" Análisis temporal y referenciales - Archivo individual")

    archivo_seleccionado = st.selectbox(
        "Selecciona el archivo a analizar:",
        nombres,
        key="select_archivo_fecha_individual"
    )

    idx = nombres.index(archivo_seleccionado)
    df_sel = leer_excel(archivos[idx])

    st.info(f" Archivo seleccionado: **{archivo_seleccionado}** ({len(df_sel)} registros)")

    if st.button(" Ejecutar análisis temporal para este archivo", type="primary"):
        df_proc = procesar_fechas(df_sel.copy())

        if "Fecha Inicio" not in df_proc.columns:
            st.warning(" No se encontró la columna 'Fecha Inicio'. No se puede detectar referenciales en este archivo.")
            st.dataframe(df_proc.head(20), use_container_width=True)
            return

        # Marcar referenciales
        df_proc["Es Referencial"] = ~df_proc["Fecha Inicio"].apply(tiene_fecha_valida)

        total_reg = len(df_proc)
        total_ref = int(df_proc["Es Referencial"].sum())
        con_fecha = total_reg - total_ref

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total registros", total_reg)
        with col2:
            pct_ref = (total_ref / total_reg * 100) if total_reg > 0 else 0.0
            st.metric("Referenciales", f"{total_ref} ({pct_ref:.1f}%)")
        with col3:
            st.metric("Con fechas válidas", con_fecha)

        st.markdown("###  Tabla de registros referenciales")
        st.dataframe(df_proc[df_proc["Es Referencial"]].head(50), use_container_width=True)

        # Cobertura temporal
        registros_temporales = df_proc[df_proc["Es Referencial"] == False].copy()
        if registros_temporales.empty:
            st.warning(" No hay registros con Fecha Inicio válida para calcular cobertura temporal en este archivo.")
            return

        if "Rango Calculado" in registros_temporales.columns:
            registros_temporales["Año Inicio"] = (
                registros_temporales["Rango Calculado"]
                .astype(str)
                .str.extract(r"(\d{4})", expand=False)
                .astype(float)
            )
            registros_temporales["Año Fin"] = (
                registros_temporales["Rango Calculado"]
                .astype(str)
                .str.extract(r"-\s*(\d{4})", expand=False)
                .astype(float)
            )
            registros_temporales["Duración (años)"] = (
                registros_temporales["Año Fin"] - registros_temporales["Año Inicio"]
            )

            registros_temporales = registros_temporales.dropna(subset=["Año Inicio", "Año Fin"])

            if not registros_temporales.empty:
                duracion_prom = registros_temporales["Duración (años)"].mean()
                duracion_min = registros_temporales["Duración (años)"].min()
                duracion_max = registros_temporales["Duración (años)"].max()

                st.markdown("###  Cobertura temporal del archivo")
                st.write(
                    f"Registros analizados: **{len(registros_temporales)}** | "
                    f"Duración promedio: **{duracion_prom:.1f} años** "
                    f"(mín: {duracion_min:.1f}, máx: {duracion_max:.1f})"
                )

                fig = px.histogram(
                    registros_temporales,
                    x="Año Inicio",
                    nbins=20,
                    title="Distribución de años de inicio de cobertura"
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning(" No fue posible extraer años válidos desde 'Rango Calculado' en este archivo.")
        else:
            st.warning(" No se encontró la columna 'Rango Calculado'. Verifica que existan columnas de fechas compatibles.")


# ======================================================
# ANÁLISIS ARCHIVO INDIVIDUAL - ELIMINAR DUPLICADOS
# ======================================================
def eliminar_duplicados_archivo_individual(archivos, nombres, columnas_sugeridas):
    """Permite eliminar duplicados de un archivo individual según columnas seleccionadas."""
    st.divider()
    st.subheader(" Eliminar duplicados - Archivo individual")

    archivo_seleccionado = st.selectbox(
        "Selecciona el archivo a limpiar:",
        nombres,
        key="select_archivo_dup_individual"
    )

    idx = nombres.index(archivo_seleccionado)
    df_sel = leer_excel(archivos[idx])

    st.info(f" Archivo seleccionado: **{archivo_seleccionado}** ({len(df_sel)} registros)")

    columnas_disponibles = df_sel.columns.tolist()

    # Columnas sugeridas: intersección entre sugeridas y disponibles
    columnas_por_defecto = [c for c in columnas_sugeridas if c in columnas_disponibles]
    if not columnas_por_defecto:
        columnas_por_defecto = columnas_disponibles  # fallback

    columnas_clave = st.multiselect(
        "Selecciona las columnas que se usarán para identificar duplicados",
        options=columnas_disponibles,
        default=columnas_por_defecto,
        help="Las filas que tengan el mismo valor en TODAS estas columnas se considerarán duplicadas."
    )

    if columnas_clave and st.button("Eliminar duplicados de este archivo", type="primary"):
        df_sin_duplicados = df_sel.drop_duplicates(subset=columnas_clave, keep="first")

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Filas originales", len(df_sel))
        with col2:
            st.metric("Filas después de limpiar", len(df_sin_duplicados))

        st.markdown("### Vista previa del archivo sin duplicados")
        st.dataframe(df_sin_duplicados.head(50), use_container_width=True)

        # Preparar Excel descargable
        nombre_base = os.path.splitext(archivo_seleccionado)[0]
        excel_sin_dup = crear_excel_descargable(
            {f"{nombre_base}_sin_duplicados": df_sin_duplicados}
        )
        st.download_button(
            label=" Descargar archivo sin duplicados (XLSX)",
            data=excel_sin_dup,
            file_name=f"{nombre_base}_sin_duplicados.xlsx",
            mime=EXCEL_MIME,
            key="btn_descargar_dup_individual"
        )


# ======================================================
# PROCESO PRINCIPAL
# ======================================================
if archivos:
    dfs = [leer_excel(a) for a in archivos]
    nombres = [a.name for a in archivos]

    # Diccionario global para el análisis completo
    resultados_completos = {}
    
    # ---- ANÁLISIS INDIVIDUAL (OpenAlex) ----
    if consultar_solo_uno and len(archivos) > 0:
        analizar_openalex_individual(archivos, nombres, correo_openalex, resultados_completos)

    # ---- ANÁLISIS INDIVIDUAL (Fechas) ----
    if analizar_tiempo_individual and len(archivos) > 0:
        analizar_fechas_archivo_individual(archivos, nombres)

    # ---- ANÁLISIS INDIVIDUAL (Eliminar duplicados) ----
    if limpiar_duplicados_individual and len(archivos) > 0:
        # Si hay varios archivos, usamos columnas comunes como sugerencia;
        # si hay solo uno, usamos todas sus columnas.
        if len(dfs) > 1:
            columnas_comunes = list(set.intersection(*(set(df.columns) for df in dfs)))
        else:
            columnas_comunes = dfs[0].columns.tolist()
        eliminar_duplicados_archivo_individual(archivos, nombres, columnas_comunes)
    
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
                "Columnas": [df.shape[1] for df in dfs]
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

                # Si el usuario pidió limpiar duplicados, deduplicar por (Archivo, __clave__)
                if limpiar_duplicados_final and "__clave__" in coincidencias_total.columns:
                    coincidencias_total = coincidencias_total.drop_duplicates(
                        subset=["Archivo", "__clave__"],
                        keep="first"
                    )

                # luego quitar la columna interna de clave
                coincidencias_total = coincidencias_total.drop(columns=["__clave__"], errors="ignore")
                
                total_exclusivos = sum(len(df) for df in exclusivos_por_archivo)
                total_registros = sum(len(df) for df in dfs)

                # Guardar coincidencias en resultados completos
                resultados_completos["Coincidencias"] = coincidencias_total
                
                # ---- RESUMEN GENERAL ----
                st.divider()
                st.subheader("Resumen general")
                
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

                # DataFrame de resumen para el Excel completo
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
                        "Cantidad": [len(coincidencias_total), total_exclusivos]
                    }),
                    names="Tipo",
                    values="Cantidad",
                    title="Distribución general de registros",
                    color="Tipo",
                    color_discrete_map={"Coincidencias": "#2ECC71", "Exclusivos": "#3498DB"}
                )
                fig_general.update_traces(textinfo="percent+value")
                st.plotly_chart(fig_general, use_container_width=True)
                
                # Mostrar coincidencias
                if modo == "Avanzado":
                    with st.expander(" Ver tabla de coincidencias completa"):
                        st.dataframe(coincidencias_total, use_container_width=True)
                        
                        # Descargar coincidencias como XLSX
                        excel_coinc = crear_excel_descargable({"Coincidencias": coincidencias_total})
                        st.download_button(
                            label=" Descargar coincidencias (XLSX)",
                            data=excel_coinc,
                            file_name="coincidencias.xlsx",
                            mime=EXCEL_MIME
                        )
                else:
                    # Modo rápido: solo primeras 10 filas
                    st.markdown("###  Muestra de coincidencias")
                    st.dataframe(coincidencias_total.head(10), use_container_width=True)
                    if len(coincidencias_total) > 10:
                        st.info(f"Mostrando 10 de {len(coincidencias_total)} coincidencias. Cambia a Modo Avanzado para ver todas.")
                
                # ---- ANÁLISIS SOBRE COINCIDENCIAS ----
                if comparar_fechas:
                    coincidencias_total = analizar_fechas_coincidencias(
                        coincidencias_total, 
                        modo_avanzado=(modo == "Avanzado"),
                        resultados=resultados_completos
                    )
                    # Actualizar la versión almacenada
                    resultados_completos["Coincidencias"] = coincidencias_total
                
                if usar_openalex:
                    analizar_openalex_coincidencias(
                        coincidencias_total, 
                        correo_openalex,
                        modo_avanzado=(modo == "Avanzado"),
                        resultados=resultados_completos
                    )
                
                # ---- MOSTRAR EXCLUSIVOS (solo en modo avanzado) ----
                if modo == "Avanzado" and total_exclusivos > 0:
                    st.divider()
                    st.subheader(" Registros exclusivos por archivo")
                    st.caption("Registros que solo aparecen en un archivo")
                    
                    for i, (df_excl, nombre) in enumerate(zip(exclusivos_por_archivo, nombres)):
                        if not df_excl.empty:
                            # Guardar en resultados completos
                            clave = f"Exclusivos_{os.path.splitext(nombre)[0]}"
                            resultados_completos[clave] = df_excl

                            with st.expander(f"**{nombre}** — {len(df_excl)} exclusivos"):
                                st.dataframe(df_excl.head(20), use_container_width=True)
                                
                                excel_excl = crear_excel_descargable(
                                    {clave: df_excl}
                                )
                                st.download_button(
                                    label=f" Descargar exclusivos de {nombre} (XLSX)",
                                    data=excel_excl,
                                    file_name=f"exclusivos_{os.path.splitext(nombre)[0]}.xlsx",
                                    mime=EXCEL_MIME,
                                    key=f"btn_excl_{i}"
                                )

                # ---- BOTÓN DE DESCARGA COMPLETA ----
                if resultados_completos:
                    st.divider()
                    st.subheader(" Descargar análisis completo")
                    st.caption("Descarga un único archivo Excel con todas las hojas de análisis disponibles.")
                    excel_full = crear_excel_descargable(resultados_completos, incluir_graficos=True)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label=" Descargar análisis completo (.xlsx)",
                        data=excel_full,
                        file_name=f"analisis_completo_{timestamp}.xlsx",
                        mime=EXCEL_MIME,
                        type="primary"
                    )

        else:
            st.error("❌ No se encontraron columnas comunes entre los archivos.")
    
    elif len(archivos) == 1:
        st.info("ℹ️ Sube al menos 2 archivos para realizar comparaciones.")
        if modo == "Avanzado":
            st.info(" Puedes usar la opción de consulta individual de OpenAlex en el panel lateral.")
        st.dataframe(dfs[0].head(20))

else:
    st.info(" Sube al menos un archivo Excel en el panel lateral para comenzar.")
    
    # Mostrar guía según el modo
    if modo == "Rápido":
        st.markdown("""
        ###  Modo Rápido - Guía de uso
        
        1. **Sube 2 o más archivos Excel** en el panel lateral
        2. **Selecciona las columnas clave** para comparar (ej: ISSN, ISBN, Título)
        3. **Resultados rápidos** con menos visualizaciones 
        
        **Pensado para:** Comparaciones rápidas y análisis básicos
        """)
    else:
        st.markdown("""
        ###  Modo Avanzado - Guía de uso
        
        1. **Sube 2 o más archivos Excel** en el panel lateral
        2. **Selecciona las columnas clave para comparar** (ej: ISSN, ISBN, Título)
        3. **Activa las opciones avanzadas** que necesites:
           -  **Análisis temporal y referenciales**: Detecta recursos sin fechas y analiza cobertura
           -  **OpenAlex en lote**: Consulta información de acceso abierto de revistas de las coincidencias
           -  **OpenAlex individual**: Analiza un archivo específico
           -  **Análisis temporal individual**: Analiza fechas y referenciales de un archivo específico
        4. **Explora visualizaciones detalladas** y descarga todos los resultados en **Excel (.xlsx)**
        
        **Pensado para:** Análisis más completos y detallados.
        
        ---
        
        ####  Opciones disponibles:
        -  Normalización de ISSN/ISBN
        -  Estadísticas detalladas (min, max, promedios)
        -  Visualizaciones adicionales (top 10, distribuciones por país)
        -  Descarga de todos los resultados en Excel
        -  Vista completa de exclusivos por archivo
        """)
    
    # Tips generales
    st.divider()
    st.markdown("""
    ###  Consejos para mejores resultados:
    
    - **Columnas clave**: Usa identificadores únicos como ISSN, ISBN, DOI,Título o Autor
    - **Normalización**: Actívala para ignorar diferencias de formato en ISSN/ISBN
    - **OpenAlex**: Requiere un correo institucional válido para mejores resultados
    - **Fechas**: Las columnas deben llamarse exactamente "Fecha Inicio", "Fecha Termino" y "Retraso"
    """)
