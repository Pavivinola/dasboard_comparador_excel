import streamlit as st
import pandas as pd
import io
import os
import requests
import time
from datetime import datetime
import xlsxwriter
import plotly.express as px

# ======================================================
# CONFIGURACIÓN DE LA APLICACIÓN
# ======================================================
st.set_page_config(page_title="Comparador de Excels", layout="wide") # Configuración de la página, título y diseño
st.title("Dashboard Comparador de Excels") # Título principal de la aplicación
st.markdown( # Descripción breve de la aplicación
    "Esta herramienta permite comparar varios archivos Excel, detectar coincidencias y registros exclusivos, "
    "y consultar metadatos de revistas en OpenAlex."
)
st.divider() # Línea divisoria

# ======================================================
# PANEL LATERAL
# ======================================================
st.sidebar.header("Configuración") # Encabezado del panel lateral

modo = st.sidebar.radio("Selecciona el modo de ejecución:", ["Rápido", "Avanzado"]) # Modo de ejecución
usar_openalex = st.sidebar.checkbox("Consultar información en OpenAlex (batch)", value=False) # Opción para consultar OpenAlex
consultar_solo_uno = st.sidebar.checkbox("Consultar OpenAlex para un solo archivo", value=False) # Opción para consultar OpenAlex solo para un archivo

correo_openalex = st.sidebar.text_input( # Campo para ingresar correo institucional
    "Correo para identificarte ante OpenAlex (recomendado)",
    placeholder="tucorreo@institucion.cl" # Placeholder del campo
)

archivos = st.sidebar.file_uploader( # Esta variable permite subir archivos Excel
    "Sube uno o más archivos Excel (.xlsx)",
    type="xlsx", # Tipo de archivo permitido
    accept_multiple_files=True # Permitir subir múltiples archivos
)

# ======================================================
# FUNCIONES AUXILIARES
# ======================================================
@st.cache_data # Esto es para cachear los datos y evitar recargas innecesarias
def leer_excel(archivo): # Función para leer un archivo Excel y manejar errores
    try:  # Intentar leer el archivo Excel
        return pd.read_excel(archivo) # Si lo lee, devuelve el DataFrame
    except Exception as e: # Si hay un error, lo muestra en pantalla
        st.error(f"Error al leer {archivo.name}: {e}") # Muestra el error
        return pd.DataFrame() # Devuelve un DataFrame vacío en caso de error

def normalizar_valor(valor): # Esta función normaliza valores como ISSN, ISBN, etc.
    """Normaliza ISSN, ISBN, EISSN, etc."""
    if pd.isna(valor): # Si el valor es NaN, devuelve cadena vacía
        return "" # Devuelve cadena vacía para valores NaN
    valor = str(valor).strip().upper() # Convierte a cadena, elimina espacios y pone en mayúsculas
    valor = valor.replace("-", "").replace(" ", "").replace(".", "") # Elimina guiones, espacios y puntos
    if valor.isdigit() and len(valor) == 8: # Si es un número de 8 dígitos, formatea como ISSN 
        return valor # Retorna el valor normalizado
    return valor # Retorna el valor normalizado

def generar_clave_prioritaria(row, columnas, normalizar=False):  # Función para generar clave prioritaria
    """Devuelve la primera columna con valor válido, con o sin normalización."""
    for col in columnas: # Recorre las columnas especificadas
        valor = row[col] # Obtiene el valor de la columna actual
        if normalizar: # Si se requiere normalización
            valor = normalizar_valor(valor) # Normaliza el valor
        if valor and str(valor).lower() != "nan": # Si el valor es válido
            return valor # Retorna el valor válido
    return None # Si no hay valores válidos, retorna None
@st.cache_data # Cachea los datos para evitar recargas innecesarias
def consultar_openalex_batch(lista_issn, correo_openalex=None): # Esta función consulta OpenAlex en lotes
    """
    Consulta OpenAlex en lotes de 50 ISSN.
    Incluye modo debug para ver las URLs generadas y las respuestas parciales.
    """
    resultados = [] # Lista para almacenar los resultados
    base_url = "https://api.openalex.org/sources?filter=issn:" # URL base de la API de OpenAlex
    batch_size = 50 # Tamaño del lote de consultas

    for i in range(0, len(lista_issn), batch_size): # Recorre la lista de ISSN en lotes
        lote = lista_issn[i:i + batch_size] # Obtiene el lote actual
        url = base_url + "|".join(lote) # Construye la URL de consulta para el lote
        if correo_openalex: # Si se porporciona correo, lo añade a la URL
            url += f"&mailto={correo_openalex}" # Añade el correo a la URL

        # === DEBUG: Mostrar información en consola ===
        print("\n🔍 Lote consultado:", lote) # Muestra en consola
        print("🌐 URL enviada:", url) # 

        try: # Intenta realizar la consulta
            r = requests.get(url) # Realiza la solicitud GET a la API
            print(" Código HTTP:", r.status_code) # Muestra el código de estado HTTP

            if r.status_code == 200: # Is la respuesta es exitosa (200)
                data = r.json() # Parsea la respuesta JSON

                # Mostrar parte del JSON para inspección (solo en el primer lote)
                if i == 0:
                    st.markdown("### 🧩 Respuesta de OpenAlex (vista parcial)")
                    st.code(str(data)[:800], language="json")

                resultados_lote = data.get("results", [])
                print("📦 Resultados recibidos:", len(resultados_lote))

                for item in resultados_lote:
                    resultados.append({
                        "ISSN": item.get("issn_l"),
                        "Nombre revista": item.get("display_name", ""),
                        "País": item.get("country_code", ""),
                        "Tipo": item.get("type", ""),
                        "Acceso abierto": "Sí" if item.get("is_oa") else "No",
                        "Última actualización": item.get("updated_date", "")
                    })
            else:
                print(f"⚠️ Error {r.status_code} en la consulta. Texto: {r.text[:200]}")
                st.warning(f"Error {r.status_code} al consultar OpenAlex.")
            
            time.sleep(0.3)  # pequeña pausa entre lotes

        except Exception as e:
            print("❌ Error de conexión:", e)
            st.error(f"Error consultando OpenAlex: {e}")

    print(f"\n✅ Total general de resultados recibidos: {len(resultados)}")
    return pd.DataFrame(resultados)

# @st.cache_data
# def consultar_openalex_batch(lista_issn, correo_openalex=None):
#     """Consulta OpenAlex en lotes de 50 ISSN con opción de incluir correo institucional."""
#     resultados = []
#     base_url = "https://api.openalex.org/sources?filter=issn:"
#     batch_size = 50

#     for i in range(0, len(lista_issn), batch_size):
#         lote = lista_issn[i:i + batch_size]
#         url = base_url + "|".join(lote)
#         if correo_openalex:
#             url += f"&mailto={correo_openalex}"

#         try:
#             r = requests.get(url)
#             if r.status_code == 200:
#                 data = r.json()
#                 for item in data.get("results", []):
#                     resultados.append({
#                         "ISSN": item.get("issn_l"),
#                         "Nombre revista": item.get("display_name", ""),
#                         "País": item.get("country_code", ""),
#                         "Tipo": item.get("type", ""),
#                         "Acceso abierto": "Sí" if item.get("is_oa") else "No",
#                         "Última actualización": item.get("updated_date", "")
#                     })
#             else:
#                 st.warning(f"Error {r.status_code} en la consulta a OpenAlex.")
#             time.sleep(0.3)
#         except Exception as e:
#             st.error(f"Error consultando OpenAlex: {e}")

#     return pd.DataFrame(resultados)

# ======================================================
# PROCESO PRINCIPAL
# ======================================================
if archivos:
    dfs = [leer_excel(a) for a in archivos]
    nombres = [a.name for a in archivos]

    # ------------------------------------------------------
    # CASO: CONSULTAR SOLO UN ARCHIVO CON OPENALEX
    # ------------------------------------------------------
    if len(archivos) == 1 and consultar_solo_uno:
        st.subheader("Vista previa del archivo")
        df = dfs[0]
        filas, columnas = df.shape
        st.markdown(f"**{nombres[0]}** — {filas} filas × {columnas} columnas")
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

    # ------------------------------------------------------
    # CASO: COMPARACIÓN ENTRE VARIOS ARCHIVOS
    # ------------------------------------------------------
    if len(archivos) > 1:
        # Vista previa
        st.subheader("Vista previa de los archivos cargados")
        for nombre, df in zip(nombres, dfs):
            filas, columnas = df.shape
            st.markdown(f"**{nombre}** — {filas} filas × {columnas} columnas")
            st.dataframe(df.head(10))
            st.markdown("---")
        st.divider()

        # Selección de columnas clave
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

                # Unir claves y comparar
                claves = pd.concat([df[["__clave__"]] for df in dfs], keys=range(len(dfs)))
                claves["Archivo"] = claves.index.get_level_values(0)
                conteo = claves.groupby("__clave__")["Archivo"].nunique()

                # Coincidencias
                claves_comunes = conteo[conteo > 1].index
                coincidencias_total = pd.concat([
                    df[df["__clave__"].isin(claves_comunes)] for df in dfs
                ]).drop(columns=["__clave__"])

                # Exclusivos
                exclusivos_por_archivo = [
                    df[df["__clave__"].isin(conteo[conteo == 1].index)].drop(columns=["__clave__"])
                    for df in dfs
                ]

                # Resultados básicos
                total_exclusivos = sum(len(df) for df in exclusivos_por_archivo)
                st.divider()
                st.subheader("Resumen general")
                c1, c2, c3 = st.columns(3)
                c1.metric("Archivos cargados", len(archivos))
                c2.metric("Coincidencias encontradas", len(coincidencias_total))
                c3.metric("Registros exclusivos", total_exclusivos)

                # Gráficos
                st.markdown("### Visualización de resultados")
                fig1 = px.pie(
                    pd.DataFrame({
                        "Tipo": ["Coincidencias", "Exclusivos"],
                        "Cantidad": [len(coincidencias_total), total_exclusivos]
                    }),
                    names="Tipo", values="Cantidad",
                    title="Coincidencias vs Exclusivos",
                    color="Tipo",
                    color_discrete_map={"Coincidencias": "#2ECC71", "Exclusivos": "#3498DB"}
                )
                fig1.update_traces(textinfo="percent+value")
                st.plotly_chart(fig1, use_container_width=True)

                if len(archivos) > 1:
                    fig2 = px.pie(
                        pd.DataFrame({
                            "Archivo": nombres,
                            "Exclusivos": [len(df) for df in exclusivos_por_archivo],
                        }),
                        names="Archivo", values="Exclusivos",
                        title="Distribución de Exclusivos por Archivo",
                    )
                    fig2.update_traces(textinfo="percent+value")
                    st.plotly_chart(fig2, use_container_width=True)

                # Consultar OpenAlex (solo coincidencias)
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

                # Generar Excel
                fecha = datetime.now().strftime("%Y-%m-%d_%H-%M")
                nombre_salida = f"resultado_comparacion_{fecha}.xlsx"
                ruta_salida = os.path.join(os.getcwd(), nombre_salida)
                output = io.BytesIO()

                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    workbook = writer.book
                    formato_titulo = workbook.add_format({
                        "bold": True, "font_color": "white",
                        "bg_color": "#004D40", "border": 1
                    })
                    formato_general = workbook.add_format({"border": 1})
                    formato_texto = workbook.add_format({"text_wrap": True, "border": 1})

                    resumen = pd.DataFrame({
                        "Parámetro": ["Fecha", "Modo", "Archivos", "Columnas clave", "Coincidencias", "Exclusivos totales"],
                        "Valor": [datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                  modo, ", ".join(nombres),
                                  ", ".join(columnas_clave),
                                  len(coincidencias_total), total_exclusivos]
                    })
                    resumen.to_excel(writer, sheet_name="Resumen", index=False)
                    hoja_resumen = writer.sheets["Resumen"]
                    hoja_resumen.set_column("A:A", 35, formato_general)
                    hoja_resumen.set_column("B:B", 60, formato_general)
                    hoja_resumen.write_row("A1", ["Parámetro", "Valor"], formato_titulo)

                    coincidencias_total.to_excel(writer, sheet_name="Coincidencias", index=False)
                    for i, exclusivos in enumerate(exclusivos_por_archivo):
    # Limpiar el nombre del archivo (sin extensión ni caracteres prohibidos)
                        nombre_limpio = os.path.splitext(nombres[i])[0]  # quita .xlsx
                        nombre_limpio = "".join(c for c in nombre_limpio if c.isalnum() or c in (" ", "_", "-"))
                        nombre_hoja = f"Exclusivos_{nombre_limpio}"[:31]  # máximo 31 caracteres

                        exclusivos.to_excel(writer, sheet_name=nombre_hoja, index=False)

                    # for i, exclusivos in enumerate(exclusivos_por_archivo):
                    #     exclusivos.to_excel(writer, sheet_name=f"Exclusivos_{nombres[i][:25]}", index=False)

                    if not df_openalex.empty:
                        df_openalex.to_excel(writer, sheet_name="OpenAlex_Resultados", index=False)

                with open(ruta_salida, "wb") as f:
                    f.write(output.getvalue())
                output.seek(0)
                st.download_button(
                    "Descargar archivo Excel con resultados",
                    data=output,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Selecciona al menos una columna clave para realizar la comparación.")
else:
    st.info("Sube al menos un archivo Excel para comenzar la comparación.")
