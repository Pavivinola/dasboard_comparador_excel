#  Comparador de Excels — Dashboard para análisis de duplicados y coincidencias

**Desarrollado por:** Jorge Moreno  
**Institución:** Biblioteca Patrimonial San Ignacio — Universidad Alberto Hurtado  
**Proyecto:** Herramienta de análisis comparativo y metadatos con integración OpenAlex  
**Framework:** Streamlit  

---

##  Descripción general

**Comparador de Excels** es un dashboard interactivo desarrollado en **Python + Streamlit**  
para comparar múltiples archivos Excel (.xlsx), detectar coincidencias, identificar registros exclusivos  
y consultar metadatos de revistas académicas mediante la API de **OpenAlex**.

El sistema fue diseñado para apoyar procesos bibliográficos, adquisiciones y control de colecciones en entornos académicos o patrimoniales,  
ofreciendo una interfaz visual simple y resultados descargables en Excel.

---

##  Funcionalidades principales

 Carga de múltiples archivos Excel (.xlsx)  
 Selección flexible de columnas clave (ISSN, E-ISSN, Título, etc.)  
 Comparación avanzada con normalización de formatos  
 Modo **Rápido** y **Avanzado**  
 Consulta a **OpenAlex** en lotes (batch)  
 Descarga de resultados consolidados en Excel  
 Visualización de estadísticas y gráficos interactivos (Plotly)  
 Compatible con archivos grandes y sin instalación local (versión web)  

---

##  Estructura del proyecto

compareitor_dashboard/
│
├── dashboard_comparar_excel.py # Código principal del dashboard
├── requirements.txt # Dependencias
├── README.md # Este documento
├── data/ # (Opcional) Archivos de entrada/salida
└── assets/ # (Opcional) Recursos visuales

##  Ejecución local

Si deseas ejecutar la aplicación en tu equipo local:

1. Clona este repositorio:
   
   git clone https://github.com/Pavivinola/compareitor-dashboard.git
   cd compareitor-dashboard

2. Crea y activa un entorno virtual:

  python -m venv venv
  venv\Scripts\activate

3. Instala las dependencias:

  pip install -r requirements.txt

4. Ejecuta la aplicación:

  streamlit run dashboard_comparar_excel.py

5.Abre en tu navegador:

  http://localhost:8501


