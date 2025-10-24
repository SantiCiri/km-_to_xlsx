import xml.etree.ElementTree as ET
import zipfile
import tempfile
import os
from glob import glob
from shapely.geometry import Polygon
import pandas as pd
import streamlit as st
import pandas as pd
from io import BytesIO

def parse_kml(kml_content):
    """
    Parse un contenido XML de KML y extrae las coordenadas.
    Devuelve una lista de tuplas (lon, lat, alt).
    """
    coords = []
    root = ET.fromstring(kml_content)

    # KML tiene namespaces, los manejamos
    ns = {"kml": "http://www.opengis.net/kml/2.2"}

    for elem in root.findall(".//kml:coordinates", ns):
        text = elem.text.strip()
        for line in text.split():
            parts = line.split(",")
            lon, lat = float(parts[0]), float(parts[1])
            coords.append((lon, lat))
    return coords


def read_kml_file(filepath):
    """Lee un archivo KML y devuelve lista de coordenadas"""
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()
    return parse_kml(content)


def read_kmz_file(filepath):
    """Lee un archivo KMZ, extrae el KML y devuelve lista de coordenadas"""
    coords = []
    with zipfile.ZipFile(filepath, "r") as z:
        for name in z.namelist():
            if name.endswith(".kml"):
                with z.open(name) as f:
                    content = f.read().decode("utf-8")
                    coords.extend(parse_kml(content))
    return coords


def read_file(file):
    """Detecta si es KML o KMZ y devuelve coordenadas"""
    # Si es un UploadedFile, lo tratamos distinto
    if not isinstance(file, (str, bytes, os.PathLike)):
        filename = file.name
        ext = os.path.splitext(filename)[1].lower()
        data = file.read()  # leer contenido binario
        # Crear archivo temporal en disco (solo mientras se procesa)
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        try:
            if ext == ".kml":
                return read_kml_file(tmp_path)
            elif ext == ".kmz":
                return read_kmz_file(tmp_path)
            else:
                raise ValueError("El archivo debe ser .kml o .kmz")
        finally:
            os.remove(tmp_path)


def calcular_superficie_hectareas(coords):
    """
    Calcula la superficie de un polígono definido por coordenadas (lat, lon).
    Devuelve el área en hectáreas.
    
    Parámetros:
        coords (list of tuple): lista de (lat, lon)
    
    Retorno:
        float: superficie en hectáreas
    """
    if coords[0] != coords[-1]:
        # cerrar el polígono si no está cerrado
        coords.append(coords[0])
    
    # Shapely espera (lon, lat)
    poly = Polygon([(lon, lat) for lat, lon in coords])

    area_ha=round(poly.area*1000000,1)
    x=poly.centroid.x
    y=poly.centroid.y

    centroid=f"{y}, {x}"
    print(centroid)
    return area_ha,centroid

st.set_page_config(page_title="Conversor KMZ/KML a Excel", page_icon="📍",layout="wide")

st.title("📍 Conversor de archivos KMZ / KML a Excel")

st.write("Subí uno o varios archivos KMZ o KML para convertirlos en una tabla con coordenadas, centroides y superficie.")
st.write("No guardamos tus archivos en memoria, por lo que tu información está segura y no podremos recuperar nada una vez que cierres la sesión.")
col1, col2 = st.columns([1, 1])  # podés ajustar la proporción
with col1:
    st.write("Por ayuda, opiniones y transformadores de cartas de porte a Excel, comunicate con Santiago Cirigliano al 11-4048-6131 o click en:")
with col2:
    st.link_button("💬 WhatsApp", "https://wa.me/541140486131")

# Cargar múltiples archivos
uploaded_files = st.file_uploader(
    "Seleccioná tus archivos kmz o kml. Puedes subir varios a la vez",
    type=["kmz", "kml"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} archivo(s) cargado(s) correctamente.")
else:
    st.info("Esperando archivos...")

# Botón para ejecutar el proceso
if st.button("🚀 Transformar"):
    if not uploaded_files:
        st.warning("Por favor, subí al menos un archivo antes de transformar.")
    else:
        st.write("Este es el formato de Excel requerido por Visec para cargar tus polígonos. Cada polígono debe ingresarse una única vez en la plataforma, dentro del proceso denominado “Registro de UPs”. Además, deberás completar el resto de la planilla con información que no está contenida en los archivos KMZ que nos enviaste, a fin de completar correctamente el archivo denominado Template-UP-ORIGINAL-Sistema-VISEC-MRV.")
        df = pd.DataFrame(columns=["Archivo", "Polígono", "Punto Referencia", "Superficie"])
        errores = []
        
        # Procesar cada archivo
        for file in uploaded_files:
            filename = file.name
            if not (filename.endswith(".kmz") or filename.endswith(".kml")):
                errores.append(filename)
                continue
            
            try:
                coords = read_file(file)
                sup, centroid = calcular_superficie_hectareas(coords)
                df.loc[len(df)] = [filename, coords, centroid, sup]
            except Exception as e:
                errores.append(f"{filename} (error: {e})")

        if not df.empty:
            # Mostrar preview
            st.dataframe(df)

            # Convertir a Excel en memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Datos")
            output.seek(0)

            st.download_button(
                label="📥 Descargar Excel",
                data=output,
                file_name="resultados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            if st.download_button==True:st.balloons()
        else:
            st.warning("No se generaron datos válidos.")

        # Mostrar archivos con errores
        if errores:
            st.error(f"⚠️ Archivos ignorados o con error:\n" + "\n".join(errores))

