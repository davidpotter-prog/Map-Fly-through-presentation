import streamlit as st
import os
import json
import requests
import geopandas as gpd
from shapely.geometry import LineString
import numpy as np
from pptx import Presentation
from pptx.util import Inches
from google import genai
from google.genai import types
from pydantic import BaseModel
import fiona
import tempfile
import zipfile
import pandas as pd

# Enable KML/KMZ support in GeoPandas
fiona.drvsupport.supported_drivers['KML'] = 'rw'

# Gemini Data Structure
class SlideContent(BaseModel):
    title: str
    ai_summary: str

# --- CORE FUNCTIONS ---
def sample_polyline_waypoints(gdf, num_stops):
    line = gdf.geometry.iloc[0]
    total_length = line.length
    distances = np.linspace(0, total_length, num_stops)
    stops = []
    for d in distances:
        pt = line.interpolate(d)
        lookahead_pt = line.interpolate(min(d + (total_length * 0.01), total_length))
        heading = np.degrees(np.arctan2(lookahead_pt.x - pt.x, lookahead_pt.y - pt.y)) % 360
        stops.append({"geometry": pt, "distance_km": round(d / 1000, 2), "heading": heading})
    return stops

def fetch_waypoint_maps(stops_data, api_key, map_type):
    os.makedirs("temp_maps", exist_ok=True)
    base_url = "https://maps.googleapis.com/maps/api/staticmap"
    for idx, stop in enumerate(stops_data):
        lat, lon = stop["geometry"].y, stop["geometry"].x  
        params = {
            "center": f"{lat},{lon}",
            "zoom": 15,
            "size": "800x450",
            "maptype": map_type.lower(),
            "markers": f"color:red|label:{idx+1}|{lat},{lon}",
            "key": api_key
        }
        response = requests.get(base_url, params=params)
        if response.status_code == 200:
            file_path = f"temp_maps/stop_{idx+1}.png"
            with open(file_path, "wb") as f: f.write(response.content)
            stop["map_image_path"] = file_path
        else:
            stop["map_image_path"] = None
    return stops_data

def generate_waypoints_narrative(stops_data, api_key):
    client = genai.Client(api_key=api_key)
    for idx, stop in enumerate(stops_data):
        prompt = f"We are at waypoint {idx + 1}. Distance: {stop['distance_km']}km, Heading: {stop['heading']}°. Generate a concise, engaging slide title and a brief 2-3 sentence narrative summary for this stop."
        try:
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=SlideContent, temperature=0.7)
            )
            data = json.loads(response.text)
            stop["title"] = data.get("title", f"Waypoint {idx + 1}")
            stop["ai_summary"] = data.get("ai_summary", "")
        except Exception as e:
            stop["title"] = f"Waypoint {idx + 1}"
            stop["ai_summary"] = f"Metrics: {stop['distance_km']} km, {stop['heading']}° heading."
    return stops_data

def create_google_earth_kml(stops_data, output_path="earth_presentation.kml"):
    """
    Generates a KML file with sequential placemarks, 3D camera angles, 
    and HTML-formatted popup bubbles containing the AI narrative.
    """
    kml = ['<?xml version="1.0" encoding="UTF-8"?>']
    kml.append('<kml xmlns="http://www.opengis.net/kml/2.2">')
    kml.append('  <Document>')
    kml.append('    <name>AI Generated Earth Tour</name>')

    for idx, stop in enumerate(stops_data):
        lon, lat = stop["geometry"].x, stop["geometry"].y
        heading = stop["heading"]
        title = stop.get("title", f"Stop {idx+1}")
        summary = stop.get("ai_summary", "")

        # Format the popup bubble text using HTML
        description = f"<![CDATA[<h3>{title}</h3><p>{summary}</p>]]>"

        kml.append('    <Placemark>')
        kml.append(f'      <name>{idx + 1}. {title}</name>')
        kml.append(f'      <description>{description}</description>')
        
        # The Camera View: controls what the user sees when they click "Next"
        kml.append('      <LookAt>')
        kml.append(f'        <longitude>{lon}</longitude>')
        kml.append(f'        <latitude>{lat}</latitude>')
        kml.append('        <altitude>0</altitude>')
        kml.append(f'        <heading>{heading}</heading>') # Looks down the path
        kml.append('        <tilt>65</tilt>') # Gives a cinematic 3D angled view
        kml.append('        <range>800</range>') # Camera distance from the ground (meters)
        kml.append('        <altitudeMode>relativeToGround</altitudeMode>')
        kml.append('      </LookAt>')
        
        kml.append('      <Point>')
        kml.append(f'        <coordinates>{lon},{lat},0</coordinates>')
        kml.append('      </Point>')
        kml.append('    </Placemark>')

    kml.append('  </Document>')
    kml.append('</kml>')

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(kml))
        
    return output_path

# --- UI LAYOUT ---
st.set_page_config(page_title="Route-to-Slide Generator", layout="wide")
st.title("🗺️ Automated Geospatial Slideshow Builder")

# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Settings")
    
    # Try to load secrets first, otherwise leave blank for manual entry
    maps_key = st.text_input("Google Maps API Key", value=st.secrets.get("GOOGLE_MAPS_KEY", ""), type="password")
    gemini_key = st.text_input("Gemini API Key", value=st.secrets.get("GEMINI_API_KEY", ""), type="password")
    
    st.markdown("---")
    slide_count = st.slider("Number of Waypoints / Slides", min_value=3, max_value=20, value=8)
    map_type = st.selectbox("Map Style", ["Satellite", "Terrain", "Hybrid", "Roadmap"])
    
    st.markdown("---")
    st.subheader("🎨 Custom Branding (Optional)")
    uploaded_template = st.file_uploader("Upload Corporate Template (.pptx)", type=["pptx"])
    layout_index = st.number_input("Template Layout Index (Default: 6 for Blank)", min_value=0, max_value=15, value=6)

# Main Dashboard
uploaded_route = st.file_uploader("📂 Upload Route File (KMZ, KML, GeoJSON, GPX)", type=["kmz", "kml", "geojson", "json", "gpx"])

if uploaded_route:
    st.success(f"Loaded {uploaded_route.name} successfully!")
    
    if st.button("🚀 Generate Presentation", type="primary"):
        if not maps_key or not gemini_key:
            st.error("⚠️ Please provide both API keys in the sidebar.")
        else:
            with st.spinner("Analyzing geometry, capturing maps, and generating narrative..."):
                # 1. Save uploaded file to a temporary location (required for GeoPandas to read zipped KMZ)
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_route.name.split('.')[-1]}") as tmp_route:
                    tmp_route.write(uploaded_route.getvalue())
                    tmp_route_path = tmp_route.name

                # 2. Save uploaded template to a temp file (if provided)
                tmp_template_path = None
                if uploaded_template:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_template:
                        tmp_template.write(uploaded_template.getvalue())
                        tmp_template_path = tmp_template.name

              # 3. Process the file
                try:
                    # A. Unzip KMZ if necessary
                    if uploaded_route.name.lower().endswith('.kmz'):
                        with zipfile.ZipFile(tmp_route_path, 'r') as kmz:
                            kml_filename = [f for f in kmz.namelist() if f.endswith('.kml')][0]
                            read_path = kmz.extract(kml_filename, path=tempfile.gettempdir())
                    else:
                        read_path = tmp_route_path

                    # B. Read ALL layers
                    layers = fiona.listlayers(read_path)
                    gdfs = [gpd.read_file(read_path, layer=l) for l in layers if not gpd.read_file(read_path, layer=l).empty]
                    
                    if not gdfs:
                        st.error("⚠️ Could not find any map data in this file.")
                        st.stop()
                        
                    gdf = pd.concat(gdfs, ignore_index=True)
                    gdf = gdf[gdf.geometry.type.isin(['LineString', 'MultiLineString'])]
                    
                    if gdf.empty:
                        st.error("⚠️ No routes found! Please upload a file containing a drawn path.")
                        st.stop()

                    # C. Generate Data & KML
                    # Notice we entirely removed fetch_waypoint_maps()
                    stops = sample_polyline_waypoints(gdf, num_stops=slide_count)
                    stops = generate_waypoints_narrative(stops, api_key=gemini_key)
                    
                    # Generate the KML instead of PPTX
                    kml_file = create_google_earth_kml(stops)
                    st.success("✅ Google Earth Tour complete!")
                    
                    # D. Download Button
                    with open(kml_file, "rb") as file:
                        st.download_button(
                            label="📥 Download Earth Tour (.kml)",
                            data=file,
                            file_name="Automated_Earth_Tour.kml",
                            mime="application/vnd.google-earth.kml+xml"
                        )
                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
                    
                    # 4. Provide the download button
                    with open(pptx_file, "rb") as file:
                        st.download_button(
                            label="📥 Download Slideshow (.pptx)",
                            data=file,
                            file_name="Automated_Route_Deck.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
