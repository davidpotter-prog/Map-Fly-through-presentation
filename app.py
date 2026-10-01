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

def create_presentation(stops_data, template_path=None, layout_index=6):
    prs = Presentation(template_path) if template_path else Presentation()
    try:
        slide_layout = prs.slide_layouts[layout_index]
    except IndexError:
        slide_layout = prs.slide_layouts[0]
        
    for idx, stop in enumerate(stops_data):
        slide = prs.slides.add_slide(slide_layout)
        # Fallback placement for dynamic boxes
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
        title_box.text_frame.text = f"Stop {idx + 1}: {stop.get('title', '')}"
        
        if stop.get("map_image_path"):
            slide.shapes.add_picture(stop["map_image_path"], Inches(0.5), Inches(1.2), width=Inches(6.5))
            
        notes_box = slide.shapes.add_textbox(Inches(7.2), Inches(1.2), Inches(2.5), Inches(4.5))
        notes_box.text_frame.word_wrap = True
        notes_box.text_frame.text = stop.get("ai_summary", "")
        
    output_path = "route_presentation.pptx"
    prs.save(output_path)
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

                    # B. Read ALL layers in the file (bypassing the empty first-folder problem)
                    layers = fiona.listlayers(read_path)
                    gdfs = []
                    for layer in layers:
                        layer_gdf = gpd.read_file(read_path, layer=layer)
                        if not layer_gdf.empty:
                            gdfs.append(layer_gdf)
                            
                    if not gdfs:
                        st.error("⚠️ Could not find any map data in this file.")
                        st.stop() # Stops execution cleanly
                        
                    # Combine all layers into one master table
                    gdf = pd.concat(gdfs, ignore_index=True)
                    
                    # C. Filter to keep ONLY routes/lines (ignores standalone pins/placemarks)
                    gdf = gdf[gdf.geometry.type.isin(['LineString', 'MultiLineString'])]
                    
                    if gdf.empty:
                        st.error("⚠️ No routes found! Make sure your file contains a drawn path (LineString), not just individual dropped pins.")
                        st.stop()

                    # D. Proceed with slide generation
                    stops = sample_polyline_waypoints(gdf, num_stops=slide_count)
                    stops = fetch_waypoint_maps(stops, api_key=maps_key, map_type=map_type)
                    stops = generate_waypoints_narrative(stops, api_key=gemini_key)
                    
                    pptx_file = create_presentation(stops, template_path=tmp_template_path, layout_index=layout_index)
                    
                    st.success("✅ Presentation complete!")
                    
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
