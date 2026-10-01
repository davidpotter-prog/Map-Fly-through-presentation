import streamlit as st
import os
import json
import geopandas as gpd
from shapely.geometry import LineString
from shapely.ops import linemerge
import numpy as np
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
    merged_geom = gdf.geometry.unary_union
    if merged_geom.geom_type == 'MultiLineString':
        merged_geom = linemerge(merged_geom)
    if merged_geom.geom_type == 'MultiLineString':
        line = list(merged_geom.geoms)[0]
    else:
        line = merged_geom

    total_length = line.length
    distances = np.linspace(0, total_length, num_stops)
    stops = []
    for d in distances:
        pt = line.interpolate(d)
        lookahead_pt = line.interpolate(min(d + (total_length * 0.01), total_length))
        heading = np.degrees(np.arctan2(lookahead_pt.x - pt.x, lookahead_pt.y - pt.y)) % 360
        stops.append({"geometry": pt, "distance_km": round(d / 1000, 2), "heading": heading})
    return stops

def sample_polyline_waypoints_by_distance(gdf, interval_ft=800):
    merged_geom = gdf.geometry.unary_union
    if merged_geom.geom_type == 'MultiLineString':
        merged_geom = linemerge(merged_geom)
        
    if merged_geom.geom_type == 'MultiLineString':
        line_wgs = list(merged_geom.geoms)[0]
    else:
        line_wgs = merged_geom

    temp_gdf = gpd.GeoDataFrame(geometry=[line_wgs], crs="EPSG:4326")
    utm_crs = temp_gdf.estimate_utm_crs()
    line_utm = temp_gdf.to_crs(utm_crs).geometry.iloc[0]

    total_length_m = line_utm.length
    total_length_ft = total_length_m * 3.28084
    interval_m = interval_ft * 0.3048

    if total_length_m == 0:
        return [], 0, 0

    distances_m = list(np.arange(0, total_length_m, interval_m))
    
    if (total_length_m - distances_m[-1]) > (interval_m / 2):
        distances_m.append(total_length_m)

    stops = []
    for d in distances_m:
        fraction = d / total_length_m if total_length_m > 0 else 0
        pt = line_wgs.interpolate(fraction, normalized=True)
        lookahead_pt = line_wgs.interpolate(min(fraction + 0.01, 1.0), normalized=True)
        heading = np.degrees(np.arctan2(lookahead_pt.x - pt.x, lookahead_pt.y - pt.y)) % 360
        
        stops.append({
            "geometry": pt,
            "distance_km": round(d / 1000, 2),
            "distance_ft": round(d * 3.28084, 0),
            "heading": heading
        })

    return stops, total_length_ft, round(total_length_ft / 5280, 2)

def generate_waypoints_narrative(stops_data, api_key):
    client = genai.Client(api_key=api_key)
    for idx, stop in enumerate(stops_data):
        prompt = f"We are at waypoint {idx + 1}. Heading: {stop['heading']}°. Generate a concise, engaging slide title and a brief 2-3 sentence narrative summary for this stop."
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
            stop["ai_summary"] = f"Heading: {stop['heading']}°"
    return stops_data

def create_google_earth_kml(stops_data, gdf, tilt=65, range_meters=800, output_path="earth_presentation.kml"):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>']
    kml.append('<kml xmlns="http://www.opengis.net/kml/2.2">')
    kml.append('  <Document>')
    kml.append('    <name>AI Generated Earth Tour</name>')

    kml.append('    <Style id="routeLineStyle">')
    kml.append('      <LineStyle><color>ff00aaff</color><width>5</width></LineStyle>')
    kml.append('    </Style>')

    kml.append('    <Placemark>')
    kml.append('      <name>Route Path</name>')
    kml.append('      <styleUrl>#routeLineStyle</styleUrl>')
    kml.append('      <MultiGeometry>')
    for geom in gdf.geometry:
        if geom.type == 'LineString':
            coords = " ".join([f"{pt[0]},{pt[1]},0" for pt in geom.coords])
            kml.append(f'        <LineString><coordinates>{coords}</coordinates></LineString>')
        elif geom.type == 'MultiLineString':
            for line in geom.geoms:
                coords = " ".join([f"{pt[0]},{pt[1]},0" for pt in line.coords])
                kml.append(f'        <LineString><coordinates>{coords}</coordinates></LineString>')
    kml.append('      </MultiGeometry>')
    kml.append('    </Placemark>')

    for idx, stop in enumerate(stops_data):
        lon, lat = stop["geometry"].x, stop["geometry"].y
        heading = stop["heading"]
        title = stop.get("title", f"Stop {idx+1}")
        summary = stop.get("ai_summary", "")

        description = f"<![CDATA[<h3>{title}</h3><p>{summary}</p>]]>"

        kml.append('    <Placemark>')
        kml.append(f'      <name>{idx + 1}. {title}</name>')
        kml.append(f'      <description>{description}</description>')
        
        kml.append('      <LookAt>')
        kml.append(f'        <longitude>{lon}</longitude>')
        kml.append(f'        <latitude>{lat}</latitude>')
        kml.append('        <altitude>0</altitude>')
        kml.append(f'        <heading>{heading}</heading>')
        kml.append(f'        <tilt>{tilt}</tilt>')
        kml.append(f'        <range>{range_meters}</range>')
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
st.set_page_config(page_title="Earth Tour Generator", layout="wide")
st.title("🌍 Automated Google Earth Tour Builder")

with st.sidebar:
    st.header("⚙️ Settings")
    gemini_key = st.text_input("Gemini API Key", value=st.secrets.get("GEMINI_API_KEY", ""), type="password")
    
    st.markdown("---")
    st.subheader("📍 Waypoint Mode")
    sampling_mode = st.radio("Waypoint Spacing Mode", ["Every X Feet", "Fixed Number of Stops"])
    
    if sampling_mode == "Every X Feet":
        interval_ft = st.number_input("Waypoint Interval (feet)", min_value=100, max_value=10000, value=800, step=100)
    else:
        slide_count = st.slider("Total Waypoints", min_value=3, max_value=100, value=10)

    st.markdown("---")
    st.subheader("🎥 Camera Angle")
    camera_tilt = st.slider("Camera Tilt", min_value=0, max_value=80, value=65)
    camera_range = st.slider("Camera Height (meters)", min_value=100, max_value=5000, value=800, step=100)

uploaded_route = st.file_uploader("📂 Upload Route File (KMZ, KML, GeoJSON)", type=["kmz", "kml", "geojson", "json"])

if uploaded_route:
    st.success(f"Loaded {uploaded_route.name} successfully!")
    
    if st.button("🚀 Generate Earth Tour", type="primary"):
        if not gemini_key:
            st.error("⚠️ Please provide the Gemini API key in the sidebar.")
        else:
            with st.spinner("Analyzing geometry and writing narrative..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_route.name.split('.')[-1]}") as tmp_route:
                    tmp_route.write(uploaded_route.getvalue())
                    tmp_route_path = tmp_route.name

                try:
                    if uploaded_route.name.lower().endswith('.kmz'):
                        with zipfile.ZipFile(tmp_route_path, 'r') as kmz:
                            kml_filename = [f for f in kmz.namelist() if f.endswith('.kml')][0]
                            read_path = kmz.extract(kml_filename, path=tempfile.gettempdir())
                    else:
                        read_path = tmp_route_path

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

                    if sampling_mode == "Every X Feet":
                        stops, total_ft, total_miles = sample_polyline_waypoints_by_distance(gdf, interval_ft=interval_ft)
                        st.info(f"📏 Route Length: **{total_ft:,.0f} feet** ({total_miles} miles) — Generated **{len(stops)} waypoints** every {interval_ft} ft.")
                    else:
                        stops = sample_polyline_waypoints(gdf, num_stops=slide_count)

                    stops = generate_waypoints_narrative(stops, api_key=gemini_key)
                    
                    kml_file = create_google_earth_kml(
                        stops, 
                        gdf=gdf, 
                        tilt=camera_tilt, 
                        range_meters=camera_range
                    )
                    
                    st.success("✅ Google Earth Tour complete!")
                    
                    with open(kml_file, "rb") as file:
                        st.download_button(
                            label="📥 Download Earth Tour (.kml)",
                            data=file,
                            file_name="Automated_Earth_Tour.kml",
                            mime="application/vnd.google-earth.kml+xml"
                        )
                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
