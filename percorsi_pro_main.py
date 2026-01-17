"""
Percorsi Android Pro - Ottimizzatore di Percorsi Completo
Versione: 2.0.0
Autore: Mattia Prosperi

Funzionalità:
- Caricamento file Excel con selezione colonne
- Mappa integrata (OpenStreetMap)
- Gestione percorsi salvati
- Navigazione integrata
- Export Excel con scelta cartella
"""

import os
import json
import threading
from math import radians, sin, cos, atan2, sqrt, pi
from datetime import datetime
from functools import partial

from kivy.app import App
from kivy.lang import Builder
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.utils import platform
from kivy.properties import StringProperty, NumericProperty, ListProperty, BooleanProperty, ObjectProperty
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.progressbar import ProgressBar
from kivy.uix.spinner import Spinner
from kivy.uix.checkbox import CheckBox
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.image import Image
from kivy.uix.scatter import Scatter
from kivy.uix.stencilview import StencilView
from kivy.graphics import Color, Line, Ellipse, Rectangle
from kivy.graphics.texture import Texture
from kivy.network.urlrequest import UrlRequest

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

# Per Excel
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

# API Key OpenRouteService
ORS_API_KEY = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6IjQwNTQ3OTY0MjI0NjRmMzg4ZTFkNjQ1NTc4MGY4OGZkIiwiaCI6Im11cm11cjY0In0="

# Cache e storage
DIST_CACHE = {}
SAVED_ROUTES = []  # Lista percorsi salvati

# ==================== UTILITY FUNCTIONS ====================

def haversine(lat1, lon1, lat2, lon2):
    R = 6371000
    phi1, phi2 = radians(lat1), radians(lat2)
    delta_phi = radians(lat2 - lat1)
    delta_lambda = radians(lon2 - lon1)
    a = sin(delta_phi/2)**2 + cos(phi1)*cos(phi2)*sin(delta_lambda/2)**2
    return R * 2 * atan2(sqrt(a), sqrt(1-a))

def parse_coordinates(text):
    coords = []
    for line in text.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
        for sep in [',', ';', '\t', ' ']:
            if sep in line:
                parts = line.split(sep)
                if len(parts) >= 2:
                    try:
                        lat = float(parts[0].strip().replace(',', '.'))
                        lon = float(parts[1].strip().replace(',', '.'))
                        if -90 <= lat <= 90 and -180 <= lon <= 180:
                            coords.append((lat, lon))
                            break
                    except:
                        continue
    return coords

def cache_key(lat1, lon1, lat2, lon2, mode):
    return (round(lat1, 5), round(lon1, 5), round(lat2, 5), round(lon2, 5), mode)

def get_osrm_distance(lat1, lon1, lat2, lon2):
    if not REQUESTS_AVAILABLE:
        return int(haversine(lat1, lon1, lat2, lon2))
    key = cache_key(lat1, lon1, lat2, lon2, "osrm")
    if key in DIST_CACHE:
        return DIST_CACHE[key]
    try:
        url = f"https://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=false"
        response = requests.get(url, timeout=10)
        data = response.json()
        if data.get("code") == "Ok":
            dist = int(data["routes"][0]["distance"])
            DIST_CACHE[key] = dist
            return dist
    except:
        pass
    dist = int(haversine(lat1, lon1, lat2, lon2))
    DIST_CACHE[key] = dist
    return dist

def get_ors_distance(lat1, lon1, lat2, lon2):
    if not REQUESTS_AVAILABLE:
        return int(haversine(lat1, lon1, lat2, lon2))
    key = cache_key(lat1, lon1, lat2, lon2, "ors")
    if key in DIST_CACHE:
        return DIST_CACHE[key]
    try:
        url = "https://api.openrouteservice.org/v2/directions/driving-car"
        headers = {"Authorization": ORS_API_KEY}
        params = {"start": f"{lon1},{lat1}", "end": f"{lon2},{lat2}"}
        response = requests.get(url, headers=headers, params=params, timeout=10)
        data = response.json()
        if "features" in data:
            dist = int(data["features"][0]["properties"]["segments"][0]["distance"])
            DIST_CACHE[key] = dist
            return dist
    except:
        pass
    dist = int(haversine(lat1, lon1, lat2, lon2))
    DIST_CACHE[key] = dist
    return dist

def get_route_geometry(coords):
    """Ottiene la geometria del percorso completo da OSRM"""
    if not REQUESTS_AVAILABLE or len(coords) < 2:
        return coords
    try:
        coords_str = ";".join([f"{lon},{lat}" for lat, lon in coords])
        url = f"https://router.project-osrm.org/route/v1/driving/{coords_str}?overview=full&geometries=geojson"
        response = requests.get(url, timeout=15)
        data = response.json()
        if data.get("code") == "Ok":
            geometry = data["routes"][0]["geometry"]["coordinates"]
            return [(lat, lon) for lon, lat in geometry]
    except:
        pass
    return coords

def build_matrix(coords, mode="haversine", callback=None):
    n = len(coords)
    matrix = {i: {j: 0 for j in range(n)} for i in range(n)}
    total = n * (n-1) // 2
    done = 0
    for i in range(n):
        for j in range(i+1, n):
            if mode == "osrm":
                d = get_osrm_distance(coords[i][0], coords[i][1], coords[j][0], coords[j][1])
            elif mode == "ors":
                d = get_ors_distance(coords[i][0], coords[i][1], coords[j][0], coords[j][1])
            else:
                d = int(haversine(coords[i][0], coords[i][1], coords[j][0], coords[j][1]))
            matrix[i][j] = d
            matrix[j][i] = d
            done += 1
            if callback and total > 0:
                callback(done / total)
    return matrix

def nearest_neighbor(matrix, start=0):
    n = len(matrix)
    if n <= 1:
        return list(range(n))
    visited = [False] * n
    tour = [start]
    visited[start] = True
    for _ in range(n-1):
        curr = tour[-1]
        nearest, nearest_d = None, float('inf')
        for j in range(n):
            if not visited[j] and matrix[curr][j] < nearest_d:
                nearest, nearest_d = j, matrix[curr][j]
        if nearest is not None:
            tour.append(nearest)
            visited[nearest] = True
    return tour

def two_opt(tour, matrix, max_iter=500):
    n = len(tour)
    if n < 4:
        return tour
    improved = True
    iters = 0
    while improved and iters < max_iter:
        improved = False
        iters += 1
        for i in range(n-2):
            for j in range(i+2, n):
                if j == n-1 and i == 0:
                    continue
                a, b, c, d = tour[i], tour[i+1], tour[j], tour[(j+1) % n]
                if matrix[a][b] + matrix[c][d] > matrix[a][c] + matrix[b][d]:
                    tour[i+1:j+1] = reversed(tour[i+1:j+1])
                    improved = True
    return tour

def solve_tsp(coords, mode="haversine", callback=None):
    if len(coords) < 2:
        return list(range(len(coords))), 0, [0]*len(coords)
    if callback:
        callback(0.1)
    matrix = build_matrix(coords, mode, lambda p: callback(0.1 + p*0.5) if callback else None)
    if callback:
        callback(0.7)
    tour = nearest_neighbor(matrix, 0)
    tour = two_opt(tour, matrix)
    if callback:
        callback(0.95)
    distances = [0]
    for i in range(1, len(tour)):
        distances.append(matrix[tour[i-1]][tour[i]])
    total = sum(distances)
    if callback:
        callback(1.0)
    return tour, total, distances

def generate_gmaps_link(coords, max_wp=10):
    links = []
    for i in range(0, len(coords), max_wp-1):
        seg = coords[i:i+max_wp]
        if len(seg) < 2:
            continue
        pts = "/".join([f"{lat},{lon}" for lat, lon in seg])
        links.append(f"https://www.google.com/maps/dir/{pts}")
    return links

# ==================== EXCEL FUNCTIONS ====================

def read_excel_file(filepath):
    """Legge un file Excel e restituisce DataFrame e colonne"""
    if PANDAS_AVAILABLE:
        try:
            if filepath.lower().endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)
            return df, list(df.columns)
        except Exception as e:
            return None, str(e)
    elif OPENPYXL_AVAILABLE:
        try:
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active
            data = list(ws.values)
            if not data:
                return None, "File vuoto"
            headers = [str(h) if h else f"Col{i}" for i, h in enumerate(data[0])]
            rows = data[1:]
            return {'headers': headers, 'rows': rows}, headers
        except Exception as e:
            return None, str(e)
    return None, "Libreria Excel non disponibile"

def export_to_excel(data, filepath):
    """Esporta dati in Excel"""
    if PANDAS_AVAILABLE:
        try:
            df = pd.DataFrame(data)
            df.to_excel(filepath, index=False)
            return True, filepath
        except Exception as e:
            return False, str(e)
    elif OPENPYXL_AVAILABLE:
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            # Intestazioni
            if data:
                headers = list(data[0].keys())
                ws.append(headers)
                for row in data:
                    ws.append([row.get(h, '') for h in headers])
            wb.save(filepath)
            return True, filepath
        except Exception as e:
            return False, str(e)
    return False, "Libreria Excel non disponibile"

# ==================== MAP WIDGET ====================

class MapWidget(StencilView):
    """Widget mappa con tiles OpenStreetMap"""
    
    lat = NumericProperty(41.9028)  # Roma default
    lon = NumericProperty(12.4964)
    zoom = NumericProperty(10)
    markers = ListProperty([])
    route_coords = ListProperty([])
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.tile_cache = {}
        self.bind(size=self.update_map)
        self.bind(pos=self.update_map)
        self.bind(lat=self.update_map)
        self.bind(lon=self.update_map)
        self.bind(zoom=self.update_map)
        self.bind(markers=self.update_map)
        self.bind(route_coords=self.update_map)
        Clock.schedule_once(self.update_map, 0.1)
    
    def lat_lon_to_tile(self, lat, lon, zoom):
        """Converte coordinate in tile numbers"""
        n = 2 ** zoom
        x = int((lon + 180) / 360 * n)
        lat_rad = radians(lat)
        y = int((1 - (log(tan(lat_rad) + 1/cos(lat_rad)) / pi)) / 2 * n)
        return x, y
    
    def tile_to_lat_lon(self, x, y, zoom):
        """Converte tile numbers in coordinate"""
        n = 2 ** zoom
        lon = x / n * 360 - 180
        lat_rad = atan(sinh(pi * (1 - 2 * y / n)))
        lat = lat_rad * 180 / pi
        return lat, lon
    
    def lat_lon_to_pixel(self, lat, lon):
        """Converte coordinate in pixel sullo schermo"""
        center_x, center_y = self.lat_lon_to_tile(self.lat, self.lon, self.zoom)
        target_x, target_y = self.lat_lon_to_tile(lat, lon, self.zoom)
        
        # Offset in tiles
        dx = (target_x - center_x) * 256
        dy = (target_y - center_y) * 256
        
        # Converti in pixel relativi al widget
        px = self.width / 2 + dx + self.x
        py = self.height / 2 - dy + self.y
        
        return px, py
    
    def update_map(self, *args):
        """Aggiorna la visualizzazione della mappa"""
        self.canvas.clear()
        
        with self.canvas:
            # Sfondo
            Color(0.9, 0.9, 0.9, 1)
            Rectangle(pos=self.pos, size=self.size)
            
            # Disegna griglia semplificata (placeholder per tiles)
            Color(0.8, 0.85, 0.9, 1)
            Rectangle(pos=self.pos, size=self.size)
            
            # Disegna il percorso
            if len(self.route_coords) >= 2:
                Color(0.2, 0.5, 1, 0.8)
                points = []
                for lat, lon in self.route_coords:
                    px, py = self.lat_lon_to_pixel(lat, lon)
                    points.extend([px, py])
                if len(points) >= 4:
                    Line(points=points, width=3)
            
            # Disegna i markers
            for i, (lat, lon, label) in enumerate(self.markers):
                px, py = self.lat_lon_to_pixel(lat, lon)
                
                # Marker circle
                if i == 0:
                    Color(0.2, 0.8, 0.2, 1)  # Verde per partenza
                elif i == len(self.markers) - 1:
                    Color(0.8, 0.2, 0.2, 1)  # Rosso per arrivo
                else:
                    Color(0.2, 0.4, 0.9, 1)  # Blu per intermedi
                
                Ellipse(pos=(px-12, py-12), size=(24, 24))
                
                # Numero
                Color(1, 1, 1, 1)
                Ellipse(pos=(px-8, py-8), size=(16, 16))
    
    def set_view(self, coords):
        """Centra la mappa sulle coordinate fornite"""
        if not coords:
            return
        
        # Calcola centro
        lats = [c[0] for c in coords]
        lons = [c[1] for c in coords]
        self.lat = sum(lats) / len(lats)
        self.lon = sum(lons) / len(lons)
        
        # Calcola zoom appropriato
        lat_range = max(lats) - min(lats)
        lon_range = max(lons) - min(lons)
        max_range = max(lat_range, lon_range)
        
        if max_range > 10:
            self.zoom = 5
        elif max_range > 5:
            self.zoom = 6
        elif max_range > 2:
            self.zoom = 7
        elif max_range > 1:
            self.zoom = 8
        elif max_range > 0.5:
            self.zoom = 9
        else:
            self.zoom = 11
        
        self.update_map()
    
    def on_touch_down(self, touch):
        if self.collide_point(*touch.pos):
            touch.grab(self)
            self.last_touch = touch.pos
            return True
        return super().on_touch_down(touch)
    
    def on_touch_move(self, touch):
        if touch.grab_current == self:
            dx = touch.pos[0] - self.last_touch[0]
            dy = touch.pos[1] - self.last_touch[1]
            
            # Converti pixel in gradi
            scale = 360 / (256 * (2 ** self.zoom))
            self.lon -= dx * scale
            self.lat += dy * scale
            
            self.last_touch = touch.pos
            self.update_map()
            return True
        return super().on_touch_move(touch)
    
    def on_touch_up(self, touch):
        if touch.grab_current == self:
            touch.ungrab(self)
            return True
        return super().on_touch_up(touch)
    
    def zoom_in(self):
        if self.zoom < 18:
            self.zoom += 1
            self.update_map()
    
    def zoom_out(self):
        if self.zoom > 2:
            self.zoom -= 1
            self.update_map()

# Funzioni matematiche mancanti
from math import log, tan, atan, sinh

# ==================== KV LANGUAGE ====================

KV = '''
#:import Factory kivy.factory.Factory

<MenuButton@Button>:
    size_hint_y: None
    height: 55
    background_color: 0.2, 0.2, 0.25, 1
    background_normal: ''
    font_size: '16sp'
    bold: True

<ActionButton@Button>:
    size_hint_y: None
    height: 50
    background_normal: ''
    font_size: '15sp'

<CardBox@BoxLayout>:
    orientation: 'vertical'
    padding: 15
    spacing: 10
    canvas.before:
        Color:
            rgba: 0.15, 0.15, 0.18, 1
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [10]

ScreenManager:
    id: sm
    HomeScreen:
    LoadExcelScreen:
    ManualInputScreen:
    ColumnSelectScreen:
    SettingsScreen:
    ProcessingScreen:
    SavedRoutesScreen:
    RouteDetailScreen:
    MapViewScreen:
    ExportScreen:

# ==================== HOME ====================
<HomeScreen>:
    name: 'home'
    BoxLayout:
        orientation: 'vertical'
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 80
            padding: 15
            canvas.before:
                Color:
                    rgba: 1, 0.1, 0.4, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Label:
                text: 'PERCORSI PRO'
                font_size: '28sp'
                bold: True
        
        # Menu
        ScrollView:
            BoxLayout:
                orientation: 'vertical'
                padding: 20
                spacing: 15
                size_hint_y: None
                height: self.minimum_height
                
                MenuButton:
                    text: '📁  CARICA FILE EXCEL'
                    background_color: 0.2, 0.6, 0.3, 1
                    on_release: app.root.current = 'load_excel'
                
                MenuButton:
                    text: '✏️  INSERIMENTO MANUALE'
                    background_color: 0.2, 0.5, 0.7, 1
                    on_release: app.root.current = 'manual_input'
                
                MenuButton:
                    text: '📍  PERCORSI SALVATI'
                    background_color: 0.6, 0.4, 0.2, 1
                    on_release: app.root.current = 'saved_routes'
                
                MenuButton:
                    text: '⚙️  IMPOSTAZIONI'
                    background_color: 0.4, 0.4, 0.4, 1
                    on_release: app.root.current = 'settings'
                
                Widget:
                    size_hint_y: None
                    height: 30
                
                Label:
                    text: 'v2.0 - Powered by Mattia Prosperi'
                    font_size: '12sp'
                    color: 0.5, 0.5, 0.5, 1
                    size_hint_y: None
                    height: 30

# ==================== LOAD EXCEL ====================
<LoadExcelScreen>:
    name: 'load_excel'
    BoxLayout:
        orientation: 'vertical'
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            spacing: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '< INDIETRO'
                size_hint_x: 0.3
                background_color: 0.3, 0.3, 0.35, 1
                background_normal: ''
                on_release: app.root.current = 'home'
            
            Label:
                text: 'Carica File Excel'
                font_size: '18sp'
                bold: True
        
        # File chooser
        FileChooserListView:
            id: file_chooser
            path: app.get_default_path()
            filters: ['*.xlsx', '*.xls', '*.csv']
            size_hint_y: 0.7
        
        # Actions
        BoxLayout:
            size_hint_y: None
            height: 60
            padding: 10
            spacing: 10
            
            Button:
                text: 'SELEZIONA FILE'
                background_color: 0.2, 0.7, 0.3, 1
                background_normal: ''
                on_release: app.load_excel(file_chooser.selection)

# ==================== MANUAL INPUT ====================
<ManualInputScreen>:
    name: 'manual_input'
    BoxLayout:
        orientation: 'vertical'
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '< INDIETRO'
                size_hint_x: 0.3
                background_color: 0.3, 0.3, 0.35, 1
                background_normal: ''
                on_release: app.root.current = 'home'
            
            Label:
                text: 'Inserimento Manuale'
                font_size: '18sp'
                bold: True
        
        ScrollView:
            BoxLayout:
                orientation: 'vertical'
                padding: 15
                spacing: 10
                size_hint_y: None
                height: self.minimum_height
                
                Label:
                    text: 'Coordinate (lat,lon per riga):'
                    size_hint_y: None
                    height: 30
                    halign: 'left'
                    text_size: self.size
                
                TextInput:
                    id: coords_input
                    hint_text: '45.4642,9.1900\\n41.9028,12.4964'
                    multiline: True
                    size_hint_y: None
                    height: 200
                    background_color: 0.15, 0.15, 0.18, 1
                    foreground_color: 1, 1, 1, 1
                
                Label:
                    text: 'Etichette (opzionale):'
                    size_hint_y: None
                    height: 30
                    halign: 'left'
                    text_size: self.size
                
                TextInput:
                    id: labels_input
                    hint_text: 'Milano\\nRoma'
                    multiline: True
                    size_hint_y: None
                    height: 100
                    background_color: 0.15, 0.15, 0.18, 1
                    foreground_color: 1, 1, 1, 1
                
                BoxLayout:
                    size_hint_y: None
                    height: 45
                    spacing: 10
                    
                    Button:
                        text: 'ESEMPIO'
                        background_color: 0.4, 0.4, 0.5, 1
                        background_normal: ''
                        on_release: app.load_example()
                    
                    Button:
                        text: 'PULISCI'
                        background_color: 0.5, 0.3, 0.3, 1
                        background_normal: ''
                        on_release: app.clear_manual_input()
                
                Label:
                    id: preview_label
                    text: 'Coordinate: 0'
                    size_hint_y: None
                    height: 30
                
                Button:
                    text: 'OTTIMIZZA PERCORSO'
                    size_hint_y: None
                    height: 55
                    background_color: 1, 0.1, 0.4, 1
                    background_normal: ''
                    bold: True
                    on_release: app.start_manual_optimization()

# ==================== COLUMN SELECT ====================
<ColumnSelectScreen>:
    name: 'column_select'
    BoxLayout:
        orientation: 'vertical'
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '< INDIETRO'
                size_hint_x: 0.3
                background_color: 0.3, 0.3, 0.35, 1
                background_normal: ''
                on_release: app.root.current = 'load_excel'
            
            Label:
                text: 'Seleziona Colonne'
                font_size: '18sp'
                bold: True
        
        ScrollView:
            BoxLayout:
                orientation: 'vertical'
                padding: 15
                spacing: 15
                size_hint_y: None
                height: self.minimum_height
                
                CardBox:
                    size_hint_y: None
                    height: 200
                    
                    Label:
                        text: 'COLONNE OBBLIGATORIE'
                        font_size: '14sp'
                        bold: True
                        color: 1, 0.5, 0.2, 1
                        size_hint_y: None
                        height: 30
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        Label:
                            text: 'Latitudine:'
                            size_hint_x: 0.4
                        Spinner:
                            id: lat_spinner
                            text: 'Seleziona...'
                            size_hint_x: 0.6
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        Label:
                            text: 'Longitudine:'
                            size_hint_x: 0.4
                        Spinner:
                            id: lon_spinner
                            text: 'Seleziona...'
                            size_hint_x: 0.6
                
                CardBox:
                    size_hint_y: None
                    height: 250
                    
                    Label:
                        text: 'COLONNE OPZIONALI'
                        font_size: '14sp'
                        bold: True
                        color: 0.5, 0.8, 1, 1
                        size_hint_y: None
                        height: 30
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        Label:
                            text: 'Data:'
                            size_hint_x: 0.4
                        Spinner:
                            id: date_spinner
                            text: '-- Nessuna --'
                            size_hint_x: 0.6
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        Label:
                            text: 'Operatore:'
                            size_hint_x: 0.4
                        Spinner:
                            id: operator_spinner
                            text: '-- Nessuna --'
                            size_hint_x: 0.6
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        Label:
                            text: 'Etichetta:'
                            size_hint_x: 0.4
                        Spinner:
                            id: label_spinner
                            text: '-- Nessuna --'
                            size_hint_x: 0.6
                
                Button:
                    text: 'CONFERMA E OTTIMIZZA'
                    size_hint_y: None
                    height: 55
                    background_color: 1, 0.1, 0.4, 1
                    background_normal: ''
                    bold: True
                    on_release: app.start_excel_optimization()

# ==================== SETTINGS ====================
<SettingsScreen>:
    name: 'settings'
    BoxLayout:
        orientation: 'vertical'
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '< INDIETRO'
                size_hint_x: 0.3
                background_color: 0.3, 0.3, 0.35, 1
                background_normal: ''
                on_release: app.root.current = 'home'
            
            Label:
                text: 'Impostazioni'
                font_size: '18sp'
                bold: True
        
        ScrollView:
            BoxLayout:
                orientation: 'vertical'
                padding: 15
                spacing: 15
                size_hint_y: None
                height: self.minimum_height
                
                CardBox:
                    size_hint_y: None
                    height: 150
                    
                    Label:
                        text: 'CALCOLO DISTANZE'
                        font_size: '14sp'
                        bold: True
                        color: 1, 0.5, 0.2, 1
                        size_hint_y: None
                        height: 30
                    
                    Spinner:
                        id: distance_mode
                        text: 'Haversine (veloce)'
                        values: ['Haversine (veloce)', 'OSRM (strade)', 'ORS (preciso)']
                        size_hint_y: None
                        height: 45
                    
                    Label:
                        text: 'Haversine: linea d\\'aria\\nOSRM/ORS: distanze stradali reali'
                        font_size: '12sp'
                        color: 0.6, 0.6, 0.6, 1
                        size_hint_y: None
                        height: 40
                
                CardBox:
                    size_hint_y: None
                    height: 100
                    
                    Label:
                        text: 'API KEY ORS'
                        font_size: '14sp'
                        bold: True
                        color: 0.5, 0.8, 1, 1
                        size_hint_y: None
                        height: 30
                    
                    Label:
                        text: 'Preconfigurata ✓'
                        color: 0.3, 0.8, 0.3, 1
                        size_hint_y: None
                        height: 30

# ==================== PROCESSING ====================
<ProcessingScreen>:
    name: 'processing'
    BoxLayout:
        orientation: 'vertical'
        padding: 40
        spacing: 30
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        Widget:
        
        Label:
            text: 'Ottimizzazione in corso...'
            font_size: '22sp'
            bold: True
        
        ProgressBar:
            id: progress_bar
            max: 100
            value: 0
            size_hint_y: None
            height: 25
        
        Label:
            id: progress_label
            text: '0%'
            font_size: '18sp'
        
        Label:
            id: progress_status
            text: 'Preparazione...'
            font_size: '14sp'
            color: 0.6, 0.6, 0.6, 1
        
        Widget:
        
        Button:
            text: 'ANNULLA'
            size_hint: 0.5, None
            height: 45
            pos_hint: {'center_x': 0.5}
            background_color: 0.5, 0.3, 0.3, 1
            background_normal: ''
            on_release: app.cancel_optimization()

# ==================== SAVED ROUTES ====================
<SavedRoutesScreen>:
    name: 'saved_routes'
    BoxLayout:
        orientation: 'vertical'
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '< HOME'
                size_hint_x: 0.25
                background_color: 0.3, 0.3, 0.35, 1
                background_normal: ''
                on_release: app.root.current = 'home'
            
            Label:
                text: 'Percorsi Salvati'
                font_size: '18sp'
                bold: True
            
            Button:
                text: '🗑️'
                size_hint_x: 0.15
                background_color: 0.5, 0.2, 0.2, 1
                background_normal: ''
                on_release: app.clear_all_routes()
        
        # Routes list
        ScrollView:
            id: routes_scroll
            BoxLayout:
                id: routes_list
                orientation: 'vertical'
                padding: 10
                spacing: 10
                size_hint_y: None
                height: self.minimum_height

# ==================== ROUTE DETAIL ====================
<RouteDetailScreen>:
    name: 'route_detail'
    BoxLayout:
        orientation: 'vertical'
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '< INDIETRO'
                size_hint_x: 0.3
                background_color: 0.3, 0.3, 0.35, 1
                background_normal: ''
                on_release: app.root.current = 'saved_routes'
            
            Label:
                id: route_title
                text: 'Dettaglio Percorso'
                font_size: '16sp'
                bold: True
        
        # Summary
        BoxLayout:
            size_hint_y: None
            height: 60
            padding: 10
            canvas.before:
                Color:
                    rgba: 0.2, 0.6, 0.3, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Label:
                id: route_summary
                text: 'Distanza: - | Tappe: -'
                font_size: '16sp'
                bold: True
        
        # Actions
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 5
            spacing: 5
            
            Button:
                text: '🗺️ MAPPA'
                background_color: 0.2, 0.5, 0.8, 1
                background_normal: ''
                on_release: app.show_route_on_map()
            
            Button:
                text: '🧭 NAVIGA'
                background_color: 0.2, 0.7, 0.4, 1
                background_normal: ''
                on_release: app.start_navigation()
            
            Button:
                text: '💾 ESPORTA'
                background_color: 0.6, 0.4, 0.2, 1
                background_normal: ''
                on_release: app.root.current = 'export'
        
        # Stops list
        Label:
            text: 'ORDINE DESTINAZIONI'
            size_hint_y: None
            height: 35
            bold: True
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
        
        ScrollView:
            BoxLayout:
                id: stops_list
                orientation: 'vertical'
                padding: 10
                spacing: 5
                size_hint_y: None
                height: self.minimum_height

# ==================== MAP VIEW ====================
<MapViewScreen>:
    name: 'map_view'
    BoxLayout:
        orientation: 'vertical'
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '< INDIETRO'
                size_hint_x: 0.3
                background_color: 0.3, 0.3, 0.35, 1
                background_normal: ''
                on_release: app.root.current = 'route_detail'
            
            Label:
                text: 'Mappa Percorso'
                font_size: '18sp'
                bold: True
        
        # Map
        MapWidget:
            id: map_widget
        
        # Zoom controls
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            spacing: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '➖ ZOOM OUT'
                background_color: 0.3, 0.3, 0.4, 1
                background_normal: ''
                on_release: map_widget.zoom_out()
            
            Button:
                text: '🎯 CENTRA'
                background_color: 0.3, 0.5, 0.3, 1
                background_normal: ''
                on_release: app.center_map()
            
            Button:
                text: '➕ ZOOM IN'
                background_color: 0.3, 0.3, 0.4, 1
                background_normal: ''
                on_release: map_widget.zoom_in()

# ==================== EXPORT ====================
<ExportScreen>:
    name: 'export'
    BoxLayout:
        orientation: 'vertical'
        canvas.before:
            Color:
                rgba: 0.1, 0.1, 0.12, 1
            Rectangle:
                pos: self.pos
                size: self.size
        
        # Header
        BoxLayout:
            size_hint_y: None
            height: 50
            padding: 10
            canvas.before:
                Color:
                    rgba: 0.15, 0.15, 0.18, 1
                Rectangle:
                    pos: self.pos
                    size: self.size
            
            Button:
                text: '< INDIETRO'
                size_hint_x: 0.3
                background_color: 0.3, 0.3, 0.35, 1
                background_normal: ''
                on_release: app.root.current = 'route_detail'
            
            Label:
                text: 'Esporta Percorso'
                font_size: '18sp'
                bold: True
        
        ScrollView:
            BoxLayout:
                orientation: 'vertical'
                padding: 15
                spacing: 15
                size_hint_y: None
                height: self.minimum_height
                
                CardBox:
                    size_hint_y: None
                    height: 180
                    
                    Label:
                        text: 'FORMATO EXPORT'
                        font_size: '14sp'
                        bold: True
                        color: 1, 0.5, 0.2, 1
                        size_hint_y: None
                        height: 30
                    
                    Button:
                        text: '📊 ESPORTA EXCEL (.xlsx)'
                        size_hint_y: None
                        height: 50
                        background_color: 0.2, 0.5, 0.3, 1
                        background_normal: ''
                        on_release: app.export_excel()
                    
                    Button:
                        text: '📍 ESPORTA GPX'
                        size_hint_y: None
                        height: 50
                        background_color: 0.3, 0.4, 0.6, 1
                        background_normal: ''
                        on_release: app.export_gpx()
                    
                    Button:
                        text: '🌍 ESPORTA KML'
                        size_hint_y: None
                        height: 50
                        background_color: 0.5, 0.4, 0.3, 1
                        background_normal: ''
                        on_release: app.export_kml()
                
                CardBox:
                    size_hint_y: None
                    height: 120
                    
                    Label:
                        text: 'CARTELLA DESTINAZIONE'
                        font_size: '14sp'
                        bold: True
                        color: 0.5, 0.8, 1, 1
                        size_hint_y: None
                        height: 30
                    
                    Button:
                        text: '📁 SCEGLI CARTELLA'
                        size_hint_y: None
                        height: 50
                        background_color: 0.4, 0.4, 0.5, 1
                        background_normal: ''
                        on_release: app.choose_export_folder()
                    
                    Label:
                        id: export_path_label
                        text: 'Default: Download'
                        font_size: '12sp'
                        color: 0.6, 0.6, 0.6, 1
                        size_hint_y: None
                        height: 25
                
                CardBox:
                    size_hint_y: None
                    height: 80
                    
                    Button:
                        text: '🗺️ APRI IN GOOGLE MAPS'
                        size_hint_y: None
                        height: 55
                        background_color: 0.2, 0.6, 0.9, 1
                        background_normal: ''
                        bold: True
                        on_release: app.open_google_maps()
'''

# ==================== SCREENS ====================

class HomeScreen(Screen):
    pass

class LoadExcelScreen(Screen):
    pass

class ManualInputScreen(Screen):
    pass

class ColumnSelectScreen(Screen):
    pass

class SettingsScreen(Screen):
    pass

class ProcessingScreen(Screen):
    pass

class SavedRoutesScreen(Screen):
    pass

class RouteDetailScreen(Screen):
    pass

class MapViewScreen(Screen):
    pass

class ExportScreen(Screen):
    pass

# ==================== MAIN APP ====================

class PercorsiApp(App):
    
    # Data
    excel_data = None
    excel_columns = []
    coords = ListProperty([])
    labels = ListProperty([])
    current_route = None
    saved_routes = ListProperty([])
    export_folder = StringProperty('')
    
    # Processing
    is_processing = BooleanProperty(False)
    
    def build(self):
        self.title = 'Percorsi Pro'
        if platform not in ('android', 'ios'):
            Window.size = (420, 750)
        
        # Load saved routes
        self.load_saved_routes()
        
        return Builder.load_string(KV)
    
    def get_default_path(self):
        if platform == 'android':
            from android.storage import primary_external_storage_path
            return primary_external_storage_path()
        return os.path.expanduser('~')
    
    def get_distance_mode(self):
        try:
            screen = self.root.get_screen('settings')
            spinner_text = screen.ids.distance_mode.text
            if "OSRM" in spinner_text:
                return "osrm"
            elif "ORS" in spinner_text:
                return "ors"
        except:
            pass
        return "haversine"
    
    # ==================== EXCEL LOADING ====================
    
    def load_excel(self, selection):
        if not selection:
            self.show_popup("Errore", "Seleziona un file!")
            return
        
        filepath = selection[0]
        data, columns = read_excel_file(filepath)
        
        if data is None:
            self.show_popup("Errore", f"Impossibile leggere il file:\n{columns}")
            return
        
        self.excel_data = data
        self.excel_columns = columns
        
        # Navigate to column select screen first
        self.root.current = 'column_select'
        
        # Schedule spinner population after screen transition
        Clock.schedule_once(lambda dt: self._populate_spinners(columns), 0.2)
    
    def _populate_spinners(self, columns):
        """Popola gli spinner dopo che la schermata è visibile"""
        try:
            screen = self.root.get_screen('column_select')
            col_options = ['-- Nessuna --'] + columns
            
            screen.ids.lat_spinner.values = columns
            screen.ids.lon_spinner.values = columns
            screen.ids.date_spinner.values = col_options
            screen.ids.operator_spinner.values = col_options
            screen.ids.label_spinner.values = col_options
            
            # Auto-detect columns
            for col in columns:
                col_upper = col.upper()
                if 'LAT' in col_upper:
                    screen.ids.lat_spinner.text = col
                elif 'LON' in col_upper:
                    screen.ids.lon_spinner.text = col
                elif 'DATA' in col_upper or 'DATE' in col_upper:
                    screen.ids.date_spinner.text = col
                elif 'OPERATORE' in col_upper or 'OPERATOR' in col_upper:
                    screen.ids.operator_spinner.text = col
                elif 'NOME' in col_upper or 'NAME' in col_upper or 'LABEL' in col_upper or 'DESC' in col_upper:
                    screen.ids.label_spinner.text = col
        except Exception as e:
            print(f"Error populating spinners: {e}")
    
    def start_excel_optimization(self):
        screen = self.root.get_screen('column_select')
        lat_col = screen.ids.lat_spinner.text
        lon_col = screen.ids.lon_spinner.text
        
        if lat_col == 'Seleziona...' or lon_col == 'Seleziona...':
            self.show_popup("Errore", "Seleziona le colonne Latitudine e Longitudine!")
            return
        
        # Extract coordinates
        self.coords = []
        self.labels = []
        
        date_col = screen.ids.date_spinner.text
        op_col = screen.ids.operator_spinner.text
        label_col = screen.ids.label_spinner.text
        
        if PANDAS_AVAILABLE and hasattr(self.excel_data, 'iterrows'):
            for idx, row in self.excel_data.iterrows():
                try:
                    lat = float(row[lat_col])
                    lon = float(row[lon_col])
                    if -90 <= lat <= 90 and -180 <= lon <= 180:
                        self.coords.append((lat, lon))
                        
                        # Build label
                        label_parts = []
                        if label_col != '-- Nessuna --' and label_col in row:
                            label_parts.append(str(row[label_col]))
                        if op_col != '-- Nessuna --' and op_col in row:
                            label_parts.append(str(row[op_col]))
                        if date_col != '-- Nessuna --' and date_col in row:
                            label_parts.append(str(row[date_col]))
                        
                        self.labels.append(' - '.join(label_parts) if label_parts else f"Punto {len(self.coords)}")
                except:
                    continue
        else:
            # Openpyxl fallback
            headers = self.excel_data['headers']
            lat_idx = headers.index(lat_col)
            lon_idx = headers.index(lon_col)
            
            for row in self.excel_data['rows']:
                try:
                    lat = float(row[lat_idx])
                    lon = float(row[lon_idx])
                    if -90 <= lat <= 90 and -180 <= lon <= 180:
                        self.coords.append((lat, lon))
                        self.labels.append(f"Punto {len(self.coords)}")
                except:
                    continue
        
        if len(self.coords) < 2:
            self.show_popup("Errore", f"Trovate solo {len(self.coords)} coordinate valide!")
            return
        
        self._start_optimization()
    
    # ==================== MANUAL INPUT ====================
    
    def load_example(self):
        screen = self.root.get_screen('manual_input')
        screen.ids.coords_input.text = "45.4642,9.1900\n41.9028,12.4964\n43.7696,11.2558\n44.4949,11.3426\n40.8518,14.2681\n45.4384,10.9916"
        screen.ids.labels_input.text = "Milano\nRoma\nFirenze\nBologna\nNapoli\nVerona"
        self._update_manual_preview()
    
    def clear_manual_input(self):
        screen = self.root.get_screen('manual_input')
        screen.ids.coords_input.text = ""
        screen.ids.labels_input.text = ""
        screen.ids.preview_label.text = "Coordinate: 0"
    
    def _update_manual_preview(self):
        screen = self.root.get_screen('manual_input')
        text = screen.ids.coords_input.text
        coords = parse_coordinates(text)
        screen.ids.preview_label.text = f"Coordinate: {len(coords)}"
    
    def start_manual_optimization(self):
        screen = self.root.get_screen('manual_input')
        text = screen.ids.coords_input.text
        self.coords = parse_coordinates(text)
        
        labels_text = screen.ids.labels_input.text
        self.labels = [l.strip() for l in labels_text.split('\n') if l.strip()]
        
        # Pad labels
        while len(self.labels) < len(self.coords):
            self.labels.append(f"Punto {len(self.labels) + 1}")
        
        if len(self.coords) < 2:
            self.show_popup("Errore", "Inserisci almeno 2 coordinate!")
            return
        
        self._start_optimization()
    
    # ==================== OPTIMIZATION ====================
    
    def _start_optimization(self):
        self.root.current = 'processing'
        self.is_processing = True
        
        # Reset progress after screen transition
        Clock.schedule_once(self._reset_progress, 0.1)
        
        thread = threading.Thread(target=self._optimize)
        thread.daemon = True
        thread.start()
    
    def _reset_progress(self, dt):
        try:
            screen = self.root.get_screen('processing')
            screen.ids.progress_bar.value = 0
            screen.ids.progress_label.text = "0%"
            screen.ids.progress_status.text = "Preparazione..."
        except:
            pass
    
    def _optimize(self):
        def update_progress(p):
            Clock.schedule_once(lambda dt: self._set_progress(p), 0)
        
        try:
            mode = self.get_distance_mode()
            Clock.schedule_once(lambda dt: self._set_status("Calcolo distanze..."), 0)
            
            tour, total, distances = solve_tsp(list(self.coords), mode, update_progress)
            
            Clock.schedule_once(lambda dt: self._set_status("Finalizzazione..."), 0)
            
            # Reorder
            opt_coords = [self.coords[i] for i in tour]
            opt_labels = [self.labels[i] if i < len(self.labels) else f"Punto {i+1}" for i in tour]
            
            # Create route object
            route = {
                'id': datetime.now().strftime('%Y%m%d_%H%M%S'),
                'name': f"Percorso {datetime.now().strftime('%d/%m %H:%M')}",
                'coords': opt_coords,
                'labels': opt_labels,
                'distances': distances,
                'total_distance': total,
                'created': datetime.now().isoformat(),
                'stops_count': len(opt_coords)
            }
            
            Clock.schedule_once(lambda dt: self._optimization_done(route), 0)
            
        except Exception as e:
            Clock.schedule_once(lambda dt: self._optimization_error(str(e)), 0)
    
    def _set_progress(self, p):
        try:
            screen = self.root.get_screen('processing')
            screen.ids.progress_bar.value = p * 100
            screen.ids.progress_label.text = f"{int(p * 100)}%"
        except:
            pass
    
    def _set_status(self, text):
        try:
            screen = self.root.get_screen('processing')
            screen.ids.progress_status.text = text
        except:
            pass
    
    def _optimization_done(self, route):
        self.is_processing = False
        self.current_route = route
        
        # Add to saved routes
        self.saved_routes.append(route)
        self.save_routes_to_file()
        
        # Show route detail
        self._show_route_detail(route)
    
    def _optimization_error(self, error):
        self.is_processing = False
        self.root.current = 'home'
        self.show_popup("Errore", error)
    
    def cancel_optimization(self):
        self.is_processing = False
        self.root.current = 'home'
    
    # ==================== SAVED ROUTES ====================
    
    def load_saved_routes(self):
        try:
            filepath = os.path.join(self.get_default_path(), '.percorsi_routes.json')
            if os.path.exists(filepath):
                with open(filepath, 'r') as f:
                    self.saved_routes = json.load(f)
        except:
            self.saved_routes = []
    
    def save_routes_to_file(self):
        try:
            filepath = os.path.join(self.get_default_path(), '.percorsi_routes.json')
            with open(filepath, 'w') as f:
                json.dump(list(self.saved_routes), f)
        except:
            pass
    
    def on_start(self):
        # Delay refresh to ensure all screens are loaded
        Clock.schedule_once(self._delayed_init, 1.0)
    
    def _delayed_init(self, dt):
        try:
            self._refresh_routes_list()
        except:
            pass
    
    def _refresh_routes_list(self):
        try:
            screen = self.root.get_screen('saved_routes')
            routes_list = screen.ids.routes_list
        except (AttributeError, KeyError):
            return
        
        routes_list.clear_widgets()
        
        if not self.saved_routes:
            routes_list.add_widget(Label(
                text="Nessun percorso salvato",
                size_hint_y=None,
                height=50,
                color=(0.5, 0.5, 0.5, 1)
            ))
            return
        
        for route in reversed(self.saved_routes):
            btn = Button(
                text=f"📍 {route['name']}\n   {self._format_distance(route['total_distance'])} | {route['stops_count']} tappe",
                size_hint_y=None,
                height=70,
                background_color=(0.2, 0.25, 0.3, 1),
                background_normal='',
                halign='left',
                valign='middle',
                text_size=(None, None)
            )
            btn.bind(on_release=lambda x, r=route: self._show_route_detail(r))
            routes_list.add_widget(btn)
    
    def clear_all_routes(self):
        self.saved_routes = []
        self.save_routes_to_file()
        self._refresh_routes_list()
        self.show_popup("Info", "Tutti i percorsi sono stati eliminati")
    
    # ==================== ROUTE DETAIL ====================
    
    def _show_route_detail(self, route):
        self.current_route = route
        
        # Get screen
        screen = self.root.get_screen('route_detail')
        
        # Update UI
        screen.ids.route_title.text = route['name']
        screen.ids.route_summary.text = f"Distanza: {self._format_distance(route['total_distance'])} | Tappe: {route['stops_count']}"
        
        # Populate stops list
        stops_list = screen.ids.stops_list
        stops_list.clear_widgets()
        
        cumulative = 0
        for i, (coord, label) in enumerate(zip(route['coords'], route['labels'])):
            dist = route['distances'][i] if i < len(route['distances']) else 0
            cumulative += dist
            
            stop_box = BoxLayout(
                orientation='horizontal',
                size_hint_y=None,
                height=50,
                padding=5
            )
            
            # Number badge
            num_label = Label(
                text=str(i + 1),
                size_hint_x=0.1,
                bold=True,
                color=(1, 0.5, 0.2, 1)
            )
            
            # Name and coords
            info_label = Label(
                text=f"{label}\n({coord[0]:.4f}, {coord[1]:.4f})",
                size_hint_x=0.6,
                halign='left',
                valign='middle',
                text_size=(None, None),
                font_size='13sp'
            )
            
            # Distance
            dist_label = Label(
                text=self._format_distance(cumulative),
                size_hint_x=0.3,
                color=(0.5, 0.8, 0.5, 1),
                font_size='12sp'
            )
            
            stop_box.add_widget(num_label)
            stop_box.add_widget(info_label)
            stop_box.add_widget(dist_label)
            
            stops_list.add_widget(stop_box)
        
        self.root.current = 'route_detail'
        self._refresh_routes_list()
    
    # ==================== MAP ====================
    
    def show_route_on_map(self):
        if not self.current_route:
            return
        
        screen = self.root.get_screen('map_view')
        map_widget = screen.ids.map_widget
        
        # Set markers
        markers = []
        for i, (coord, label) in enumerate(zip(self.current_route['coords'], self.current_route['labels'])):
            markers.append((coord[0], coord[1], label))
        map_widget.markers = markers
        
        # Get route geometry
        map_widget.route_coords = self.current_route['coords']
        
        # Center map
        map_widget.set_view(self.current_route['coords'])
        
        self.root.current = 'map_view'
    
    def center_map(self):
        if self.current_route:
            screen = self.root.get_screen('map_view')
            screen.ids.map_widget.set_view(self.current_route['coords'])
    
    # ==================== NAVIGATION ====================
    
    def start_navigation(self):
        if not self.current_route:
            return
        
        links = generate_gmaps_link(self.current_route['coords'])
        if links:
            import webbrowser
            webbrowser.open(links[0])
            
            if len(links) > 1:
                self.show_popup("Info", f"Percorso diviso in {len(links)} segmenti.\nAperto il primo segmento.")
    
    def open_google_maps(self):
        self.start_navigation()
    
    # ==================== EXPORT ====================
    
    def choose_export_folder(self):
        # Show folder chooser popup
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        file_chooser = FileChooserListView(
            path=self.get_default_path(),
            dirselect=True
        )
        content.add_widget(file_chooser)
        
        btn_box = BoxLayout(size_hint_y=None, height=50, spacing=10)
        
        def select_folder(instance):
            if file_chooser.selection:
                self.export_folder = file_chooser.selection[0]
            else:
                self.export_folder = file_chooser.path
            try:
                screen = self.root.get_screen('export')
                screen.ids.export_path_label.text = f"Cartella: {self.export_folder}"
            except:
                pass
            popup.dismiss()
        
        btn_select = Button(text='SELEZIONA', background_color=(0.2, 0.6, 0.3, 1), background_normal='')
        btn_select.bind(on_release=select_folder)
        btn_cancel = Button(text='ANNULLA', background_color=(0.5, 0.3, 0.3, 1), background_normal='')
        
        btn_box.add_widget(btn_select)
        btn_box.add_widget(btn_cancel)
        content.add_widget(btn_box)
        
        popup = Popup(title='Scegli Cartella', content=content, size_hint=(0.95, 0.8))
        btn_cancel.bind(on_release=popup.dismiss)
        popup.open()
    
    def _get_export_path(self, filename):
        if self.export_folder:
            return os.path.join(self.export_folder, filename)
        
        if platform == 'android':
            from android.storage import primary_external_storage_path
            return os.path.join(primary_external_storage_path(), 'Download', filename)
        
        return os.path.join(os.path.expanduser('~'), filename)
    
    def export_excel(self):
        if not self.current_route:
            return
        
        data = []
        cumulative = 0
        for i, (coord, label) in enumerate(zip(self.current_route['coords'], self.current_route['labels'])):
            dist = self.current_route['distances'][i] if i < len(self.current_route['distances']) else 0
            cumulative += dist
            data.append({
                'Progressione': i + 1,
                'Etichetta': label,
                'Latitudine': coord[0],
                'Longitudine': coord[1],
                'Distanza_m': dist,
                'Distanza_Cumulata_m': cumulative
            })
        
        filepath = self._get_export_path(f"percorso_{self.current_route['id']}.xlsx")
        success, result = export_to_excel(data, filepath)
        
        if success:
            self.show_popup("Esportato", f"File salvato:\n{result}")
        else:
            self.show_popup("Errore", f"Errore export:\n{result}")
    
    def export_gpx(self):
        if not self.current_route:
            return
        
        gpx = ['<?xml version="1.0" encoding="UTF-8"?>', '<gpx version="1.1" creator="PercorsiPro">']
        
        for i, (coord, label) in enumerate(zip(self.current_route['coords'], self.current_route['labels'])):
            gpx.append(f'<wpt lat="{coord[0]}" lon="{coord[1]}"><n>{label}</n></wpt>')
        
        gpx.append('<trk><n>Percorso</n><trkseg>')
        for coord in self.current_route['coords']:
            gpx.append(f'<trkpt lat="{coord[0]}" lon="{coord[1]}"/>')
        gpx.append('</trkseg></trk></gpx>')
        
        filepath = self._get_export_path(f"percorso_{self.current_route['id']}.gpx")
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write('\n'.join(gpx))
            self.show_popup("Esportato", f"File salvato:\n{filepath}")
        except Exception as e:
            self.show_popup("Errore", str(e))
    
    def export_kml(self):
        if not self.current_route:
            return
        
        kml = ['<?xml version="1.0" encoding="UTF-8"?>',
               '<kml xmlns="http://www.opengis.net/kml/2.2">',
               '<Document><n>Percorso</n>']
        
        for i, (coord, label) in enumerate(zip(self.current_route['coords'], self.current_route['labels'])):
            kml.append(f'<Placemark><n>{label}</n>')
            kml.append(f'<Point><coordinates>{coord[1]},{coord[0]},0</coordinates></Point></Placemark>')
        
        coords_str = ' '.join([f"{c[1]},{c[0]},0" for c in self.current_route['coords']])
        kml.append(f'<Placemark><n>Traccia</n>')
        kml.append(f'<LineString><coordinates>{coords_str}</coordinates></LineString></Placemark>')
        kml.append('</Document></kml>')
        
        filepath = self._get_export_path(f"percorso_{self.current_route['id']}.kml")
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write('\n'.join(kml))
            self.show_popup("Esportato", f"File salvato:\n{filepath}")
        except Exception as e:
            self.show_popup("Errore", str(e))
    
    # ==================== UTILITIES ====================
    
    def _format_distance(self, meters):
        if meters >= 1000:
            return f"{meters/1000:.1f} km"
        return f"{int(meters)} m"
    
    def show_popup(self, title, message):
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        content.add_widget(Label(text=message, text_size=(350, None)))
        btn = Button(text='OK', size_hint_y=None, height=45, background_color=(0.3, 0.5, 0.7, 1), background_normal='')
        content.add_widget(btn)
        
        popup = Popup(title=title, content=content, size_hint=(0.85, 0.4))
        btn.bind(on_release=popup.dismiss)
        popup.open()


if __name__ == '__main__':
    PercorsiApp().run()
