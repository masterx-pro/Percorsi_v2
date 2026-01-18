"""
Percorsi Android Pro - Ottimizzatore di Percorsi Completo
Versione: 3.0.0
Autore: Mattia Prosperi

Funzionalità:
- Caricamento file Excel con selezione colonne
- Multi-operatore con divisione intelligente percorsi
- Gestione indirizzi duplicati
- Mappa integrata (OpenStreetMap)
- Gestione percorsi salvati
- Export Excel completo con tutte le colonne originali
"""

import os
import json
import threading
import copy
from math import radians, sin, cos, atan2, sqrt, pi, log, tan, atan, sinh
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
from kivy.uix.stencilview import StencilView
from kivy.graphics import Color, Line, Ellipse, Rectangle

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Pandas NON disponibile su Android - usa solo openpyxl
PANDAS_AVAILABLE = False

# API Key OpenRouteService
ORS_API_KEY = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6IjQwNTQ3OTY0MjI0NjRmMzg4ZTFkNjQ1NTc4MGY4OGZkIiwiaCI6Im11cm11cjY0In0="

# Cache
DIST_CACHE = {}

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

def coords_are_same(coord1, coord2, tolerance=0.0001):
    """Verifica se due coordinate sono uguali (stesso indirizzo)"""
    return abs(coord1[0] - coord2[0]) < tolerance and abs(coord1[1] - coord2[1]) < tolerance

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

# ==================== MULTI-OPERATOR FUNCTIONS ====================

def calculate_centroid(coords):
    """Calcola il centroide di un gruppo di coordinate"""
    if not coords:
        return (0, 0)
    lat = sum(c[0] for c in coords) / len(coords)
    lon = sum(c[1] for c in coords) / len(coords)
    return (lat, lon)

def divide_for_operators(indices, coords, original_data, num_operators, items_per_operator, 
                         address_col=None, mode="haversine", callback=None):
    """
    Divide gli indici tra gli operatori in modo intelligente.
    
    - Ogni operatore riceve items_per_operator indirizzi
    - Nessun indirizzo duplicato tra operatori diversi
    - Se l'ultimo indirizzo ha duplicati, aggiungi fino a 20 indirizzi allo stesso operatore
    - Ottimizza per dare a ogni operatore il percorso più breve possibile
    """
    if num_operators <= 1:
        return [indices]
    
    n = len(indices)
    if n == 0:
        return [[] for _ in range(num_operators)]
    
    # Crea gruppi per indirizzo (coordinate simili)
    address_groups = {}  # coord_key -> [indices]
    
    for idx in indices:
        coord = coords[idx]
        # Arrotonda per raggruppare indirizzi vicini
        coord_key = (round(coord[0], 4), round(coord[1], 4))
        if coord_key not in address_groups:
            address_groups[coord_key] = []
        address_groups[coord_key].append(idx)
    
    # Converti in lista di gruppi
    groups = list(address_groups.values())
    
    # Calcola centroide di ogni gruppo
    group_centroids = []
    for group in groups:
        group_coords = [coords[i] for i in group]
        centroid = calculate_centroid(group_coords)
        group_centroids.append(centroid)
    
    # Dividi i gruppi tra operatori usando clustering geografico
    # Usa k-means semplificato basato sulla distanza dal centroide globale
    global_centroid = calculate_centroid([coords[i] for i in indices])
    
    # Ordina gruppi per distanza dal centroide globale
    group_distances = []
    for i, centroid in enumerate(group_centroids):
        dist = haversine(centroid[0], centroid[1], global_centroid[0], global_centroid[1])
        group_distances.append((dist, i, groups[i]))
    group_distances.sort()
    
    # Assegna gruppi agli operatori
    operator_assignments = [[] for _ in range(num_operators)]
    operator_counts = [0] * num_operators
    
    used_addresses = set()  # Traccia indirizzi già assegnati
    
    for dist, group_idx, group in group_distances:
        # Trova l'operatore con meno indirizzi che può ancora ricevere
        best_op = None
        min_count = float('inf')
        
        for op in range(num_operators):
            if operator_counts[op] < items_per_operator and operator_counts[op] < min_count:
                # Verifica che nessun indirizzo del gruppo sia già assegnato
                group_key = (round(coords[group[0]][0], 4), round(coords[group[0]][1], 4))
                if group_key not in used_addresses:
                    best_op = op
                    min_count = operator_counts[op]
        
        if best_op is None:
            # Tutti gli operatori sono pieni o indirizzo già usato
            # Cerca operatore con spazio per duplicati (fino a 20)
            for op in range(num_operators):
                if operator_counts[op] < items_per_operator + 10:  # Margine per duplicati
                    best_op = op
                    break
        
        if best_op is not None:
            # Assegna gruppo all'operatore
            for idx in group:
                if operator_counts[best_op] < items_per_operator:
                    operator_assignments[best_op].append(idx)
                    operator_counts[best_op] += 1
                elif operator_counts[best_op] < items_per_operator + 10:
                    # Gestione duplicati: aggiungi fino a 10 extra
                    # Verifica se è stesso indirizzo dell'ultimo
                    if operator_assignments[best_op]:
                        last_idx = operator_assignments[best_op][-1]
                        last_coord = coords[last_idx]
                        curr_coord = coords[idx]
                        if coords_are_same(last_coord, curr_coord):
                            operator_assignments[best_op].append(idx)
                            operator_counts[best_op] += 1
            
            # Marca indirizzo come usato
            group_key = (round(coords[group[0]][0], 4), round(coords[group[0]][1], 4))
            used_addresses.add(group_key)
    
    # Ottimizza ogni percorso operatore
    optimized_assignments = []
    for op_indices in operator_assignments:
        if len(op_indices) > 1:
            op_coords = [coords[i] for i in op_indices]
            tour, _, _ = solve_tsp(op_coords, mode)
            optimized = [op_indices[i] for i in tour]
            optimized_assignments.append(optimized)
        else:
            optimized_assignments.append(op_indices)
    
    return optimized_assignments

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
    """Legge un file Excel e restituisce dati e colonne (solo openpyxl)"""
    if not OPENPYXL_AVAILABLE:
        return None, "Libreria openpyxl non disponibile"
    
    try:
        if filepath.lower().endswith('.csv'):
            # Leggi CSV manualmente
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            if not lines:
                return None, "File vuoto"
            
            # Detect separator
            first_line = lines[0]
            if ';' in first_line:
                sep = ';'
            elif '\t' in first_line:
                sep = '\t'
            else:
                sep = ','
            
            headers = [h.strip().strip('"') for h in first_line.strip().split(sep)]
            rows = []
            for line in lines[1:]:
                if line.strip():
                    values = [v.strip().strip('"') for v in line.strip().split(sep)]
                    # Pad row if needed
                    while len(values) < len(headers):
                        values.append('')
                    rows.append(values[:len(headers)])
            
            return {'headers': headers, 'rows': rows, 'is_dict': True}, headers
        else:
            # Excel file
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            data = list(ws.values)
            if not data:
                return None, "File vuoto"
            
            headers = [str(h) if h is not None else f"Col{i}" for i, h in enumerate(data[0])]
            rows = []
            for row in data[1:]:
                rows.append(list(row))
            
            return {'headers': headers, 'rows': rows, 'is_dict': True}, headers
    except Exception as e:
        return None, str(e)

def export_to_excel_full(data_rows, original_columns, filepath, extra_columns=None):
    """Esporta dati in Excel con tutte le colonne"""
    if not OPENPYXL_AVAILABLE:
        return False, "Libreria openpyxl non disponibile"
    
    if extra_columns is None:
        extra_columns = ['Progressione', 'Operatore', 'Distanza_m', 'Distanza_Cumulata_m']
    
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Determina tutte le colonne
        all_cols = set()
        for row in data_rows:
            all_cols.update(row.keys())
        
        # Ordina: prima extra_columns, poi original_columns, poi il resto
        ordered_cols = []
        for c in extra_columns:
            if c in all_cols:
                ordered_cols.append(c)
                all_cols.discard(c)
        for c in original_columns:
            if c in all_cols:
                ordered_cols.append(c)
                all_cols.discard(c)
        ordered_cols.extend(sorted(all_cols))
        
        # Scrivi intestazioni
        ws.append(ordered_cols)
        
        # Scrivi dati
        for row in data_rows:
            row_data = []
            for col in ordered_cols:
                val = row.get(col, '')
                if val is None:
                    val = ''
                row_data.append(val)
            ws.append(row_data)
        
        wb.save(filepath)
        return True, filepath
    except Exception as e:
        return False, str(e)

# ==================== MAP WIDGET ====================

class MapWidget(StencilView):
    """Widget mappa semplificato"""
    
    lat = NumericProperty(41.9028)
    lon = NumericProperty(12.4964)
    zoom = NumericProperty(10)
    markers = ListProperty([])
    route_coords = ListProperty([])
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.bind(size=self.update_map)
        self.bind(pos=self.update_map)
        self.bind(lat=self.update_map)
        self.bind(lon=self.update_map)
        self.bind(zoom=self.update_map)
        self.bind(markers=self.update_map)
        self.bind(route_coords=self.update_map)
        Clock.schedule_once(self.update_map, 0.1)
        self.last_touch = None
    
    def lat_lon_to_pixel(self, lat, lon):
        """Converte coordinate in pixel"""
        try:
            scale = 256 * (2 ** self.zoom) / 360
            x = (lon - self.lon) * scale + self.width / 2 + self.x
            
            # Mercator projection
            lat_rad = radians(lat)
            center_lat_rad = radians(self.lat)
            y_scale = 256 * (2 ** self.zoom) / (2 * pi)
            
            y_target = log(tan(pi/4 + lat_rad/2))
            y_center = log(tan(pi/4 + center_lat_rad/2))
            
            y = (y_center - y_target) * y_scale + self.height / 2 + self.y
            
            return x, y
        except:
            return self.width / 2 + self.x, self.height / 2 + self.y
    
    def update_map(self, *args):
        """Aggiorna la visualizzazione della mappa"""
        self.canvas.clear()
        
        with self.canvas:
            # Sfondo
            Color(0.85, 0.9, 0.85, 1)
            Rectangle(pos=self.pos, size=self.size)
            
            # Griglia
            Color(0.75, 0.8, 0.75, 1)
            grid_size = 50
            for i in range(int(self.width / grid_size) + 1):
                x = self.x + i * grid_size
                Line(points=[x, self.y, x, self.y + self.height], width=1)
            for i in range(int(self.height / grid_size) + 1):
                y = self.y + i * grid_size
                Line(points=[self.x, y, self.x + self.width, y], width=1)
            
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
            for i, marker in enumerate(self.markers):
                if len(marker) >= 2:
                    lat, lon = marker[0], marker[1]
                    px, py = self.lat_lon_to_pixel(lat, lon)
                    
                    # Colore marker
                    if i == 0:
                        Color(0.2, 0.8, 0.2, 1)  # Verde partenza
                    elif i == len(self.markers) - 1:
                        Color(0.8, 0.2, 0.2, 1)  # Rosso arrivo
                    else:
                        Color(0.2, 0.4, 0.9, 1)  # Blu intermedi
                    
                    Ellipse(pos=(px-15, py-15), size=(30, 30))
                    
                    # Numero bianco
                    Color(1, 1, 1, 1)
                    Ellipse(pos=(px-10, py-10), size=(20, 20))
    
    def set_view(self, coords):
        """Centra la mappa sulle coordinate"""
        if not coords:
            return
        
        lats = [c[0] for c in coords]
        lons = [c[1] for c in coords]
        self.lat = sum(lats) / len(lats)
        self.lon = sum(lons) / len(lons)
        
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
        if touch.grab_current == self and self.last_touch:
            dx = touch.pos[0] - self.last_touch[0]
            dy = touch.pos[1] - self.last_touch[1]
            
            scale = 360 / (256 * (2 ** self.zoom))
            self.lon -= dx * scale
            self.lat += dy * scale * 0.7
            
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

# ==================== KV LANGUAGE ====================

KV = '''
<MenuButton@Button>:
    size_hint_y: None
    height: 55
    background_color: 0.2, 0.2, 0.25, 1
    background_normal: ''
    font_size: '16sp'
    bold: True

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
                text: 'PERCORSI PRO v3'
                font_size: '26sp'
                bold: True
        
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
                    on_release: 
                        app.root.current = 'saved_routes'
                        app._refresh_routes_list()
                
                MenuButton:
                    text: '⚙️  IMPOSTAZIONI'
                    background_color: 0.4, 0.4, 0.4, 1
                    on_release: app.root.current = 'settings'
                
                Widget:
                    size_hint_y: None
                    height: 20
                
                Label:
                    text: 'v3.0 - Multi-Operatore - Mattia Prosperi'
                    font_size: '12sp'
                    color: 0.5, 0.5, 0.5, 1
                    size_hint_y: None
                    height: 30

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
        
        FileChooserListView:
            id: file_chooser
            path: app.get_default_path()
            filters: ['*.xlsx', '*.xls', '*.csv']
            size_hint_y: 0.7
        
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
                    height: 180
                    background_color: 0.15, 0.15, 0.18, 1
                    foreground_color: 1, 1, 1, 1
                
                Label:
                    text: 'Etichette (opzionale):'
                    size_hint_y: None
                    height: 30
                
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
                    height: 280
                    
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
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        Label:
                            text: 'Indirizzo:'
                            size_hint_x: 0.4
                        Spinner:
                            id: address_spinner
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
                        text: 'OSRM/ORS: distanze stradali reali'
                        font_size: '12sp'
                        color: 0.6, 0.6, 0.6, 1
                        size_hint_y: None
                        height: 30
                
                CardBox:
                    size_hint_y: None
                    height: 220
                    
                    Label:
                        text: 'MULTI-OPERATORE'
                        font_size: '14sp'
                        bold: True
                        color: 0.5, 0.8, 1, 1
                        size_hint_y: None
                        height: 30
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        
                        CheckBox:
                            id: multi_operator_check
                            size_hint_x: 0.15
                            active: False
                        
                        Label:
                            text: 'Abilita divisione operatori'
                            halign: 'left'
                            text_size: self.size
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        Label:
                            text: 'Numero operatori:'
                            size_hint_x: 0.6
                        TextInput:
                            id: num_operators
                            text: '2'
                            multiline: False
                            input_filter: 'int'
                            size_hint_x: 0.4
                            background_color: 0.2, 0.2, 0.25, 1
                            foreground_color: 1, 1, 1, 1
                    
                    BoxLayout:
                        size_hint_y: None
                        height: 40
                        spacing: 10
                        Label:
                            text: 'Indirizzi per operatore:'
                            size_hint_x: 0.6
                        TextInput:
                            id: items_per_operator
                            text: '10'
                            multiline: False
                            input_filter: 'int'
                            size_hint_x: 0.4
                            background_color: 0.2, 0.2, 0.25, 1
                            foreground_color: 1, 1, 1, 1
                    
                    Label:
                        text: 'Duplicati: max +10 indirizzi extra'
                        font_size: '11sp'
                        color: 0.5, 0.5, 0.5, 1
                        size_hint_y: None
                        height: 25
                
                CardBox:
                    size_hint_y: None
                    height: 80
                    
                    Label:
                        text: 'API KEY ORS'
                        font_size: '14sp'
                        bold: True
                        color: 0.3, 0.8, 0.3, 1
                        size_hint_y: None
                        height: 30
                    
                    Label:
                        text: 'Preconfigurata ✓'
                        color: 0.3, 0.8, 0.3, 1
                        size_hint_y: None
                        height: 30

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
        
        ScrollView:
            BoxLayout:
                id: routes_list
                orientation: 'vertical'
                padding: 10
                spacing: 10
                size_hint_y: None
                height: self.minimum_height

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
                font_size: '14sp'
                bold: True
        
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
                font_size: '15sp'
                bold: True
        
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
                text: '💾 EXPORT'
                background_color: 0.6, 0.4, 0.2, 1
                background_normal: ''
                on_release: app.root.current = 'export'
        
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

<MapViewScreen>:
    name: 'map_view'
    BoxLayout:
        orientation: 'vertical'
        
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
        
        MapWidget:
            id: map_widget
        
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
                text: '➖'
                size_hint_x: 0.25
                background_color: 0.3, 0.3, 0.4, 1
                background_normal: ''
                on_release: map_widget.zoom_out()
            
            Button:
                text: '🎯 CENTRA'
                size_hint_x: 0.5
                background_color: 0.3, 0.5, 0.3, 1
                background_normal: ''
                on_release: app.center_map()
            
            Button:
                text: '➕'
                size_hint_x: 0.25
                background_color: 0.3, 0.3, 0.4, 1
                background_normal: ''
                on_release: map_widget.zoom_in()

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
                    height: 220
                    
                    Label:
                        text: 'FORMATO EXPORT'
                        font_size: '14sp'
                        bold: True
                        color: 1, 0.5, 0.2, 1
                        size_hint_y: None
                        height: 30
                    
                    Button:
                        text: '📊 EXCEL COMPLETO (.xlsx)'
                        size_hint_y: None
                        height: 50
                        background_color: 0.2, 0.6, 0.3, 1
                        background_normal: ''
                        on_release: app.export_excel()
                    
                    Button:
                        text: '📍 GPX (Navigatori)'
                        size_hint_y: None
                        height: 50
                        background_color: 0.3, 0.4, 0.6, 1
                        background_normal: ''
                        on_release: app.export_gpx()
                    
                    Button:
                        text: '🌍 KML (Google Earth)'
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
    original_rows_data = []  # Righe originali per export completo
    coords = ListProperty([])
    labels = ListProperty([])
    original_indices = ListProperty([])
    current_route = None
    saved_routes = ListProperty([])
    export_folder = StringProperty('')
    
    # Processing
    is_processing = BooleanProperty(False)
    
    def build(self):
        self.title = 'Percorsi Pro v3'
        if platform not in ('android', 'ios'):
            Window.size = (420, 780)
        self.load_saved_routes()
        return Builder.load_string(KV)
    
    def on_start(self):
        Clock.schedule_once(self._delayed_init, 1.0)
    
    def _delayed_init(self, dt):
        try:
            self._refresh_routes_list()
        except:
            pass
    
    def get_default_path(self):
        if platform == 'android':
            try:
                from android.storage import primary_external_storage_path
                return primary_external_storage_path()
            except:
                return '/storage/emulated/0'
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
    
    def get_multi_operator_settings(self):
        """Ritorna le impostazioni multi-operatore"""
        try:
            screen = self.root.get_screen('settings')
            enabled = screen.ids.multi_operator_check.active
            num_ops = int(screen.ids.num_operators.text or '2')
            items_per = int(screen.ids.items_per_operator.text or '10')
            return enabled, max(1, num_ops), max(1, items_per)
        except:
            return False, 2, 10
    
    # ==================== EXCEL LOADING ====================
    
    def load_excel(self, selection):
        if not selection:
            self.show_popup("Errore", "Seleziona un file!")
            return
        
        filepath = selection[0]
        data, columns = read_excel_file(filepath)
        
        if data is None:
            self.show_popup("Errore", f"Impossibile leggere:\n{columns}")
            return
        
        self.excel_data = data
        self.excel_columns = columns
        self.original_rows_data = []  # Reset
        
        self.root.current = 'column_select'
        Clock.schedule_once(lambda dt: self._populate_spinners(columns), 0.2)
    
    def _populate_spinners(self, columns):
        try:
            screen = self.root.get_screen('column_select')
            col_options = ['-- Nessuna --'] + columns
            
            screen.ids.lat_spinner.values = columns
            screen.ids.lon_spinner.values = columns
            screen.ids.date_spinner.values = col_options
            screen.ids.operator_spinner.values = col_options
            screen.ids.label_spinner.values = col_options
            screen.ids.address_spinner.values = col_options
            
            # Auto-detect
            for col in columns:
                u = col.upper()
                if 'LAT' in u:
                    screen.ids.lat_spinner.text = col
                elif 'LON' in u:
                    screen.ids.lon_spinner.text = col
                elif 'DATA' in u or 'DATE' in u:
                    screen.ids.date_spinner.text = col
                elif 'OPERATORE' in u or 'OPERATOR' in u:
                    screen.ids.operator_spinner.text = col
                elif 'NOME' in u or 'NAME' in u or 'LABEL' in u or 'DESC' in u or 'ETICHETTA' in u:
                    screen.ids.label_spinner.text = col
                elif 'INDIRIZZO' in u or 'ADDRESS' in u or 'VIA' in u:
                    screen.ids.address_spinner.text = col
        except Exception as e:
            print(f"Spinner error: {e}")
    
    def start_excel_optimization(self):
        screen = self.root.get_screen('column_select')
        lat_col = screen.ids.lat_spinner.text
        lon_col = screen.ids.lon_spinner.text
        
        if lat_col == 'Seleziona...' or lon_col == 'Seleziona...':
            self.show_popup("Errore", "Seleziona Latitudine e Longitudine!")
            return
        
        self.coords = []
        self.labels = []
        self.original_indices = []
        self.original_rows_data = []  # Salva righe originali
        
        label_col = screen.ids.label_spinner.text
        address_col = screen.ids.address_spinner.text
        
        # Usa formato dizionario (openpyxl)
        if self.excel_data and 'headers' in self.excel_data:
            headers = self.excel_data['headers']
            rows = self.excel_data['rows']
            
            try:
                lat_idx = headers.index(lat_col)
                lon_idx = headers.index(lon_col)
            except ValueError:
                self.show_popup("Errore", "Colonne non trovate!")
                return
            
            label_idx = headers.index(label_col) if label_col in headers else -1
            address_idx = headers.index(address_col) if address_col in headers else -1
            
            for row_idx, row in enumerate(rows):
                try:
                    lat_val = row[lat_idx] if lat_idx < len(row) else None
                    lon_val = row[lon_idx] if lon_idx < len(row) else None
                    
                    if lat_val is None or lon_val is None:
                        continue
                    
                    # Converti in float
                    lat = float(str(lat_val).replace(',', '.'))
                    lon = float(str(lon_val).replace(',', '.'))
                    
                    if -90 <= lat <= 90 and -180 <= lon <= 180:
                        self.coords.append((lat, lon))
                        self.original_indices.append(row_idx)
                        
                        # Salva riga originale come dizionario
                        row_dict = {}
                        for i, h in enumerate(headers):
                            if i < len(row):
                                row_dict[h] = row[i] if row[i] is not None else ''
                            else:
                                row_dict[h] = ''
                        self.original_rows_data.append(row_dict)
                        
                        # Label
                        label_parts = []
                        if label_idx >= 0 and label_idx < len(row) and row[label_idx]:
                            label_parts.append(str(row[label_idx]))
                        if address_idx >= 0 and address_idx < len(row) and row[address_idx]:
                            label_parts.append(str(row[address_idx]))
                        
                        self.labels.append(' - '.join(label_parts) if label_parts else f"Punto {len(self.coords)}")
                except Exception as e:
                    print(f"Errore riga {row_idx}: {e}")
                    continue
        
        if len(self.coords) < 2:
            self.show_popup("Errore", f"Solo {len(self.coords)} coordinate valide!")
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
        self.original_indices = list(range(len(self.coords)))
        self.original_rows_data = []  # No original data for manual input
        
        while len(self.labels) < len(self.coords):
            self.labels.append(f"Punto {len(self.labels) + 1}")
        
        if len(self.coords) < 2:
            self.show_popup("Errore", "Minimo 2 coordinate!")
            return
        
        self._start_optimization()
    
    # ==================== OPTIMIZATION ====================
    
    def _start_optimization(self):
        self.root.current = 'processing'
        self.is_processing = True
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
            multi_enabled, num_ops, items_per = self.get_multi_operator_settings()
            
            Clock.schedule_once(lambda dt: self._set_status("Calcolo distanze..."), 0)
            
            all_indices = list(range(len(self.coords)))
            
            # Prepara dati originali se disponibili
            orig_rows = getattr(self, 'original_rows_data', [])
            
            if multi_enabled and num_ops > 1:
                # Multi-operatore
                Clock.schedule_once(lambda dt: self._set_status(f"Divisione per {num_ops} operatori..."), 0)
                
                operator_groups = divide_for_operators(
                    all_indices, 
                    list(self.coords), 
                    None,
                    num_ops, 
                    items_per,
                    mode=mode
                )
                
                routes = []
                for op_idx, op_indices in enumerate(operator_groups):
                    if not op_indices:
                        continue
                    
                    op_coords = [self.coords[i] for i in op_indices]
                    op_labels = [self.labels[i] for i in op_indices]
                    op_orig_indices = [self.original_indices[i] if i < len(self.original_indices) else i for i in op_indices]
                    op_orig_rows = [orig_rows[i] if i < len(orig_rows) else {} for i in op_indices]
                    
                    if len(op_coords) > 1:
                        tour, total, distances = solve_tsp(op_coords, mode)
                        opt_coords = [op_coords[i] for i in tour]
                        opt_labels = [op_labels[i] for i in tour]
                        opt_orig = [op_orig_indices[i] for i in tour]
                        opt_rows = [op_orig_rows[i] for i in tour]
                    else:
                        opt_coords = op_coords
                        opt_labels = op_labels
                        opt_orig = op_orig_indices
                        opt_rows = op_orig_rows
                        total = 0
                        distances = [0]
                    
                    route = {
                        'id': f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{op_idx+1}",
                        'name': f"Operatore {op_idx + 1} - {datetime.now().strftime('%d/%m %H:%M')}",
                        'operator': f"Operatore {op_idx + 1}",
                        'coords': opt_coords,
                        'labels': opt_labels,
                        'original_indices': opt_orig,
                        'original_row_data': opt_rows,
                        'distances': distances,
                        'total_distance': total,
                        'created': datetime.now().isoformat(),
                        'stops_count': len(opt_coords)
                    }
                    routes.append(route)
                
                Clock.schedule_once(lambda dt: self._multi_optimization_done(routes), 0)
            else:
                # Singolo percorso
                tour, total, distances = solve_tsp(list(self.coords), mode, update_progress)
                
                opt_coords = [self.coords[i] for i in tour]
                opt_labels = [self.labels[i] for i in tour]
                opt_orig = [self.original_indices[i] if i < len(self.original_indices) else i for i in tour]
                opt_rows = [orig_rows[i] if i < len(orig_rows) else {} for i in tour]
                
                route = {
                    'id': datetime.now().strftime('%Y%m%d_%H%M%S'),
                    'name': f"Percorso {datetime.now().strftime('%d/%m %H:%M')}",
                    'operator': 'Unico',
                    'coords': opt_coords,
                    'labels': opt_labels,
                    'original_indices': opt_orig,
                    'original_row_data': opt_rows,
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
        self.saved_routes.append(route)
        self.save_routes_to_file()
        self._show_route_detail(route)
    
    def _multi_optimization_done(self, routes):
        self.is_processing = False
        for route in routes:
            self.saved_routes.append(route)
        self.save_routes_to_file()
        
        if routes:
            self.current_route = routes[0]
            self.show_popup("Completato", f"Creati {len(routes)} percorsi operatore!")
            self._show_route_detail(routes[0])
        else:
            self.root.current = 'home'
    
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
            filepath = os.path.join(self.get_default_path(), '.percorsi_routes_v3.json')
            if os.path.exists(filepath):
                with open(filepath, 'r') as f:
                    self.saved_routes = json.load(f)
        except:
            self.saved_routes = []
    
    def save_routes_to_file(self):
        try:
            filepath = os.path.join(self.get_default_path(), '.percorsi_routes_v3.json')
            with open(filepath, 'w') as f:
                json.dump(list(self.saved_routes), f)
        except:
            pass
    
    def _refresh_routes_list(self):
        try:
            screen = self.root.get_screen('saved_routes')
            routes_list = screen.ids.routes_list
        except:
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
            op_text = f"[{route.get('operator', 'N/A')}] " if route.get('operator') else ""
            btn = Button(
                text=f"📍 {op_text}{route['name']}\n   {self._format_distance(route['total_distance'])} | {route['stops_count']} tappe",
                size_hint_y=None,
                height=70,
                background_color=(0.2, 0.25, 0.3, 1),
                background_normal='',
                halign='left',
                valign='middle'
            )
            btn.bind(on_release=lambda x, r=route: self._show_route_detail(r))
            routes_list.add_widget(btn)
    
    def clear_all_routes(self):
        self.saved_routes = []
        self.save_routes_to_file()
        self._refresh_routes_list()
        self.show_popup("Info", "Percorsi eliminati")
    
    # ==================== ROUTE DETAIL ====================
    
    def _show_route_detail(self, route):
        self.current_route = route
        screen = self.root.get_screen('route_detail')
        
        op_text = f"[{route.get('operator', '')}] " if route.get('operator') else ""
        screen.ids.route_title.text = f"{op_text}{route['name']}"
        screen.ids.route_summary.text = f"Distanza: {self._format_distance(route['total_distance'])} | Tappe: {route['stops_count']}"
        
        stops_list = screen.ids.stops_list
        stops_list.clear_widgets()
        
        cumulative = 0
        for i, (coord, label) in enumerate(zip(route['coords'], route['labels'])):
            dist = route['distances'][i] if i < len(route['distances']) else 0
            cumulative += dist
            
            stop_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=50, padding=5)
            
            num_label = Label(text=str(i + 1), size_hint_x=0.1, bold=True, color=(1, 0.5, 0.2, 1))
            info_label = Label(text=f"{label[:30]}", size_hint_x=0.6, halign='left', font_size='13sp')
            dist_label = Label(text=self._format_distance(cumulative), size_hint_x=0.3, color=(0.5, 0.8, 0.5, 1), font_size='12sp')
            
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
        
        markers = [(c[0], c[1], l) for c, l in zip(self.current_route['coords'], self.current_route['labels'])]
        map_widget.markers = markers
        map_widget.route_coords = self.current_route['coords']
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
                self.show_popup("Info", f"Percorso diviso in {len(links)} segmenti")
    
    def open_google_maps(self):
        self.start_navigation()
    
    # ==================== EXPORT ====================
    
    def choose_export_folder(self):
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        file_chooser = FileChooserListView(path=self.get_default_path(), dirselect=True)
        content.add_widget(file_chooser)
        
        btn_box = BoxLayout(size_hint_y=None, height=50, spacing=10)
        
        def select_folder(instance):
            self.export_folder = file_chooser.selection[0] if file_chooser.selection else file_chooser.path
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
            try:
                from android.storage import primary_external_storage_path
                return os.path.join(primary_external_storage_path(), 'Download', filename)
            except:
                pass
        return os.path.join(os.path.expanduser('~'), filename)
    
    def export_excel(self):
        """Esporta in Excel con TUTTE le colonne originali"""
        if not self.current_route:
            return
        
        data_rows = []
        cumulative = 0
        
        for i, (coord, label) in enumerate(zip(self.current_route['coords'], self.current_route['labels'])):
            dist = self.current_route['distances'][i] if i < len(self.current_route['distances']) else 0
            cumulative += dist
            
            row = {
                'Progressione': i + 1,
                'Operatore': self.current_route.get('operator', ''),
                'Distanza_m': dist,
                'Distanza_Cumulata_m': cumulative
            }
            
            # Aggiungi dati originali se disponibili
            orig_data = self.current_route.get('original_row_data', [])
            if i < len(orig_data) and orig_data[i]:
                for key, val in orig_data[i].items():
                    if key not in row:
                        row[key] = val if val is not None else ''
            
            # Fallback
            if 'Latitudine' not in row:
                row['Latitudine'] = coord[0]
                row['Longitudine'] = coord[1]
                row['Etichetta'] = label
            
            data_rows.append(row)
        
        filepath = self._get_export_path(f"percorso_{self.current_route['id']}.xlsx")
        original_cols = self.excel_columns if self.excel_columns else []
        success, result = export_to_excel_full(data_rows, original_cols, filepath)
        
        if success:
            self.show_popup("Esportato", f"Salvato:\n{result}")
        else:
            self.show_popup("Errore", f"{result}")
    
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
        
        kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document><n>Percorso</n>']
        for i, (coord, label) in enumerate(zip(self.current_route['coords'], self.current_route['labels'])):
            kml.append(f'<Placemark><n>{label}</n><Point><coordinates>{coord[1]},{coord[0]},0</coordinates></Point></Placemark>')
        coords_str = ' '.join([f"{c[1]},{c[0]},0" for c in self.current_route['coords']])
        kml.append(f'<Placemark><n>Traccia</n><LineString><coordinates>{coords_str}</coordinates></LineString></Placemark>')
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
