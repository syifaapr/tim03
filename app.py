# -*- coding: utf-8 -*-
"""
Dashboard Pelatihan Realtime - Kementerian Keuangan
Dengan Auto-Refresh dari Google Drive + Backup Lokal
"""

import dash
from dash import dcc, html, Input, Output, State
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import calendar
import io
import requests
import os
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# URL Google Drive file (view-only)
GOOGLE_DRIVE_URL = "https://drive.google.com/file/d/1TBu_i7ZNxFRUddJkbndNVjjygaOu7PP8/view?usp=sharing"

# GOOGLE_SHEET_ID = "1zUTMfqT99sAjqtBvgXKWNXzzkq8atcJc"  

# Setting
USE_GOOGLE_DRIVE = True  
REFRESH_INTERVAL = 5 * 60 * 1000 

# Nama file backup lokal
BACKUP_FILE = "kalpem_backup.xlsx"
CSV_BACKUP = "kalpem_backup.csv"

def download_from_google_drive():
    """Download Excel dari Google Drive (public link)"""
    try:
        logger.info(f"[{datetime.now().strftime('%H:%M:%S')}] Mengambil data dari Google Drive...")
        
        # Download file from gd
        response = requests.get(GOOGLE_DRIVE_URL, timeout=30)
        response.raise_for_status() 
        
        # Baca ke BytesIO
        content = io.BytesIO(response.content)
        
        # Parse Excel
        df = pd.read_excel(content, engine='openpyxl')
        
        logger.info(f"Berhasil! {len(df)} baris data")
        
        # Simpan backup lokal
        df.to_excel(BACKUP_FILE, index=False)
        df.to_csv(CSV_BACKUP, index=False)
        
        return df, True
        
    except Exception as e:
        print(f"Error Google Drive: {e}")
        print(" Menggunakan backup lokal...")
        
        try:
            # Coba baca dari Excel backup
            if os.path.exists(BACKUP_FILE):
                df = pd.read_excel(BACKUP_FILE, engine='openpyxl')
                print(f"Menggunakan backup Excel: {len(df)} baris")
                return df, False
            elif os.path.exists(CSV_BACKUP):
                df = pd.read_csv(CSV_BACKUP)
                print(f"Menggunakan backup CSV: {len(df)} baris")
                return df, False
            else:
                # if no backup, gunakan file default
                df = pd.read_csv("kalpem.csv", sep=",")
                print(f"⚠️ Menggunakan file default: {len(df)} baris")
                return df, False
                
        except Exception as backup_error:
            print(f"Error backup: {backup_error}")
            return pd.DataFrame(), False

def parse_indonesian_date(date_str):
    """Konversi tanggal Indonesia ke datetime"""
    if pd.isna(date_str):
        return pd.NaT
    
    month_map = {
        'Januari': 'January', 'Februari': 'February', 'Maret': 'March',
        'April': 'April', 'Mei': 'May', 'Juni': 'June',
        'Juli': 'July', 'Agustus': 'August', 'September': 'September',
        'Oktober': 'October', 'November': 'November', 'Desember': 'December'
    }
    
    try:
        parts = str(date_str).split()
        if len(parts) >= 3:
            day = parts[0]
            month_ind = parts[1]
            year = parts[2]
            
            if month_ind in month_map:
                month_en = month_map[month_ind]
                date_en = f"{day} {month_en} {year}"
                return pd.to_datetime(date_en, format='%d %B %Y')
    except:
        pass
    
    return pd.to_datetime(date_str, errors='coerce')

def process_data(df):
    """Process dan cleaning data"""
    if df.empty:
        return df
    
    # Buat copy untuk menghindai SettingWithCopyWarning
    df = df.copy()
    
    # Clean data
    df = df.dropna(how='all')
    
    # Hapus kolom yang tidak perlu
    cols_to_drop = ['No.', 'Unnamed: 0', 'Unnamed: 0.1']
    for col in cols_to_drop:
        if col in df.columns:
            df = df.drop(columns=[col])
    
    # Konversi tanggal
    df['Mulai'] = df['Mulai'].apply(parse_indonesian_date)
    df['Akhir'] = df['Akhir'].apply(parse_indonesian_date)
    
    # Ekstrak informasi
    df['Bulan_Num'] = df['Mulai'].dt.month
    df['Bulan_Indo'] = df['Mulai'].dt.month.apply(
        lambda x: {
            1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
            5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
            9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
        }.get(x, 'Tidak Diketahui') if pd.notna(x) else 'Tidak Diketahui'
    )
    
    df['Tahun'] = df['Mulai'].dt.year
    df['Durasi'] = (df['Akhir'] - df['Mulai']).dt.days + 1
    
    # Pastikan kolom numerik bertipe numerik
    numeric_cols = ['TotalPeserta', 'TotalJamlator']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            df[col] = df[col].fillna(0)
    
    return df

print("Memuat data...")

if USE_GOOGLE_DRIVE:
    df, connection_ok = download_from_google_drive()
else:
    df = pd.read_csv("kalpem.csv", sep=",")
    connection_ok = False

df = process_data(df)

print(f"Data berhasil dimuat: {len(df)} baris")

# Urutan bulan
bulan_urutan = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
]

# Data statistik
total_pelatihan_all = len(df)
total_peserta_all = int(df['TotalPeserta'].sum()) if 'TotalPeserta' in df.columns else 0
total_jamlator_all = int(df['TotalJamlator'].sum()) if 'TotalJamlator' in df.columns else 0

app = dash.Dash(__name__, suppress_callback_exceptions=True)

server = app.server

app.layout = html.Div([
    # Background 
    html.Div(className="dashboard-background"),
    
    # Theme storage
    dcc.Store(id='theme-store', data='light'),
    
    # Data storage 
    dcc.Store(id='data-store', data=df.to_dict('records')),
    
    # Download store
    dcc.Download(id="download-excel"),
    
    # Interval
    dcc.Interval(
        id='interval-refresh',
        interval=REFRESH_INTERVAL,
        n_intervals=0
    ),

    # Main wrapper
    html.Div(id='dashboard-wrapper', children=[
        html.Div([
            # Sidebar Header
            html.Div([
                html.H1("TIM AP 03"),
                html.P("Kementerian Keuangan RI", className="subtitle"),
                html.P("Dashboard Pelatihan", 
                       style={"fontSize": "0.85rem", "color": "rgba(255,255,255,0.7)", "marginTop": "5px"}),
                
                # Status Koneksi & Update
                html.Div([
                    html.Div([
                        html.Span(id="connection-status", style={"fontSize": "0.8rem", "fontWeight": "500"}),
                        html.Span(" | ", style={"margin": "0 5px", "color": "rgba(255,255,255,0.)"}),
                        html.Span(id="last-update", style={"fontSize": "0.8rem"})
                    ])
                ], style={"marginTop": "10px", "padding": "8px", 
                         "backgroundColor": "rgba(255,255,255,0.1)", 
                         "borderRadius": "6px", "textAlign": "center"})
            ], className="sidebar-header"),
            
            # Filter Content
            html.Div([
                # Filter Bulan
                html.Div([
                    html.Div([
                        html.Span(" ", style={"fontSize": "1.1rem"}),
                        html.H4("FILTER BULAN", style={"margin": "0"})
                    ], className="section-title"),
                    html.Div([
                        html.Label("Pilih Bulan:", className="filter-label"),
                        dcc.Dropdown(
                            id="bulan-dropdown",
                            options=[{"label": b, "value": b} for b in bulan_urutan],
                            placeholder="Semua bulan...",
                            multi=True,
                            style={"color": "#fffff"}
                        )
                    ])
                ], className="sidebar-section"),
                
                # Filter Penyelenggara
                html.Div([
                    html.Div([
                        html.Span(" ", style={"fontSize": "1.1rem"}),
                        html.H4("PENYELENGGARA", style={"margin": "0"})
                    ], className="section-title"),
                    html.Div([
                        html.Label("Pilih Unit Kerja:", className="filter-label"),
                        dcc.Dropdown(
                            id="penyelenggara-dropdown",
                            options=[{"label": str(p), "value": str(p)} 
                                     for p in sorted([str(x) for x in df['Penyelenggara'].dropna().unique()])],
                            placeholder="Semua penyelenggara...",
                            multi=True
                        )
                    ])
                ], className="sidebar-section"),
                
                # Filter Metode
                html.Div([
                    html.Div([
                        html.Span("", style={"fontSize": "1.1rem"}),
                        html.H4("METODE PELATIHAN", style={"margin": "0"})
                    ], className="section-title"),
                    html.Div([
                        html.Label("Pilih Metode:", className="filter-label"),
                        dcc.Dropdown(
                            id="metode-dropdown",
                            options=[{"label": str(m), "value": str(m)} 
                                     for m in sorted([str(x) for x in df['Metode'].dropna().unique()])],
                            placeholder="Semua metode...",
                            multi=True
                        )
                    ])
                ], className="sidebar-section"),
                
            ], className="sidebar-content"),
            
            # Sidebar Footer
            html.Div([
                # Info Refresh
                html.Div([
                    html.Span(" Auto-refresh: ", style={"fontSize": "0.75rem"}),
                    html.Span(f"setiap {REFRESH_INTERVAL//60000} menit", 
                             style={"fontSize": "0.75rem", "fontWeight": "500"})
                ], style={"textAlign": "center", "marginBottom": "10px", 
                         "color": "rgba(255,255,255,0.6)"}),
                
                # Manual Refresh Button
                html.Button([
                    html.Span(" "),
                    " Refresh Data Sekarang"
                ], id="manual-refresh-btn", n_clicks=0, 
                   style={"width": "100%", "padding": "10px", "marginBottom": "10px",
                         "backgroundColor": "#57c5b6", "color": "white", "border": "none",
                         "borderRadius": "8px", "cursor": "pointer", "fontWeight": "500",
                         "transition": "all 0.3s"}),
                
                # Theme Toggle
                html.Div([
                    html.Div([
                        html.Span(" ", className="toggle-icon"),
                        html.Span("Mode Gelap", style={"flex": "1", "textAlign": "center"})
                    ], style={"display": "flex", "alignItems": "center", "gap": "10px"}),
                    html.Div(id="theme-switch", 
                            style={"width": "40px", "height": "20px", 
                                   "backgroundColor": "#ccc", "borderRadius": "10px",
                                   "position": "relative", "cursor": "pointer"})
                ], className="theme-toggle", id="theme-toggle"),
                
                # Reset Button
                html.Button([
                    html.Span(" "),
                    " Reset Semua Filter"
                ], id="reset-btn", n_clicks=0, className="reset-btn")
            ], className="sidebar-footer")
            
        ], className="sidebar"),
        
        html.Div([
            # Top Bar
            html.Div([
                html.Div([
                    html.H2("Dashboard Monitoring Pelatihan", style={"marginBottom": "5px"}),
                    html.Div([
                        html.Span(f"📚 {total_pelatihan_all} Pelatihan • "),
                        html.Span(f"👥 {total_peserta_all:,} Peserta • "),
                        html.Span(f"⏱️ {total_jamlator_all:,} Jam Lator")
                    ], className="date-info")
                ]),
                html.Div([
                    html.Div([
                        html.Div(id="filtered-count", className="stat-value"),
                        html.Div("Data Tersaring", className="stat-label")
                    ], className="stat-item"),
                    html.Div([
                        html.Div(id="avg-peserta", className="stat-value"),
                        html.Div("Rata-rata Peserta", className="stat-label")
                    ], className="stat-item"),
                ], className="stats-info")
            ], className="top-bar"),
            
            # KPI Cards
            html.Div([
                html.Div([
                    html.Div([
                        html.Div("📚", className="kpi-icon"),
                        html.H3(id="total-pelatihan"),
                        html.P("Total Pelatihan"),
                        html.Div([
                            html.Span("📈", className="trend-up"),
                            html.Span("+12% vs bulan lalu", style={"marginLeft": "5px"})
                        ], className="kpi-trend")
                    ], className="kpi-card kpi-1"),
                    
                    html.Div([
                        html.Div("👥", className="kpi-icon"),
                        html.H3(id="total-peserta"),
                        html.P("Total Peserta"),
                        html.Div([
                            html.Span("📈", className="trend-up"),
                            html.Span("+8% vs bulan lalu", style={"marginLeft": "5px"})
                        ], className="kpi-trend")
                    ], className="kpi-card kpi-2"),
                    
                    html.Div([
                        html.Div("💻", className="kpi-icon"),
                        html.H3(id="e-learning-count"),
                        html.P("E-Learning"),
                        html.Div([
                            html.Span("📈", className="trend-up"),
                            html.Span("+25% vs bulan lalu", style={"marginLeft": "5px"})
                        ], className="kpi-trend")
                    ], className="kpi-card kpi-3"),
                    
                    html.Div([
                        html.Div("🌐", className="kpi-icon"),
                        html.H3(id="pjj-count"),
                        html.P("PJJ"),
                        html.Div([
                            html.Span("📈", className="trend-up"),
                            html.Span("+15% vs bulan lalu", style={"marginLeft": "5px"})
                        ], className="kpi-trend")
                    ], className="kpi-card kpi-4"),
                    
                    html.Div([
                        html.Div("⏱️", className="kpi-icon"),
                        html.H3(id="total-jamlator-count"),
                        html.P("Total Jam Lator"),
                        html.Div([
                            html.Span("📈", className="trend-up"),
                            html.Span("+10% vs bulan lalu", style={"marginLeft": "5px"})
                        ], className="kpi-trend")
                    ], className="kpi-card kpi-5")
                ], className="kpi-grid")
            ]),
            
            # Charts Section
            html.Div([
                html.Div([
                    html.H2("📊 Analisis Data", style={"flex": "1"}),
                    html.Div([
                        html.Button("Hari Ini", className="badge badge-primary"),
                        html.Button("Bulan Ini", className="badge badge-success"),
                        html.Button("Tahun Ini", className="badge badge-warning")
                    ], className="chart-actions")
                ], className="section-header"),
                
                html.Div([
                    html.Div([
                        html.Div([
                            html.H3("Pelatihan per Bulan"),
                            html.Span("Tren Bulanan", style={"fontSize": "0.9rem", "color": "var(--text-light)"})
                        ], className="chart-header"),
                        dcc.Graph(id="pelatihan-chart", config={'displayModeBar': True})
                    ], className="chart-container"),
                    
                    html.Div([
                        html.Div([
                            html.H3("Distribusi Metode"),
                            html.Span(" Komposisi", style={"fontSize": "0.9rem", "color": "var(--text-light)"})
                        ], className="chart-header"),
                        dcc.Graph(id="metode-chart")
                    ], className="chart-container"),
                    
                    html.Div([
                        html.Div([
                            html.H3("Top 10 Penyelenggara"),
                            html.Span(" Unit Teraktif", style={"fontSize": "0.9rem", "color": "var(--text-light)"})
                        ], className="chart-header"),
                        dcc.Graph(id="penyelenggara-chart")
                    ], className="chart-container"),
                    
                    html.Div([
                        html.Div([
                            html.H3("Level Evaluasi"),
                            html.Span("Tingkat Evaluasi", style={"fontSize": "0.9rem", "color": "var(--text-light)"})
                        ], className="chart-header"),
                        dcc.Graph(id="level-chart")
                    ], className="chart-container")
                ], className="charts-grid")
            ], className="charts-section"),
            
            # Data Table Section
            html.Div([
                html.Div([
                    html.H2("📋 Data Detail Pelatihan", style={"flex": "1"}),
                    html.Div([
                        html.Button("📥 Ekspor ke Excel", 
                                    id="export-excel-btn",
                                    n_clicks=0,
                                  style={"backgroundColor": "#1a5f7a", "color": "white",
                                         "border": "none", "padding": "10px 20px", "borderRadius": "8px",
                                         "cursor": "pointer", "fontWeight": "500", "marginRight": "10px"}),
                        html.Div(id="export-status", 
                                style={"fontSize": "0.9rem", "color": "green", 
                                       "display": "inline-block", "verticalAlign": "middle"})
                    ], style={"display": "flex", "alignItems": "center", "gap": "10px"})
                ], className="table-header"),
                html.Div(id="data-table", className="table-container")
            ], className="data-section")
            
        ], className="main-content")
    ], className="dashboard-wrapper")
])

@app.callback(
    [Output('data-store', 'data'),
     Output('last-update', 'children'),
     Output('connection-status', 'children'),
     Output('connection-status', 'style')],
    [Input('interval-refresh', 'n_intervals'),
     Input('manual-refresh-btn', 'n_clicks')]
)
def refresh_data(n_intervals, manual_clicks):
    """Auto-refresh data dari Google Drive"""
    try:
        if USE_GOOGLE_DRIVE:
            new_df, is_connected = download_from_google_drive()
        else:
            new_df = pd.read_csv("kalpem.csv", sep=",")
            is_connected = False
        
        # Process data
        new_df = process_data(new_df)
        
        # Status
        timestamp = datetime.now().strftime('%H:%M:%S')
        
        if is_connected:
            status_text = "Online"
            status_style = {"color": "#57c5b6", "fontWeight": "500"}
        else:
            status_text = "Offline (Use Backup File)"
            status_style = {"color": "#f39c12", "fontWeight": "500"}
        
        return (
            new_df.to_dict('records'), 
            f"Update: {timestamp}", 
            status_text,
            status_style
        )
        
    except Exception as e:
        print(f"Error refresh: {e}")
        
        # Jika error, gunakan data yang ada
        return (
            df.to_dict('records'), 
            f"⚠️ Error: {str(e)[:30]}", 
            "Error",
            {"color": "#e74c3c", "fontWeight": "500"}
        )
    
@app.callback(
    [Output('theme-store', 'data'),
     Output('theme-switch', 'style')],
    Input('theme-toggle', 'n_clicks'),
    State('theme-store', 'data'),
    prevent_initial_call=True
)
def toggle_theme(n_clicks, current_theme):
    if n_clicks is None:
        return current_theme, {}
    
    new_theme = 'dark' if current_theme == 'light' else 'light'
    switch_style = {
        'width': '40px',
        'height': '20px',
        'backgroundColor': '#57c5b6' if new_theme == 'dark' else '#ccc',
        'borderRadius': '10px',
        'position': 'relative',
        'cursor': 'pointer'
    }
    
    return new_theme, switch_style

@app.callback(
    [
        Output("total-pelatihan", "children"),
        Output("total-peserta", "children"),
        Output("e-learning-count", "children"),
        Output("pjj-count", "children"),
        Output("total-jamlator-count", "children"),
        Output("filtered-count", "children"),
        Output("avg-peserta", "children"),
        Output("pelatihan-chart", "figure"),
        Output("metode-chart", "figure"),
        Output("penyelenggara-chart", "figure"),
        Output("level-chart", "figure"),
        Output("data-table", "children")
    ],
    [
        Input("data-store", "data"),
        Input("bulan-dropdown", "value"),
        Input("penyelenggara-dropdown", "value"),
        Input("metode-dropdown", "value")
    ]
)
def update_dashboard(data_dict, bulan_filter, penyelenggara_filter, metode_filter):
    # Convert dari store ke DataFrame
    filtered_df = pd.DataFrame(data_dict)
    
    # Apply filters
    if bulan_filter:
        filtered_df = filtered_df[filtered_df['Bulan_Indo'].isin(bulan_filter)]
    
    if penyelenggara_filter:
        filtered_df = filtered_df[filtered_df['Penyelenggara'].astype(str).isin(penyelenggara_filter)]
    
    if metode_filter:
        filtered_df = filtered_df[filtered_df['Metode'].astype(str).isin(metode_filter)]
    
    # KPI Calculations
    total_pelatihan = len(filtered_df)
    
    if 'TotalPeserta' in filtered_df.columns:
        filtered_df['TotalPeserta'] = pd.to_numeric(filtered_df['TotalPeserta'], errors='coerce')
        total_peserta = filtered_df['TotalPeserta'].sum()
        avg_peserta = filtered_df['TotalPeserta'].mean() if len(filtered_df) > 0 else 0
    else:
        total_peserta = 0
        avg_peserta = 0
    
    if 'TotalJamlator' in filtered_df.columns:
        filtered_df['TotalJamlator'] = pd.to_numeric(filtered_df['TotalJamlator'], errors='coerce')
        total_jamlator = filtered_df['TotalJamlator'].sum()
    else:
        total_jamlator = 0
    
    e_learning_count = len(filtered_df[filtered_df['Metode'].astype(str) == 'E-Learning'])
    pjj_count = len(filtered_df[filtered_df['Metode'].astype(str) == 'PJJ'])
    
    # Chart 1: Pelatihan per Bulan
    if not filtered_df.empty:
        bulan_counts = filtered_df['Bulan_Indo'].value_counts().reindex(bulan_urutan, fill_value=0).reset_index()
        bulan_counts.columns = ['Bulan', 'Jumlah']
        
        fig1 = px.bar(
            bulan_counts,
            x='Bulan',
            y='Jumlah',
            title='',
            labels={'Bulan': 'Bulan', 'Jumlah': 'Jumlah Pelatihan'},
            color='Jumlah',
            color_continuous_scale='Viridis'
        )
        fig1.update_layout(
            height=350,
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
    else:
        fig1 = px.bar(title='')
        fig1.update_layout(height=350)
    
    # Chart 2: Distribusi Metode
    if not filtered_df.empty:
        metode_counts = filtered_df['Metode'].value_counts().reset_index()
        metode_counts.columns = ['Metode', 'Jumlah']
        
        fig2 = px.pie(
            metode_counts,
            values='Jumlah',
            names='Metode',
            title='',
            hole=0.4,
            color_discrete_sequence=['#1a5f7a', '#57c5b6', '#159895']
        )
        fig2.update_layout(height=350, showlegend=True)
    else:
        fig2 = px.pie(title='')
        fig2.update_layout(height=350)
    
    # Chart 3: Top Penyelenggara
    if not filtered_df.empty:
        penyelenggara_counts = filtered_df['Penyelenggara'].value_counts().head(10).reset_index()
        penyelenggara_counts.columns = ['Penyelenggara', 'Jumlah']
        
        fig3 = px.bar(
            penyelenggara_counts,
            x='Penyelenggara',
            y='Jumlah',
            title='',
            labels={'Penyelenggara': 'Penyelenggara', 'Jumlah': 'Jumlah Pelatihan'},
            color='Jumlah',
            color_continuous_scale='Blues'
        )
        fig3.update_layout(height=350, xaxis_tickangle=-45, showlegend=False)
    else:
        fig3 = px.bar(title='')
        fig3.update_layout(height=350)
    
    # Chart 4: Level Evaluasi
    if not filtered_df.empty and 'LevelEvaluasi' in filtered_df.columns:
        level_counts = filtered_df['LevelEvaluasi'].value_counts().reset_index()
        level_counts.columns = ['Level', 'Jumlah']
        
        fig4 = px.bar(
            level_counts,
            x='Level',
            y='Jumlah',
            title='',
            labels={'Level': 'Level Evaluasi', 'Jumlah': 'Jumlah'},
            color='Level',
            color_discrete_sequence=['#1a5f7a', '#57c5b6', '#159895', '#9b59b6']
        )
        fig4.update_layout(height=350, showlegend=False)
    else:
        fig4 = px.bar(title='')
        fig4.update_layout(height=350)
    
    # Data Table
    if not filtered_df.empty:
        columns_to_show = [
            'NamaProgramPembelajaran', 'Mulai', 'Akhir', 
            'Metode', 'Penyelenggara', 'TotalPeserta','Jumlahkelas', 'Bulan_Indo', 'TotalJamlator'
        ]
        
        available_cols = [col for col in columns_to_show if col in filtered_df.columns]
        table_data = filtered_df[available_cols].head(15)
        
        # Format table
        table_header = [html.Tr([html.Th(col.replace('_', ' ')) for col in available_cols])]
        table_rows = []
        
        for _, row in table_data.iterrows():
            row_cells = []
            for col in available_cols:
                value = row[col]
                if pd.isna(value):
                    value = ''
                elif 'Mulai' in col or 'Akhir' in col:
                    if pd.notna(value):
                        try:
                            value = pd.to_datetime(value).strftime('%d %b %Y')
                        except:
                            value = str(value)
                    else:
                        value = ''
                elif 'TotalPeserta' in col or 'TotalJamlator' in col:
                    try:
                        value = f"{int(float(value)):,}"
                    except:
                        value = str(value)
                else:
                    value = str(value)
                
                cell_style = {
                    "padding": "12px 15px",
                    "borderBottom": "1px solid var(--border-color)",
                    "color": "var(--text-color)"
                }
                row_cells.append(html.Td(value, style=cell_style))
            
            table_rows.append(html.Tr(row_cells))
        
        table = html.Table(
            [html.Thead(table_header), html.Tbody(table_rows)],
            className="data-table"
        )
    else:
        table = html.Div(
            "Tidak ada data yang sesuai dengan filter",
            style={"textAlign": "center", "padding": "40px", "color": "var(--text-light)"}
        )
    
    return (
        f"{total_pelatihan}",
        f"{int(total_peserta):,}",
        f"{e_learning_count}",
        f"{pjj_count}",
        f"{int(total_jamlator):,}",
        f"{total_pelatihan}",
        f"{int(avg_peserta)}",
        fig1, fig2, fig3, fig4, table
    )

@app.callback(
    [
        Output("bulan-dropdown", "value"),
        Output("penyelenggara-dropdown", "value"),
        Output("metode-dropdown", "value")
    ],
    Input("reset-btn", "n_clicks")
)
def reset_filters(n_clicks):
    if n_clicks:
        return [], [], []
    return dash.no_update, dash.no_update, dash.no_update

@app.callback(
    Output("download-excel", "data"),
    Input("export-excel-btn", "n_clicks"),
    [State("data-store", "data"),
     State("bulan-dropdown", "value"),
     State("penyelenggara-dropdown", "value"),
     State("metode-dropdown", "value")],
    prevent_initial_call=True
)
def export_to_excel(n_clicks, data_dict, bulan_filter, penyelenggara_filter, metode_filter):
    """Ekspor data ke file Excel"""
    if n_clicks is None or n_clicks == 0:
        return dash.no_update
    
    try:
        print(f"Memulai ekspor Excel... (klik ke-{n_clicks})")
        
        # Konversi data
        if not data_dict:
            print("Data kosong")
            return None
        
        df = pd.DataFrame(data_dict)
        
        if df.empty:
            print("DataFrame kosong")
            return None
        
        # Apply filter with same like dash
        if bulan_filter:
            df = df[df['Bulan_Indo'].isin(bulan_filter)]
        
        if penyelenggara_filter:
            df = df[df['Penyelenggara'].astype(str).isin(penyelenggara_filter)]
        
        if metode_filter:
            df = df[df['Metode'].astype(str).isin(metode_filter)]
        
        print(f"📊 Data untuk ekspor: {len(df)} baris")
        
        # 3. Pilih kolom untuk ekspor
        column_order = [
            'NamaProgramPembelajaran', 'Mulai', 'Akhir', 'Durasi',
            'Metode', 'Penyelenggara', 'TotalPeserta', 'Jumlahkelas',
            'Bulan_Indo', 'TotalJamlator', 'Tahun'
        ]
        
        # Hanya ambil kolom yang ada
        existing_cols = [col for col in column_order if col in df.columns]
        df_export = df[existing_cols].copy()
        
        # Format data
        # Tanggal
        for col in ['Mulai', 'Akhir']:
            if col in df_export.columns:
                df_export[col] = pd.to_datetime(df_export[col], errors='coerce')
                df_export[col] = df_export[col].dt.strftime('%d/%m/%Y')
        
        # Numerik
        for col in ['TotalPeserta', 'TotalJamlator']:
            if col in df_export.columns:
                df_export[col] = pd.to_numeric(df_export[col], errors='coerce')
                df_export[col] = df_export[col].fillna(0).astype(int)
                # Format dengan pemisah ribuan
                df_export[col] = df_export[col].apply(lambda x: f"{x:,}")
        
        # Buat Excel di memory
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet 1: Data Pelatihan
            df_export.to_excel(writer, sheet_name='Data Pelatihan', index=False)
            
            # Sheet 2: Ringkasan
            summary_data = {
                'Metrik': [
                    'Total Pelatihan',
                    'Total Peserta', 
                    'Total Jam Lator',
                    'Rata-rata Peserta per Pelatihan',
                    'Pelatihan Terbanyak di Bulan'
                ],
                'Nilai': [
                    len(df_export),
                    df['TotalPeserta'].sum() if 'TotalPeserta' in df.columns else 0,
                    df['TotalJamlator'].sum() if 'TotalJamlator' in df.columns else 0,
                    round(df['TotalPeserta'].mean(), 1) if 'TotalPeserta' in df.columns else 0,
                    df['Bulan_Indo'].mode()[0] if 'Bulan_Indo' in df.columns and not df.empty else '-'
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Ringkasan', index=False)
            
            # Sheet 3: Statistik per Penyedia
            if 'Penyelenggara' in df.columns:
                penyedia_stats = df.groupby('Penyelenggara').agg({
                    'NamaProgramPembelajaran': 'count',
                    'TotalPeserta': 'sum',
                    'TotalJamlator': 'sum'
                }).reset_index()
                penyedia_stats.columns = ['Penyelenggara', 'Jumlah Pelatihan', 'Total Peserta', 'Total Jam Lator']
                pd.DataFrame(penyedia_stats).to_excel(writer, sheet_name='Statistik Penyedia', index=False)
        
        output.seek(0)
        
        # 6. Nama file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"Dashboard_Pelatihan_{timestamp}.xlsx"
        
        print(f"Ekspor berhasil: {filename}")
        
        return dcc.send_bytes(
            output.getvalue(),
            filename=filename,
            type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        print(f"Error saat ekspor: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

@app.callback(
    Output("export-status", "children"),
    Input("export-excel-btn", "n_clicks"),
    State("data-store", "data"),
    prevent_initial_call=True
)
def show_export_status(n_clicks, data_dict):
    """Tampilkan status ekspor"""
    if n_clicks and n_clicks > 0:
        if data_dict:
            df = pd.DataFrame(data_dict)
            return f"Sedang mengekspor {len(df)} data..."
    return ""

if __name__ == "__main__":
    print("\n" + "="*60)
    print("DASHBOARD TIM AP 03 - KEMENTERIAN KEUANGAN RI")
    print("="*60)
    print(f"Total Data: {total_pelatihan_all:,} pelatihan")
    print(f"Total Peserta: {total_peserta_all:,} orang")
    print(f"Total Jam Pelatihan: {total_jamlator_all:,} jam")
    print("="*60)
    if USE_GOOGLE_DRIVE:
        print(f"Auto-Refresh: AKTIF (setiap {REFRESH_INTERVAL//60000} menit)")
        print(f" Sumber: Google Drive")
        print(f" URL: {GOOGLE_DRIVE_URL}")
    else:
        print("Sumber: File Lokal (kalpem.csv)")
    print(f"Backup: {BACKUP_FILE}, {CSV_BACKUP}")
    print(" Fitur: Mode Gelap/Terang • Ekspor Excel • Filter")
    print("="*60)
    print(" Buka browser dan akses: http://localhost:8050")
    print("="*60)
    
    app.run_server(debug=True, port=8050, host='0.0.0.0')