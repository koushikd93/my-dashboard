import os
import glob
import pandas as pd
from datetime import datetime, time
from dash import Dash, dcc, html, Input, Output
import plotly.express as px
from dash import dash_table
from dash import ctx
import io
import base64



# === Load and preprocess Excel files ===
def load_all_excels(folder_path):
    all_files = glob.glob(folder_path + "/*.xlsx")
    valid_files = [file for file in all_files if not os.path.basename(file).startswith('~$')]
    df_list = []
    for file in valid_files:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()  # Clean column names
        if not df.dropna(how='all').empty:
            df_list.append(df)
    return pd.concat(df_list, ignore_index=True)

# === Constants ===
folder_path = "data"
df_raw = load_all_excels(folder_path)


# === Rename column ===
df_raw.rename(columns={'LPG . No.': 'Chainage'}, inplace=True)

# === Preprocessing ===
df_raw['Alert Time'] = pd.to_datetime(df_raw['Alert Time'], errors='coerce', dayfirst=True)
df_raw['Date'] = df_raw['Alert Time'].dt.date
df_raw['Hour'] = df_raw['Alert Time'].dt.hour

def classify_alarm_type(t):
    if pd.isnull(t):
        return 'Unknown'
    if time(6, 0) <= t.time() <= time(22, 0):
        return 'Day'
    return 'Night'

df_raw['Alarm Type'] = df_raw['Alert Time'].apply(classify_alarm_type)
df_raw['Verified'] = df_raw[['Resolution remarks with time', 'Resolution remarks with time/Day Guard']].notnull().any(axis=1)

# === Unique Values ===
sections = df_raw['Section'].dropna().unique()

# === Dash App ===
app = Dash(__name__)
app.title = "Alarm Dashboard"

app.layout = html.Div(style={'fontFamily': 'Arial', 'backgroundColor': '#f8f9fa', 'padding': '20px'}, children=[
    html.H1("🚨 PIDWS Alarm Dashboard", style={"textAlign": "center", "color": "#343a40"}),

    html.Div([
        html.Div([
            html.Label("🔔 Select Alarm Type:", style={'fontWeight': 'bold'}),
            dcc.Dropdown(['Day', 'Night'], 'Day', id='alarm-type', style={'width': '200px'})
        ]),
        html.Div([
            html.Label("📅 Select Date Range:", style={'fontWeight': 'bold'}),
            dcc.DatePickerRange(
                id='date-range',
                start_date=df_raw['Date'].min(),
                end_date=df_raw['Date'].max(),
                display_format='DD-MM-YYYY',
                style={'border': '1px solid #ced4da', 'padding': '5px'}
            )
        ]),
        html.Div([
            html.Label("📍 Select Section:", style={'fontWeight': 'bold'}),
            dcc.Dropdown(list(sections), sections[0], id='section', style={'width': '250px'})
        ])
    ], style={
        'display': 'flex',
        'justifyContent': 'center',
        'gap': '30px',
        'flexWrap': 'wrap',
        'marginBottom': '30px'
    }),

    html.Div([
        html.H2("1️⃣ Number of Alarms", style={'color': '#007bff'}),
        dcc.Graph(id='alarm-count-graph')
    ], style={'width': '100%', 'padding': '10px'}),

    html.Div([
        html.H2("2️⃣ Verified vs Unverified Alarms", style={'color': '#17a2b8'}),
        dcc.Graph(id='verify-status-graph')
    ], style={'width': '100%', 'padding': '10px'}),

    html.Div([
        html.H2("3️⃣ List of Unverified Alarm Chainages", style={'color': '#dc3545'}),
        dash_table.DataTable(id='unverified-table',
            columns=[
                {'name': 'Section', 'id': 'Section'},
                {'name': 'Chainage', 'id': 'Chainage'},
                {'name': 'Alert Type/Severity', 'id': 'Alert Type/Severity'},
                {'name': 'Alert Time', 'id': 'Alert Time'},
                {'name': 'Alert Duration(HH:MM:SS)', 'id': 'Alert Duration(HH:MM:SS)'},
                {'name': 'Event Type', 'id': 'Event Type'},
                {'name': 'Location', 'id': 'Location', 'presentation': 'markdown'}
            ],
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'left', 'padding': '8px'},
            style_header={'backgroundColor': '#007bff', 'color': 'white', 'fontWeight': 'bold'},
            style_data_conditional=[{
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f2f2f2'
            }],
            page_size=10
        ),

        html.Br(),
        html.Button("⬇️ Download as Excel", id="download-excel-btn", n_clicks=0,
                    style={'backgroundColor': '#28a745', 'color': 'white', 'border': 'none',
                           'padding': '10px 20px', 'borderRadius': '5px', 'cursor': 'pointer'}),
        dcc.Download(id="download-excel")
    ], style={'width': '100%', 'padding': '10px'})
])



# === Callbacks ===


@app.callback(
    Output('alarm-count-graph', 'figure'),
    Input('alarm-type', 'value'),
    Input('date-range', 'start_date'),
    Input('date-range', 'end_date'),
    Input('section', 'value')
)
def update_alarm_count_graph(alarm_type, start_date, end_date, section):
    filtered = df_raw[(df_raw['Alarm Type'] == alarm_type) &
                      (df_raw['Date'] >= pd.to_datetime(start_date).date()) &
                      (df_raw['Date'] <= pd.to_datetime(end_date).date()) &
                      (df_raw['Section'] == section)]

    if start_date == end_date:
        fig = px.histogram(
            filtered,
            x='Hour',
            nbins=24,
            title="Alarms Distribution by Hour",
            labels={'Hour': 'Hour of Day', 'count': 'Number of Alarms'}
        )
    else:
        fig = px.histogram(
            filtered,
            x='Date',
            title="Alarms Distribution by Date",
            labels={'Date': 'Date', 'count': 'Number of Alarms'}
        )

    fig.update_layout(bargap=0.2, xaxis_tickformat="%d-%m-%Y" if start_date != end_date else None)
    return fig

@app.callback(
    Output('verify-status-graph', 'figure'),
    Input('alarm-type', 'value'),
    Input('date-range', 'start_date'),
    Input('date-range', 'end_date'),
    Input('section', 'value')
)
def update_verify_graph(alarm_type, start_date, end_date, section):
    filtered = df_raw[(df_raw['Alarm Type'] == alarm_type) &
                      (df_raw['Date'] >= pd.to_datetime(start_date).date()) &
                      (df_raw['Date'] <= pd.to_datetime(end_date).date()) &
                      (df_raw['Section'] == section)]

    summary = filtered.groupby('Date').agg(
        total_alarms=('Alert Time', 'count'),
        verified=('Verified', lambda x: x.sum()),
        unverified=('Verified', lambda x: (~x).sum())
    ).reset_index()

    fig = px.bar(summary, x='Date', y=['total_alarms', 'verified', 'unverified'],
                 title="Alarm Verification Status",
                 labels={'value': 'Count', 'variable': 'Status'},
                 color_discrete_map={
                     'total_alarms': 'yellow',
                     'verified': 'blue',
                     'unverified': 'red'
                 })
    

    
    return fig


@app.callback(
    Output('unverified-table', 'data'),
    Input('alarm-type', 'value'),
    Input('date-range', 'start_date'),
    Input('date-range', 'end_date'),
    Input('section', 'value')
)
def update_unverified_table(alarm_type, start_date, end_date, section):
    filtered = df_raw[(df_raw['Alarm Type'] == alarm_type) &
                      (df_raw['Date'] >= pd.to_datetime(start_date).date()) &
                      (df_raw['Date'] <= pd.to_datetime(end_date).date()) &
                      (df_raw['Section'] == section) &
                      (~df_raw['Verified'])]

    # Drop rows without coordinates
    filtered = filtered.dropna(subset=['Chainage', 'Latitude', 'Longitude'])

    # Create Google Maps link as Markdown
    filtered['Location'] = filtered.apply(
      lambda row: f"[📍](https://www.google.com/maps?q={row['Latitude']},{row['Longitude']}&t=k)",
      axis=1
    )

    return filtered[[
        'Section','Chainage', 'Alert Type/Severity', 'Alert Time',
        'Alert Duration(HH:MM:SS)', 'Event Type', 'Location'
    ]].to_dict('records')

@app.callback(
    Output("download-excel", "data"),
    Input("download-excel-btn", "n_clicks"),
    Input('alarm-type', 'value'),
    Input('date-range', 'start_date'),
    Input('date-range', 'end_date'),
    Input('section', 'value'),
    prevent_initial_call=True
)
def download_unverified_excel(n_clicks, alarm_type, start_date, end_date, section):
    if ctx.triggered_id != 'download-excel-btn':
        return None

    filtered = df_raw[(df_raw['Alarm Type'] == alarm_type) &
                      (df_raw['Date'] >= pd.to_datetime(start_date).date()) &
                      (df_raw['Date'] <= pd.to_datetime(end_date).date()) &
                      (df_raw['Section'] == section) &
                      (~df_raw['Verified'])]

    filtered = filtered.dropna(subset=['Chainage', 'Latitude', 'Longitude'])

    # Reorganize columns and remove hyperlink formatting
    download_df = filtered[[
        'Section','Chainage', 'Alert Type/Severity', 'Alert Time',
        'Alert Duration(HH:MM:SS)', 'Event Type', 'Latitude', 'Longitude'
    ]]

    # Convert to Excel in-memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        download_df.to_excel(writer, index=False, sheet_name='Unverified Alarms')
    output.seek(0)


    section_clean = section.replace(" ", "_")
    start_str = pd.to_datetime(start_date).strftime('%d-%m-%Y')
    end_str = pd.to_datetime(end_date).strftime('%d-%m-%Y')
    filename = f"{section_clean}_{start_str}_to_{end_str}_unverified_{alarm_type}_alarms.xlsx"

    return dcc.send_bytes(output.read(), filename=filename)


if __name__ == '__main__':
    app.run(debug=True)
