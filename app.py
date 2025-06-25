import os
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output
import dash_bootstrap_components as dbc
import calendar

# === Load & Prepare Data ===
def load_and_prepare_data(filepath, label):
    """Load and prepare data from Excel file with error handling"""
    try:
        df = pd.read_excel(filepath)
        df['APPLICATION DATE'] = pd.to_datetime(
            df['APPLICATION DATE'], 
            format='%d.%m.%Y %H:%M:%S', 
            dayfirst=True,
            errors='coerce'
        )
        df = df.dropna(subset=['APPLICATION DATE', 'REGION', 'Status'])
        df = df[df['Status'].str.strip().str.lower() == 'paid']

        df['DAY_OF_MONTH'] = df['APPLICATION DATE'].dt.day
        df['YEAR'] = df['APPLICATION DATE'].dt.year
        df['MONTH_NUM'] = df['APPLICATION DATE'].dt.month
        df['SOURCE'] = label

        print(f"Successfully loaded {len(df)} records from {filepath}")
        return df

    except FileNotFoundError:
        print(f"Warning: File {filepath} not found. Creating empty DataFrame.")
        return pd.DataFrame(columns=[
            'APPLICATION DATE', 'REGION', 'Status', 
            'DAY_OF_MONTH', 'YEAR', 'MONTH_NUM', 'SOURCE'
        ])
    except Exception as e:
        print(f"Error loading {filepath}: {str(e)}")
        return pd.DataFrame(columns=[
            'APPLICATION DATE', 'REGION', 'Status', 
            'DAY_OF_MONTH', 'YEAR', 'MONTH_NUM', 'SOURCE'
        ])

# === Load Your Excel Files ===
print("Loading data files...")
df_2024 = load_and_prepare_data('10.xlsx', '2024')
df_2025 = load_and_prepare_data('univeristy - 24-6.xlsx', '2025')

df_combined = pd.concat([df_2024, df_2025], ignore_index=True)

if df_combined.empty:
    print("Warning: No data loaded. Using placeholder data.")
    regions = ['No Data Available']
else:
    regions = sorted(df_combined['REGION'].unique())
    print(f"Data loaded successfully. Found {len(regions)} regions.")

months = [{'label': calendar.month_name[m], 'value': m} for m in range(1, 13)]

# === Dash App Initialization ===
app = Dash(
    __name__, 
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True
)

# Expose Flask server for Gunicorn
server = app.server  # Ensure this line is present

app.title = "Application Analytics Dashboard"

# === Layout ===
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            html.H1("📈 Application Analytics Dashboard", className="text-center mb-4"),
            html.P(
                "Compare paid applications by day across different regions and time periods",
                className="text-center text-muted mb-4"
            )
        ])
    ]),
    dbc.Card([
        dbc.CardBody([
            dbc.Row([
                dbc.Col([
                    html.Label("Select Region:", className="fw-bold"),
                    dcc.Dropdown(
                        id='region-filter',
                        options=[{'label': r, 'value': r} for r in regions],
                        placeholder="Choose a region...",
                        className="mb-3"
                    )
                ], md=6),
                dbc.Col([
                    html.Label("Select Month:", className="fw-bold"),
                    dcc.Dropdown(
                        id='month-filter',
                        options=months,
                        placeholder="Choose a month...",
                        className="mb-3"
                    )
                ], md=6)
            ])
        ])
    ], className="mb-4"),
    dbc.Card([
        dbc.CardBody([
            dcc.Graph(
                id='line-chart',
                style={'height': '600px'}
            )
        ])
    ]),
    html.Hr(),
    html.P(
        "Dashboard showing paid applications comparison between 2024 and 2025",
        className="text-center text-muted small mt-3"
    )
], fluid=True, className="py-4")

# === Callback ===
@app.callback(
    Output('line-chart', 'figure'),
    Input('region-filter', 'value'),
    Input('month-filter', 'value'),
)
def update_chart(selected_region, selected_month):
    if df_combined.empty:
        fig = px.line(title="No data available. Please check your Excel files.")
        fig.update_layout(
            plot_bgcolor="white",
            font=dict(size=14),
            title_font_size=18,
            height=500
        )
        return fig
    
    dff = df_combined.copy()

    if selected_region:
        dff = dff[dff['REGION'] == selected_region]

    if selected_month:
        dff = dff[dff['MONTH_NUM'] == selected_month]
    else:
        fig = px.line(title="Please select a month to display data.")
        fig.update_layout(
            plot_bgcolor="white",
            font=dict(size=14),
            title_font_size=18,
            height=500
        )
        return fig

    if dff.empty:
        title = f"No data available for the selected filters"
        if selected_region and selected_month:
            title += f": {selected_region} in {calendar.month_name[selected_month]}"
        fig = px.line(title=title)
        fig.update_layout(
            plot_bgcolor="white",
            font=dict(size=14),
            title_font_size=18,
            height=500
        )
        return fig

    dff['YEAR_MONTH'] = dff.apply(
        lambda row: f"{calendar.month_name[row['MONTH_NUM']]} {row['YEAR']}",
        axis=1
    )

    daily_group = (
        dff.groupby(['DAY_OF_MONTH', 'YEAR_MONTH'])
        .size()
        .reset_index(name='TOTAL_PAID')
    )

    max_days = calendar.monthrange(2024, selected_month)[1]

    fig = px.line(
        daily_group,
        x='DAY_OF_MONTH',
        y='TOTAL_PAID',
        color='YEAR_MONTH',
        markers=True,
        labels={
            'DAY_OF_MONTH': 'Day of Month',
            'TOTAL_PAID': 'Total Paid Applications',
            'YEAR_MONTH': 'Period'
        },
        title=f"Daily Paid Applications: {calendar.month_name[selected_month]}" +
              (f" - {selected_region}" if selected_region else "")
    )

    fig.update_layout(
        xaxis=dict(
            tickmode='linear',
            dtick=1,
            range=[1, max_days],
            title='Day of Month'
        ),
        yaxis_title="Total Paid Applications",
        plot_bgcolor="white",
        font=dict(size=12),
        title_font_size=16,
        hovermode="x unified",
        legend_title_text='Period',
        height=500,
        margin=dict(l=50, r=50, t=80, b=50)
    )
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')

    return fig

# === Application Entry Point ===
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8050))
    print(f"Starting server on port {port}")
    app.run_server(
        debug=False,
        host='0.0.0.0',
        port=port
    )
