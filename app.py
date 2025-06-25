import os
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output
import dash_bootstrap_components as dbc
import calendar

# === Load & Prepare Data ===
def load_and_prepare_data(filepath, label):
    """Loads and prepares data from an Excel file."""
    df = pd.read_excel(filepath)
    # ✅ FIX: Added dayfirst=True to correctly parse DD.MM.YYYY formats.
    df['APPLICATION DATE'] = pd.to_datetime(df['APPLICATION DATE'], errors='coerce', dayfirst=True)
    
    # Ensure required columns exist before proceeding
    required_cols = ['APPLICATION DATE', 'REGION', 'Status']
    if not all(col in df.columns for col in required_cols):
        # Raise an error if a column is missing, which will be caught below
        raise ValueError(f"File {filepath} is missing one of the required columns: {required_cols}")

    df = df.dropna(subset=['APPLICATION DATE', 'REGION', 'Status'])
    df = df[df['Status'].str.strip().str.lower() == 'paid']
    df['DAY_OF_MONTH'] = df['APPLICATION DATE'].dt.day
    df['YEAR'] = df['APPLICATION DATE'].dt.year
    df['MONTH_NUM'] = df['APPLICATION DATE'].dt.month
    df['SOURCE'] = label
    return df

# === Dash App Init ===
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server  # Expose for Gunicorn

# === Robust Data Loading ===
try:
    # Attempt to load and process the data
    df_2024 = load_and_prepare_data('10.xlsx', '2024')
    df_2025 = load_and_prepare_data('univeristy - 24-6.xlsx', '2025')
    df_combined = pd.concat([df_2024, df_2025])
    
    if df_combined.empty:
        # Handle case where files load but have no 'paid' applications
        raise ValueError("Data loaded, but no rows with status 'paid' were found.")

    regions = sorted(df_combined['REGION'].unique())
    months = [{'label': calendar.month_name[m], 'value': m} for m in range(1, 13)]

    # === Main App Layout ===
    app.layout = dbc.Container([
        html.H2("📈 Compare Paid Applications by Day: 2024 vs 2025"),
        dbc.Row([
            dbc.Col([
                html.Label("Select Region:"),
                dcc.Dropdown(
                    id='region-filter',
                    options=[{'label': r, 'value': r} for r in regions],
                    placeholder="Choose a region"
                )
            ]),
            dbc.Col([
                html.Label("Select Month:"),
                dcc.Dropdown(
                    id='month-filter',
                    options=months,
                    placeholder="Choose a month"
                )
            ])
        ]),
        dcc.Graph(id='line-chart')
    ], fluid=True)

except Exception as e:
    # ✅ ROBUSTNESS: If data loading fails, display an error message instead of crashing.
    app.layout = dbc.Container([
        html.H2("Error Loading Application Data"),
        html.Hr(),
        html.P("There was an error initializing the application. Please check the data files and column names."),
        # Display the actual error for debugging
        html.P(f"Error details: {e}"),
    ])


# === Callback ===
# This callback will only be active if the main layout was created successfully
if 'df_combined' in locals():
    @app.callback(
        Output('line-chart', 'figure'),
        Input('region-filter', 'value'),
        Input('month-filter', 'value'),
    )
    def update_chart(selected_region, selected_month):
        dff = df_combined.copy()
        
        if not selected_month:
            # Prevent update if no month is selected
            return px.line(title="Please select a month to display the chart.")

        if selected_region:
            dff = dff[dff['REGION'] == selected_region]
        
        dff = dff[dff['MONTH_NUM'] == selected_month]

        if dff.empty:
            return px.line(title=f"No 'paid' data found for the selected filters in {calendar.month_name[selected_month]}.")

        # Group and count
        daily_group = (
            dff.groupby(['DAY_OF_MONTH', 'YEAR'])
            .size()
            .reset_index(name='TOTAL_PAID')
        )

        # Handle number of days for the month
        # Use a non-leap year like 2023 for monthrange as a safe default
        max_days = calendar.monthrange(2023, selected_month)[1]

        # Plotly Express Line Chart
        fig = px.line(
            daily_group,
            x='DAY_OF_MONTH',
            y='TOTAL_PAID',
            color='YEAR', # Simpler to color by year directly
            markers=True,
            labels={
                'DAY_OF_MONTH': f'Day of {calendar.month_name[selected_month]}',
                'TOTAL_PAID': 'Total Paid Applications',
                'YEAR': 'Year'
            },
            title=f"Daily Paid Applications: {calendar.month_name[selected_month]}"
        )

        fig.update_layout(
            xaxis=dict(tickmode='linear', dtick=1, range=[1, max_days]),
            plot_bgcolor="white",
            hovermode="x unified",
        )
        
        # This is for local debugging, you may want to remove it for production
        # fig.write_html("paid_applications_comparison.html")

        return fig

# === Run Server ===
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8050))
    app.run_server(debug=True, host='0.0.0.0', port=port)
