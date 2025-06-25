import os
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output
import dash_bootstrap_components as dbc
import calendar

# === Load & Prepare Data ===
def load_and_prepare_data(filepath, label):
    df = pd.read_excel(filepath)
    df['APPLICATION DATE'] = pd.to_datetime(df['APPLICATION DATE'], errors='coerce')
    df = df.dropna(subset=['APPLICATION DATE', 'REGION', 'Status'])
    df = df[df['Status'].str.strip().str.lower() == 'paid']
    df['DAY_OF_MONTH'] = df['APPLICATION DATE'].dt.day
    df['YEAR'] = df['APPLICATION DATE'].dt.year
    df['MONTH_NUM'] = df['APPLICATION DATE'].dt.month
    df['SOURCE'] = label
    return df

# === Load Your Excel Files ===
df_2024 = load_and_prepare_data('10.xlsx', '2024')
df_2025 = load_and_prepare_data('univeristy - 24-6.xlsx', '2025')

df_combined = pd.concat([df_2024, df_2025])
regions = sorted(df_combined['REGION'].unique())
months = [{'label': calendar.month_name[m], 'value': m} for m in range(1, 13)]

# === Dash App Init ===
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server  # ✅ Expose for Gunicorn

# === Layout ===
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

# === Callback ===
@app.callback(
    Output('line-chart', 'figure'),
    Input('region-filter', 'value'),
    Input('month-filter', 'value'),
)
def update_chart(selected_region, selected_month):
    dff = df_combined.copy()
    
    if selected_region:
        dff = dff[dff['REGION'] == selected_region]
    if selected_month:
        dff = dff[dff['MONTH_NUM'] == selected_month]
    else:
        # Prevent crashing if no month selected
        return px.line(title="Please select a month to display data.")

    # Create 'Year Month' label for color
    dff['YEAR_MONTH'] = dff.apply(
        lambda row: f"{calendar.month_name[row['MONTH_NUM']]} {row['YEAR']}", axis=1
    )

    # Group and count
    daily_group = (
        dff.groupby(['DAY_OF_MONTH', 'YEAR_MONTH'])
        .size()
        .reset_index(name='TOTAL_PAID')
    )

    # Handle number of days for month (safe default for February)
    max_days = calendar.monthrange(2024, selected_month)[1]

    # Plotly Express Line Chart
    fig = px.line(
        daily_group,
        x='DAY_OF_MONTH',
        y='TOTAL_PAID',
        color='YEAR_MONTH',
        markers=True,
        labels={
            'DAY_OF_MONTH': 'Day of Month',
            'TOTAL_PAID': 'Total Paid Applications',
            'YEAR_MONTH': 'Year & Month'
        },
        title=f"Daily Paid Applications Comparison: {calendar.month_name[selected_month]}"
    )

    fig.update_layout(
        xaxis=dict(tickmode='linear', dtick=1, range=[1, max_days], title='Day of Month'),
        yaxis_title="Total Paid Applications",
        plot_bgcolor="white",
        font=dict(size=14),
        title_font_size=20,
        hovermode="x unified",
        legend_title_text='Year and Month'
    )

    # Save a local interactive HTML
    fig.write_html("paid_applications_comparison.html", include_plotlyjs='cdn')

    return fig

# === Run Server ===
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8050))  # for Render dynamic port
    app.run_server(debug=False, host='0.0.0.0', port=port)
