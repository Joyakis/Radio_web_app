import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px

# Load Excel file
file_path = "Q1 '2024 -  Q1'2025.xlsx"
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names

# Filter out 'Reach' and 'Share' from the sheet names
sheet_names_filtered = [name for name in sheet_names if name not in ['Reach', 'Share']]

# Load all sheets into a dictionary
dataframes = {sheet: xls.parse(sheet) for sheet in sheet_names_filtered}

# Initialize Dash app
app = dash.Dash(__name__)
app.title = "Excel Sheet Visualizer"

# Layout
app.layout = html.Div([
    html.H1("Radio Stations Quarterly Performance Dashboard"),
    dcc.Dropdown(
        id='sheet-selector',
        options=[{"label": name, "value": name} for name in sheet_names_filtered],
        value=sheet_names_filtered[0]  # Set the default value to the first sheet
    ),
    dcc.Graph(id='sheet-graph')
])

# Callback for updating graph
@app.callback(
    Output('sheet-graph', 'figure'),
    [Input('sheet-selector', 'value')]
)
def update_graph(sheet_name):
    df = dataframes[sheet_name]
    if df.shape[1] >= 6:
        df_subset = df.iloc[:20, :6]
        df_melted = df_subset.melt(id_vars=df.columns[0], var_name="Category", value_name="Value")

        # Convert percentage values to decimals (if they aren't already)
        #df_melted["Value"] = df_melted["Value"] / 100

        fig = px.bar(
            df_melted,
            x=df.columns[0],
            y="Value",
            color="Category",
            barmode="group",
            title=sheet_name
        )

        # Format y-axis ticks as percentages
        fig.update_layout(
            yaxis_tickformat=".2%"  # Show values like 32.65%
        )
    else:
        fig = px.histogram(df)
    return fig


if __name__ == '__main__':
    app.run(debug=True)

