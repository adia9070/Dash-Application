import os
import dash
import dash_core_components as dcc
import dash_html_components as html
import pandas as pd
import plotly.graph_objs as go
import matplotlib.pyplot as plt
from dash.dependencies import Input, Output
from datetime import datetime as dt


external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

Superstore =pd.read_excel("Global Superstore.xls", parse_date=["Order Date"], index_col="Order Date")
report = Superstore["2011":"2014"][["Profit","Sales"]].resample("Y").sum()
region_wise = Superstore.groupby("Region")[["Profit","Sales"]].sum()


app = dash.Dash(__name__,external_stylesheets=external_stylesheets)

server = app.server


app.layout = html.Div([

    html.Div([
        html.H1("Interactive DashBoard")

        ],style={'textAlign':'center'}),

    html.Div([
    dcc.Graph(id="summary_report",
    figure={
        'data': [
        go.Scatter(x = report.index, y = report.Profit, name="Profit",marker=dict(color='#851e52')),
        go.Scatter(x = report.index, y = report.Sales, name="Sales",marker=dict(color='#d3560e')),
        go.Bar(x = report.index, y = report.Sales, name="Sales", marker=dict(color='#d3560e')),
        go.Bar(x = report.index, y = report.Profit, name="Profit", marker=dict(color='#851e52'))
        ],

        'layout': go.Layout(
            title = "Report Summary 2011-2014",
            font=dict(family='Arial', size=14, color='#000000'),
            xaxis = dict(title="Year",tickfont=dict(size=14, color='#000000'), showgrid=True),
            yaxis = dict(title="USD(Million)",tickfont=dict(size=14, color='#000000')),
            legend=dict(x=1,y=1)
            )
            }

        

        )
    ]),
        html.Div([
        html.H2("Report"),
        dcc.Graph(id="Region_Wise",
        figure={
        'data':[
        go.Bar(x=region_wise.index, y=region_wise.Profit, name="Profit", marker=dict(color='#851e52')),
        go.Bar(x=region_wise.index, y=region_wise.Sales, name="Sales", marker=dict(color='#d3560e'))
        ],

        'layout': go.Layout(
            title="Region Wise Report 2011-2014",
            font=dict(family='Arial', size=14, color='#000000'),
            xaxis=dict(title="Region", tickfont=dict(size=14, color='#000000'),tickangle=-45, showgrid=True),
            yaxis=dict(title="USD(Million)",tickfont=dict(size=14, color='#000000'))
            )
        }
        )
        ]),

        html.Div([
            html.Label("Region"),
            dcc.Dropdown(
                id="region_input",
                options=[
                {'label':"Africa", 'value':"Africa"},
                {'label':"Canada", 'value':"Canada"},
                {'label':"Caribbean", 'value':"Caribbean"},
                {'label':"Central", 'value':"Central"},
                {'label':"Central Asia", 'value':"Central Asia"},
                {'label':"East", 'value':"East"},
                {'label':"EMEA", 'value':"EMEA"},
                {'label':"North", 'value':"North"},
                {'label':"North Asia", 'value':"North Asia"},
                {'label':"Oceania", 'value':"Oceania"},
                {'label':"South", 'value':"South"},
                {'label':"Southeast Asia", 'value':"Southeast Asia"},
                {'label':"West", 'value':"West"},
                ], value='Central Asia'
                ),
            dcc.Graph(id="region_output")
            ]),

        html.Div([
            dcc.DatePickerRange(
            id='my-date-picker-range',
            min_date_allowed=dt(2011, 1, 1),
            max_date_allowed=dt(2014, 12, 31),
            initial_visible_month=dt(2011, 1, 1),
            end_date=dt(2014, 12, 31)
        ), dcc.Graph(id='output-container-date-picker-range')
        ]),

        html.Div([

            ])
                 
    ],style={'columnCount':2})
@app.callback(
    Output("region_output", 'figure'),
    [Input("region_input", "value")]
    )
def update_region_summary(region):
    region_input = Superstore[Superstore.Region==region][["Profit","Sales","Country"]].groupby("Country").sum()
    return {
    'data':[
    go.Bar(x = region_input.index, y=region_input.Profit, name='Profit', marker=dict(color='#851e52')),
    go.Bar(x=region_input.index, y=region_input.Sales, name='Sales', marker=dict(color='#d3560e'))
    ],
    'layout':go.Layout(
        title=region+" Wise Summary 2011-2014",
        font=dict(family='Arial', size=14, color='#000000'),
        xaxis=dict(title="Region", tickfont=dict(size=14, color='#000000'), showgrid=True),
        yaxis=dict(title="USD(Million)",tickfont=dict(size=14, color='#000000'))
        )
    }

@app.callback(
    Output('output-container-date-picker-range', 'figure'),
    [Input('my-date-picker-range', 'start_date'),
    Input('my-date-picker-range', 'end_date')
    ]
    )
def update_date(start_date, end_date):
    try:
        report_user = Superstore[start_date:end_date][["Profit","Sales"]].resample("D").sum()
    except:
        pass

    return {
    'data':[
    go.Bar(x=report_user.index, y=report_user.Profit, name="Profit", marker=dict(color='#851e52')),
    go.Bar(x=report_user.index, y=report_user.Sales, name="Sales", marker=dict(color='#d3560e'))
    ],
    'layout':go.Layout(
            title="Region Wise Report"+ start_date+"-"+end_date,
            font=dict(family='Arial', size=14, color='#000000'),
            xaxis=dict(title="Region", tickfont=dict(size=14, color='#000000'),tickangle=-45, showgrid=True),
            yaxis=dict(title="USD(Million)",tickfont=dict(size=14, color='#000000'))
            )
    }

if __name__ == '__main__':
    app.run_server(debug=True)
