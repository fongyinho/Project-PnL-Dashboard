import dash
from dash import html, dcc, Input, Output, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
from dash.dash_table.Format import Format, Scheme
import plotly.graph_objects as go

FILE_PATH = "EXCEL_BI_ALLDATA.xlsx"
SHEET_MASTER = "Master"
SHEET_BUDGET = "项目预算数据（测试版本）"
SHEET_ACTUAL = "项目实际数据（测试版本）"

df_master = pd.read_excel(FILE_PATH, sheet_name=SHEET_MASTER)
df_master.columns = df_master.columns.str.strip().str.replace("\n", "").str.replace(" ", "")
df_master = df_master.fillna("-")
df_projects = pd.DataFrame({
    "项目编号": df_master["项目编号"],
    "项目名称": df_master["项目名称"],
    "产品经理": df_master["产品经理"],
    "立项时间": df_master["立项时间"],
    "结项预期": df_master["结项预期"],
    "一级部门": df_master["一级部门"],
    "二级部门": df_master["二级部门"],
    "项目负责人": df_master["项目负责人"],
    "项目经理": df_master["项目经理"],
    "重点项目": df_master["重点项目"],
    "状态": df_master["状态"],
    "项目类型": df_master["项目类型TDP/PDP"]
})

df_budget_all = pd.read_excel(FILE_PATH, sheet_name=SHEET_BUDGET)
df_actual_all = pd.read_excel(FILE_PATH, sheet_name=SHEET_ACTUAL)
df_actual_all["SIPM125.SQSJ"] = pd.to_datetime(df_actual_all["SIPM125.SQSJ"], errors="coerce")
df_actual_all["月份"] = df_actual_all["SIPM125.SQSJ"].dt.to_period("M")

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "Project Dashboard"
app.index_string = '''
<!DOCTYPE html>
<html>
<head>
    {%metas%}
    <title>{%title%}</title>
    {%favicon%}
    {%css%}
    <style>
        body { font-family: "Microsoft YaHei", sans-serif; }
        @media print {
            .print-page-break {
                page-break-before: always;
            }
        }
    </style>
</head>
<body>
    {%app_entry%}
    <footer>
        {%config%}
        {%scripts%}
        {%renderer%}
    </footer>
</body>
</html>
'''
CUSTOM_STYLE = {
    'fontSize': '14px',
    'lineHeight': '1.1',
    'padding': '0.2rem 0.5rem'
}

def get_otd_table_data(project_id):
    """
    一次性计算：费用大类汇总、科目明细、月度实际
    """
    CATEGORY_ORDER = [
        "Material Cost", "Tooling & Fixture Cost", "Mould Cost", "Internal Testing Cost", "Testing & Inspection Cost", "Internal Prototyping Cost", "Prototype Sample Cost",
        "Equipment Commissioning Cost", "Outsourced R&D Cost", "Installation & Modification Cost", "Repair Cost", "Fuel & Energy Cost", "Internal Simulation Cost",
        "Conference / Meeting Expenses", "Office Expenses", "Travel Expenses", "Communication Expenses", "Postage / Shipping Expenses", "Business Hospitality Expenses", "Rental Fee",
        "Product Certification Fee", "Patent Annual Fee", "Patent Application Fee", "Consulting Service Fee", "Technical Books & Reference Materials", "Employee Welfare Expenses",
        "Employee Training & Education Funds", "Recruitment Fee", "Others / Miscellaneous", "Fixed Assets", "Intangible Assets (Software)"
    ]
    FYDLIST = ["R&D Expense", "Administrative Expense", "Fixed Assets / Intangible Assets", "Labour Cost"]
    dfb = df_budget_all[df_budget_all['项目编号'] == project_id].copy()
    dfa = df_actual_all[df_actual_all['SIPM125.NO'] == project_id].copy()

    df_bcat = dfb.groupby('费用大类')['金额(元)'].sum()
    df_acat = dfa.groupby('SIPM127.FYDTYPE')['SIPM125.BXJE'].sum()
    df_summary = pd.DataFrame({
        "费用大类": FYDLIST,
        "预算金额": [df_bcat.get(cat, 0)/1000 for cat in FYDLIST],
        "实际金额": [df_acat.get(cat, 0)/1000 for cat in FYDLIST]
    })
    df_summary["占比"] = df_summary["实际金额"] / df_summary["预算金额"]
    df_summary["剩余"] = df_summary["预算金额"] - df_summary["实际金额"]

    df_bsub = dfb.groupby('科目名称')['金额(元)'].sum()
    df_asub = dfa.groupby('SIPM125.KMMC')['SIPM125.BXJE'].sum()
    df_detail = pd.DataFrame({
        "科目名称": CATEGORY_ORDER,
        "预算金额": [df_bsub.get(s, 0)/1000 for s in CATEGORY_ORDER],
        "实际金额": [df_asub.get(s, 0)/1000 for s in CATEGORY_ORDER]
    })
    df_detail["占比"] = df_detail["实际金额"] / df_detail["预算金额"]
    df_detail["剩余"] = df_detail["预算金额"] - df_detail["实际金额"]
    for col in ["预算金额","实际金额","剩余"]:
        df_summary[col] = df_summary[col].apply(lambda x: "-" if x == 0 else round(x, 2))
        df_detail[col] = df_detail[col].apply(lambda x: "-" if x == 0 else round(x, 2))

    df_actual_month = (
        dfa.groupby("月份")["SIPM125.BXJE"]
        .sum()
        .sort_index()
    )
    df_monthly = pd.DataFrame({
        "月份": df_actual_month.index.astype(str),
        "实际金额": df_actual_month.values / 1000
    })
    return df_summary, df_detail, df_monthly

def create_donut_chart(usage_ratio):
    percentage = round(usage_ratio * 100, 1)
    percentage_clamped = min(max(percentage, 0), 100)
    total_segments = 20
    filled_segments = int(round(percentage_clamped * total_segments / 100))
    unfilled_segments = total_segments - filled_segments
    fig = go.Figure(data=[go.Pie(
        values=[1]*total_segments,
        hole=0.55,
        marker_colors=["#1d3a6d"]*filled_segments + ["#e0e0e0"]*unfilled_segments,
        marker_line=dict(color="white", width=2),
        direction='clockwise',
        rotation=0,
        textinfo="none",
        hoverinfo="skip",
        hovertemplate=None,
        sort=False
    )])
    fig.update_layout(
        annotations=[dict(
            text=f"{percentage}%",
            x=0.5, y=0.5, showarrow=False,
            font_size=20,
            font_color="red" if percentage>100 else "black"
        )],
        margin=dict(t=0, b=0, l=0, r=0),
        showlegend=False
    )
    return fig
def build_budget_bar_chart(categories, actual_data, budget_data):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=categories,
        y=budget_data,
        name='budget',
        marker_color="#a4c8df",
        width=0.5,
        hovertemplate='%{x}<br>Budget: %{y:.2f} <extra></extra>'
    ))
    fig.add_trace(go.Bar(
        x=categories,
        y=actual_data,
        name='actual',
        marker_color="#1d3a6d",
        width=0.3,
        text=[f"{(act / bud * 100):.1f}%" if bud > 0 else ""
              for act, bud in zip(actual_data, budget_data)],
        textposition="outside",
        hovertemplate='%{x}<br>Actual: %{y:.2f} <extra></extra>'
    ))
    fig.update_layout(
        barmode='overlay',
        height=380,
        margin=dict(l=30, r=30, t=20, b=30),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        xaxis_title=None,
        yaxis_title="kCNY",
        font=dict(size=14),
    )
    return fig
def build_monthly_line_chart(df_monthly):
    """
    df_monthly:
        列名: '月份'  (形如 '2024-07' 或 Period)
              '实际金额' (单位：千元)
    """
    x_all = pd.to_datetime(df_monthly["月份"].astype(str))
    y_all = df_monthly["实际金额"].values

    if (y_all != 0).any():
        first_idx = (y_all != 0).argmax()  
        x = x_all[first_idx:]
        y = y_all[first_idx:]
    else:
        x = x_all
        y = y_all

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x,
        y=y,
        mode="lines+markers",
        name="Actual",
        line=dict(color="#1d3a6d", width=2),
        hovertemplate="%{x|%Y/%m}, %{y:.2f}"
    ))

    fig.update_layout(
        height=380,
        margin=dict(l=30, r=30, t=20, b=30),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        xaxis=dict(
            title=None,
            tickformat="%Y/%m",  
            #dtick="M1",         
        ),
        yaxis=dict(
            title="kCNY"
        ),
        font=dict(size=14),
    )
    return fig
def build_stage_bar_chart(project_type, budget_df, actual_df):
    stage_order = (
        ["R0", "R1", "R2", "R3", "R4", "R5"]
        if project_type == "TDP"
        else ["M0", "M1", "M2", "M3", "M4", "M5", "M6"]
    )

    budget_df = budget_df.copy()
    actual_df = actual_df.copy()
    budget_df["金额千元"] = budget_df["金额(元)"] / 1000
    actual_df["金额千元"] = actual_df["SIPM125.BXJE"] / 1000
    budget_grouped = budget_df.groupby("阶段")["金额千元"].sum().reindex(stage_order, fill_value=0)
    actual_grouped = actual_df.groupby("SIPM125.JD")["金额千元"].sum().reindex(stage_order, fill_value=0)

    max_val = max(budget_grouped.max(), actual_grouped.max())
    if max_val <= 0:
        max_val = 1

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=stage_order,
        y=budget_grouped.values,
        name="budget",
        marker_color="#a4c8df",
        width=0.5
    ))
    fig.add_trace(go.Bar(
        x=stage_order,
        y=actual_grouped.values,
        name="actual",
        marker_color="#1d3a6d",
        width=0.3
    ))

    annotations = []
    for x, bud, act in zip(stage_order, budget_grouped, actual_grouped):
  
        if bud > 0:
            percent = act / bud * 100
            label_pct = f"{percent:.1f}%"
        else:
            label_pct = "0%"
        annotations.append(dict(
            x=x,
            y=act + max_val * 0.05,
            xref="x",
            yref="y",
            text=label_pct,
            showarrow=False,
            font=dict(size=12, color="#333", family="Microsoft YaHei"),
        ))
  
        annotations.append(dict(
            x=x,
            xref="x",
            y=-0.22,             
            yref="paper",
            text=f"Budget:{bud:.1f}<br>Actual:{act:.1f}",
            showarrow=False,
            align="center",
            font=dict(size=12, color="#333", family="Microsoft YaHei"),
        ))

    ymax = max_val * 1.5
    fig.update_layout(
        annotations=annotations,
        barmode="overlay",
        height=380,
        xaxis_title=None,
        yaxis_title="kCNY",
        margin=dict(l=30, r=30, t=20, b=60),  
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(range=[0, ymax]),
        font=dict(size=14, family="Bahnschrift"),
    )
    return fig

def fmt_date(x):
    if hasattr(x, "date"):
        return x.date()
    return x
def build_budget_overview(total_budget, total_actual, pie_fig):
    diff = total_budget - total_actual
    if diff >= 0:
        diff_text = f"Total Balance：{diff:.2f} "
        diff_color = "black"
    else:
        diff_text = f"Exceeded：{abs(diff):.2f} "
        diff_color = "#cc0000"
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.P(f"Total Budget：{total_budget:.2f} ",
                       className="fw-bold",
                       style={"fontSize": "15px", "marginBottom": "5px", "marginTop": "10px"}),
                html.P(f"Total Actual：{total_actual:.2f} ",
                       className="fw-bold",
                       style={"fontSize": "15px", "marginBottom": "5px"}),
                html.P(diff_text,
                       className="fw-bold",
                       style={"fontSize": "15px", "color": diff_color}),
            ], width=5),
            dbc.Col([
                dcc.Graph(
                    figure=pie_fig,
                    config={"displayModeBar": False},
                    style={"height": "160px", "marginLeft": "-40px"}
                )
            ], width=5)
        ])
    ], style={
        "backgroundColor": "white",
        "padding": "16px",
        "borderRadius": "8px",
        "boxShadow": "0 2px 6px rgba(0,0,0,0.05)"
    })
def build_project_info(project_row):
    return dbc.Row([
        dbc.Col(
            dbc.Card(
                dbc.Row([
                    dbc.Col([
                        html.H6("Project Info", className="fw-bold", style=CUSTOM_STYLE),
                        html.Div(f"Project Name：{project_row['项目名称']}", style=CUSTOM_STYLE),
                        html.Div(f"Start Date：{fmt_date(project_row['立项时间'])}", style=CUSTOM_STYLE),
                        html.Div(f"End Date：{fmt_date(project_row['结项预期'])}", style=CUSTOM_STYLE),
                        html.Div(f"Status：{project_row['状态']}", style=CUSTOM_STYLE),
                        html.Div(f"Proj Type：{project_row['项目类型']}", style=CUSTOM_STYLE),
                    ], width=7),
                    dbc.Col([
                        html.H6(" ", className="fw-bold", style=CUSTOM_STYLE),
                        html.Div(f"Division：{project_row['一级部门']}", style=CUSTOM_STYLE),
                        html.Div(f"Department：{project_row['二级部门']}", style=CUSTOM_STYLE),
                        html.Div(f"Main PIC：{project_row['项目负责人']}", style=CUSTOM_STYLE),
                        html.Div(f"Product Manager：{project_row['产品经理']}", style=CUSTOM_STYLE),
                        html.Div(f"Project Manager：{project_row['项目经理']}", style=CUSTOM_STYLE),
                    ], width=5),
                ]),
                body=True,
                className="mt-3",
                style={"backgroundColor": "#f8f9fa", "boxShadow": "0 2px 6px rgba(0,0,0,0.05)"}
            ),
            width=6
        ),
        dbc.Col([
            html.Div(id="budget-overview")
        ], width=6)
    ])

app.layout = dbc.Container([
    html.Div([
        dbc.Row([
            dbc.Col(
                html.H4([
                    "Project P&L Dashboard",
                    html.Span("（Unit：kCNY）", style={"fontSize": "14px", "marginLeft": "6px"})
                ], className="fw-bold",
                   style={"fontSize": "22px", "color": "white", "margin": "0"}),
                width="auto",
                style={"display": "flex", "alignItems": "center"}
            ),
            dbc.Col(
                dcc.Dropdown(
                    id="project-selector",
                    options=[{"label": pid, "value": pid} for pid in df_projects["项目编号"]],
                    value=df_projects["项目编号"].iloc[0],
                    style={"width": "200px", "fontSize": "14px", "borderRadius": "4px"},
                ),
                width="auto",
                style={"display": "flex", "alignItems": "center"}
            )
        ], className="g-2")
    ], style={
        "backgroundColor": "#20448B",
        "padding": "10px",
        "borderRadius": "10px",
        "marginBottom": "25px"
    }),
    html.Div(id="project-info"),
    dbc.Row([
         dbc.Col([
            html.H6("Project P&L", className="fw-bold",
                    style={"fontSize": "16px", "color": "#20448B"}),
            dcc.Loading(id="loading-otd", type="default", children=[
                html.Div([
                    html.H6("▶ By Category", className="fw-bold",
                            style={"fontSize": "14px", "color": "#20448B"}),
                    dash_table.DataTable(
                        id="otd-table-summary",
                        merge_duplicate_headers=True,
                        style_table={"overflowX": "auto", "marginBottom": "24px"},
                        style_header={"textAlign": "center","backgroundColor": "#d6e4f5", "fontWeight": "bold"},
                        style_cell={"textAlign": "center", "fontSize": "14px"},
                        style_cell_conditional=[
                            {"if": {"column_id": "费用大类"},"textAlign": "left"},
                            {"if": {"column_id": "预算金额"},"textAlign": "right", "fontFamily": "Calibri", "paddingRight": "8px"},
                            {"if": {"column_id": "实际金额"},"textAlign": "right", "fontFamily": "Calibri", "paddingRight": "8px"},
                            {"if": {"column_id": "占比"},"textAlign": "right", "fontFamily": "Calibri", "paddingRight": "8px"},
                            {"if": {"column_id": "剩余"},"textAlign": "right", "fontFamily": "Calibri", "paddingRight": "8px"},
                        ],
                        
                    ),
                    html.H6("▶ By Subject", className="fw-bold",
                            style={"fontSize": "14px", "color": "#20448B"}),
                    dash_table.DataTable(
                        id="otd-table-detail",
                        merge_duplicate_headers=True,
                        style_table={"overflowX": "auto"},
                        style_cell={"textAlign": "center", "fontSize": "14px"},
                        style_header={"textAlign": "center", "backgroundColor": "#d6e4f5", "fontWeight": "bold"},
                        style_cell_conditional=[
                            {"if": {"column_id": "科目名称"},"textAlign": "left"},
                            {"if": {"column_id": "预算金额"},"textAlign": "right", "fontFamily": "Calibri", "paddingRight": "8px"},
                            {"if": {"column_id": "实际金额"},"textAlign": "right", "fontFamily": "Calibri", "paddingRight": "8px"},
                            {"if": {"column_id": "占比"},"textAlign": "right", "fontFamily": "Calibri", "paddingRight": "8px"},
                            {"if": {"column_id": "剩余"},"textAlign": "right", "fontFamily": "Calibri", "paddingRight": "8px"},
                        ],
                    ),
                ])
            ])
        ], width=6),
        dbc.Col([
            html.H6("▶ Budget vs Actual (By Category)", className="fw-bold",
                    style={"fontSize": "14px", "color": "#20448B"}),
            html.Div(id="budget-bar"),
            html.H6("▶ Expenses Trend by Month", className="fw-bold",
                    style={"fontSize": "14px", "color": "#20448B"}),
            html.Div(id="budget-line"),
            html.H6("▶ Budget vs Actual (By Stage)", className="fw-bold",
                    style={"fontSize": "14px", "color": "#20448B"}),
            html.Div(id="stage-bar"),
        ], width=6),
    ], className="mt-4"),
    dbc.Row([
        dbc.Col([
            html.Div([
                html.H6("▶ Budget vs Actual (By Stage)", className="fw-bold",
                        style={"fontSize": "14px", "color": "#20448B"}),
                
                    dash_table.DataTable(
                        id="otd-matrix-table",
                        merge_duplicate_headers=True,  
                        style_table={"overflowX": "auto"},
                        style_cell={
                            "textAlign": "center",
                            #"fontSize": "13px",
                            #"padding": "4px",
                            #"border": "1px solid #d6e4f5",  
                        },
                        style_header={
                            "backgroundColor": "#d6e4f5",
                        #    "fontWeight": "bold",
                            #"border": "1px solid #d6e4f5",
                        },
                    ),                    
                ], className="print-page-break") 
        ], width=12)
    ], className="mt-4"),
], fluid=True, style={"padding": "2rem"})

@app.callback(
    Output("project-info", "children"),
    Input("project-selector", "value")
)
def update_project_info(project_id):
    row = df_projects[df_projects["项目编号"] == project_id].iloc[0]
    return build_project_info(row)
@app.callback(
    Output("otd-table-summary", "data"),
    Output("otd-table-summary", "columns"),
    Output("otd-table-detail", "data"),
    Output("otd-table-detail", "columns"),
    Input("project-selector", "value")
)
def update_otd_tables(project_id):
    df_summary, df_detail, _ = get_otd_table_data(project_id)
    summary_columns = [
        {"name": "Expense Category", "id": "费用大类", "type": "text"},
        {"name": "Budget Amount", "id": "预算金额", "type": "numeric",
         "format": Format(precision=2, scheme=Scheme.fixed)},
        {"name": "Actual Amount", "id": "实际金额", "type": "numeric",
         "format": Format(precision=2, scheme=Scheme.fixed)},
        {"name": "Usage", "id": "占比", "type": "numeric",
         "format": Format(precision=1, scheme=Scheme.percentage)},
        {"name": "Balance", "id": "剩余", "type": "numeric",
         "format": Format(precision=2, scheme=Scheme.fixed)},
    ]
    detail_columns = [
        {"name": "Subject", "id": "科目名称", "type": "text"},
        {"name": "Budget Amount", "id": "预算金额", "type": "numeric",
         "format": Format(precision=2, scheme=Scheme.fixed)},
        {"name": "Actual Amount", "id": "实际金额", "type": "numeric",
         "format": Format(precision=2, scheme=Scheme.fixed)},
        {"name": "Usage", "id": "占比", "type": "numeric",
         "format": Format(precision=1, scheme=Scheme.percentage)},
        {"name": "Balance", "id": "剩余", "type": "numeric",
         "format": Format(precision=2, scheme=Scheme.fixed)},
    ]
    return (
        df_summary.to_dict("records"), summary_columns,
        df_detail.to_dict("records"), detail_columns
    )
@app.callback(
    Output("budget-overview", "children"),
    Output("budget-bar", "children"),
    Output("budget-line", "children"),
    Output("stage-bar", "children"),
    Input("project-selector", "value")
)
def update_budget_overview(project_id):
    df_summary, df_detail, df_monthly = get_otd_table_data(project_id)
    total_budget = df_summary["预算金额"].replace("-", 0).astype(float).sum()
    total_actual = df_summary["实际金额"].replace("-", 0).astype(float).sum()
    usage_ratio = total_actual / total_budget if total_budget != 0 else 0
    pie_fig = create_donut_chart(usage_ratio)

    categories = df_summary["费用大类"].tolist()
    budget_data = df_summary["预算金额"].replace("-", 0).astype(float).tolist()
    actual_data = df_summary["实际金额"].replace("-", 0).astype(float).tolist()
    bar_fig = build_budget_bar_chart(categories, actual_data, budget_data)

    line_fig = build_monthly_line_chart(df_monthly)

    project_type = df_projects[df_projects["项目编号"] == project_id]["项目类型"].values[0]
    budget_df = df_budget_all[df_budget_all["项目编号"] == project_id].copy()
    actual_df = df_actual_all[df_actual_all["SIPM125.NO"] == project_id].copy()
    bar_fig_stage = build_stage_bar_chart(project_type, budget_df, actual_df)
    return (
        build_budget_overview(total_budget, total_actual, pie_fig),
        dcc.Graph(figure=bar_fig, config={"displayModeBar": False}, style={"height": "380px"}),
        dcc.Graph(figure=line_fig, config={"displayModeBar": False}, style={"height": "380px"}),
        dcc.Graph(figure=bar_fig_stage, config={"displayModeBar": False}, style={"height": "380px"}),
    )
@app.callback(
    Output("otd-matrix-table", "data"),
    Output("otd-matrix-table", "columns"),
    Output("otd-matrix-table", "style_cell"),
    Output("otd-matrix-table", "style_cell_conditional"),
    Output("otd-matrix-table", "style_header_conditional"),
    Input("project-selector", "value")
)
def update_matrix(project_id):

    df_summary, df_detail, _ = get_otd_table_data(project_id)
    project_type = df_projects[df_projects["项目编号"] == project_id]["项目类型"].values[0]
    stages = ["R0", "R1", "R2", "R3", "R4", "R5"] if project_type == "TDP" else \
             ["M0", "M1", "M2", "M3", "M4", "M5", "M6"]
    subjects = df_detail["科目名称"].tolist()
    dfb = df_budget_all[df_budget_all["项目编号"] == project_id].copy()
    dfa = df_actual_all[df_actual_all["SIPM125.NO"] == project_id].copy()
    dfb["金额千元"] = dfb["金额(元)"] / 1000
    dfa["金额千元"] = dfa["SIPM125.BXJE"] / 1000

    columns = [{"name": ["Subject", ""], "id": "科目名称"}]
    for i, s in enumerate(stages):
        columns.append({"name": [s, "Budget"], "id": f"{s}_预算"})
        columns.append({"name": [s, "Actual"], "id": f"{s}_实际"})
        if i < len(stages) - 1:
            columns.append({"name": ["", ""], "id": f"{s}_sep"})

    data = []
    for sub in subjects:
        row = {"科目名称": sub}
        for i, s in enumerate(stages):
            bud = dfb[(dfb["科目名称"] == sub) & (dfb["阶段"] == s)]["金额千元"].sum()
            act = dfa[(dfa["SIPM125.KMMC"] == sub) & (dfa["SIPM125.JD"] == s)]["金额千元"].sum()
            row[f"{s}_预算"] = "-" if bud == 0 else f"{bud:.2f}"
            row[f"{s}_实际"] = "-" if act == 0 else f"{act:.2f}"
            if i < len(stages) - 1:
                row[f"{s}_sep"] = ""
        data.append(row)

    sep_cols = [f"{s}_sep" for s in stages[:-1]]

    style_cell = {
        "textAlign": "center",
        "fontSize": "13px",
        "padding": "4px",
        "borderTop": "1px solid #a0bde6",
        "borderBottom": "1px solid #a0bde6",
        "borderLeft": "1px solid #a0bde6",
        "borderRight": "1px solid #a0bde6",
    }

    style_cell_conditional = [
        {
            "if": {"column_id": col},
            "backgroundColor": "white",
            "borderLeft": "1px solid #a0bde6",
            "borderRight": "1px solid #a0bde6",
            "borderTop": "none",
            "borderBottom": "none",
            "padding": "0px",
            "width": "6px",
            "minWidth": "6px",
            "maxWidth": "6px",
        }
        for col in sep_cols
    ]

    number_cols = [c["id"] for c in columns if ("预算" in c["id"] or "实际" in c["id"])]
    style_cell_conditional += [
        {"if": {"column_id": col}, "fontFamily": "Calibri"}
        for col in number_cols
    ]

    style_header_conditional = [
        {
            "if": {"column_id": col},
            "backgroundColor": "white",
            "borderLeft": "1px solid #a0bde6",
            "borderRight": "1px solid #a0bde6",
            "borderTop": "1px solid #a0bde6",
            "borderBottom": "none",
            "padding": "0px",
        }
        for col in sep_cols
    ]
    return data, columns, style_cell, style_cell_conditional, style_header_conditional

if __name__ == "__main__":
    app.run(port=8080, debug=False)

