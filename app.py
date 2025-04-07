import webbrowser
import time
import dash
from dash import dcc, html, dash_table, Input, Output, State, ctx
import pandas as pd
import dash_bootstrap_components as dbc
import io
import base64
import smtplib
from email.message import EmailMessage

# Initialize Dash App
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Layout
app.layout = dbc.Container([
    html.H1("Excel Filtering & Email Automation Dashboard", className="text-center mt-4"),

    # Upload Section
    dcc.Upload(
        id="upload-excel",
        children=html.Button("Upload Excel File", className="btn btn-primary"),
        multiple=False
    ),
    
     # Dynamic Filters Section
    html.Div([
        html.Button("Add Filter", id="add-filter-btn", className="btn btn-info mt-2"),
    ]),
    html.Div(id="filters-container", children=[]),  # Holds dynamic filters

    # Apply Filter Button
    html.Button("Apply Filters", id="apply-filters-btn", className="btn btn-success mt-2"),

    # Table to Display Filtered Data
    html.Hr(),
    dash_table.DataTable(
        id="filtered-table",
        columns=[],
        data=[],
        page_size=10,
        style_table={"overflowX": "auto"}
    ),

    # Download Filtered Data
    html.Button("Download Filtered Data", id="download-btn", className="btn btn-warning mt-2"),
    dcc.Download(id="download-dataframe-xlsx"),


    html.Hr(),

    # Select Columns for Emailing
    html.Label("Select Name and Email Columns (Pairs)"),
    dbc.Row([
        dbc.Col(dcc.Dropdown(id="name-column-1", placeholder="Select Name Column 1")),
        dbc.Col(dcc.Dropdown(id="email-column-1", placeholder="Select Email Column 1")),
    ], className="mb-3"),
    dbc.Row([
        dbc.Col(dcc.Dropdown(id="name-column-2", placeholder="Select Name Column 2")),
        dbc.Col(dcc.Dropdown(id="email-column-2", placeholder="Select Email Column 2")),
    ], className="mb-3"),

    # Email Sending Section
    html.H2("Email Automation", className="text-center mt-4"),
    dbc.Row([
        dbc.Col(dbc.Input(id="sender-name", type="text", placeholder="Sender Name"), width=6),
        dbc.Col(dbc.Input(id="sender-email", type="email", placeholder="Sender Email"), width=6),
    ], className="mb-3"),
    dbc.Row([
        dbc.Col(dbc.Input(id="sender-password", type="password", placeholder="Sender App Password"), width=6),
        dbc.Col(dbc.Input(id="company-name", type="text", placeholder="Company Name"), width=6),
    ], className="mb-3"),
    dbc.Input(id="email-subject", type="text", placeholder="Enter Email Subject", className="mb-3"),
    dbc.Textarea(
        id="email-template",
        placeholder="Enter email template with {employee_name}, {company_name}, {designation}, {sender_name}",
        className="mb-3",
        rows=5,
        value="""Dear {employee_name},

Greetings from {company_name}!

We are reaching out to you regarding your role as {designation}. 
Please review the attached document and let us know if you need any further details.

Best regards,  
{sender_name}  
{company_name}"""
    ),
    html.Button("Send Emails", id="send-email", className="btn btn-success mt-3"),
    html.Div(id="email-status", className="mt-3")
], fluid=True)


# Function to Read Excel File
def parse_contents(contents):
    content_type, content_string = contents.split(",")
    decoded = base64.b64decode(content_string)
    return pd.read_excel(io.BytesIO(decoded))


# Combined Callback for Adding and Removing Filters
@app.callback(
    Output("filters-container", "children"),
    [Input("add-filter-btn", "n_clicks"), Input({"type": "remove-filter", "index": dash.ALL}, "n_clicks")],
    State("filters-container", "children"),
    State("upload-excel", "contents")
)
def update_filters(add_clicks, remove_clicks, existing_filters, contents):
    if not contents:
        return existing_filters  # No file uploaded, return unchanged

    df = parse_contents(contents)
    triggered_id = ctx.triggered_id

    if isinstance(triggered_id, dict) and triggered_id["type"] == "remove-filter":
        # Remove the clicked filter
        remove_index = triggered_id["index"]
        return [f for i, f in enumerate(existing_filters) if i != remove_index]

    if add_clicks:
        filter_id = len(existing_filters)  # Unique ID for each filter
        new_filter = html.Div([
            html.Label(f"Filter {filter_id + 1} Column"),
            dcc.Dropdown(
                id={"type": "filter-column", "index": filter_id},
                options=[{"label": col, "value": col} for col in df.columns],
                placeholder="Select a column"
            ),
            html.Label(f"Filter {filter_id + 1} Value"),
            dcc.Dropdown(id={"type": "filter-value", "index": filter_id}, placeholder="Select a value", disabled=True),
            html.Button("Remove", id={"type": "remove-filter", "index": filter_id}, className="btn btn-danger btn-sm mt-2"),
            html.Hr()
        ])
        existing_filters.append(new_filter)

    return existing_filters


# Callback to Populate Values for Selected Columns Dynamically
@app.callback(
    Output({"type": "filter-value", "index": dash.ALL}, "options"),
    Output({"type": "filter-value", "index": dash.ALL}, "disabled"),
    Input({"type": "filter-column", "index": dash.ALL}, "value"),
    State("upload-excel", "contents")
)
def update_filter_values(selected_columns, contents):
    if not contents:
        return [[]] * len(selected_columns), [True] * len(selected_columns)

    df = parse_contents(contents)
    updated_options = []
    updated_disabled = []

    for column in selected_columns:
        if column:
            values = [{"label": val, "value": val} for val in df[column].dropna().unique()]
            updated_options.append(values)
            updated_disabled.append(False)
        else:
            updated_options.append([])
            updated_disabled.append(True)

    return updated_options, updated_disabled


# Callback to Apply Multiple Filters
@app.callback(
    [Output("filtered-table", "columns"), Output("filtered-table", "data")],
    Input("apply-filters-btn", "n_clicks"),
    [State("upload-excel", "contents"), State({"type": "filter-column", "index": dash.ALL}, "value"),
     State({"type": "filter-value", "index": dash.ALL}, "value")]
)
def apply_filters(n_clicks, contents, filter_columns, filter_values):
    if not contents:
        return [], []

    df = parse_contents(contents)
    for col, val in zip(filter_columns, filter_values):
        if col and val:
            df = df[df[col] == val]

    return [{"name": i, "id": i} for i in df.columns], df.to_dict("records")


# Callback to Download Filtered Data
@app.callback(
    Output("download-dataframe-xlsx", "data"),
    Input("download-btn", "n_clicks"),
    State("filtered-table", "data")
)
def download_filtered_data(n_clicks, data):
    if n_clicks is None or not data:
        return dash.no_update

    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtered Data")

    output.seek(0)
    return dcc.send_bytes(output.getvalue(), "filtered_data.xlsx")


# Populate Dropdowns with Column Names
@app.callback(
    [Output("name-column-1", "options"), Output("email-column-1", "options"),
     Output("name-column-2", "options"), Output("email-column-2", "options")],
    Input("upload-excel", "contents")
)
def populate_dropdowns(contents):
    if not contents:
        return [], [], [], []
    df = parse_contents(contents)
    columns = [{"label": col, "value": col} for col in df.columns]
    return columns, columns, columns, columns


# Function to Send Email
def send_email(sender_email, sender_password, to_email, subject, body):
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return f"✅ Email sent to {to_email}"
    except Exception as e:
        return f"❌ Failed to send email to {to_email}: {e}"


# Callback to Send Emails

@app.callback(
    Output("email-status", "children"),
    Input("send-email", "n_clicks"),
    [State("sender-name", "value"), State("sender-email", "value"), State("sender-password", "value"),
     State("company-name", "value"), State("email-subject", "value"), State("email-template", "value"),
     State("filtered-table", "data"),  # ✅ Correct - Use filtered-table data
     State("name-column-1", "value"), State("email-column-1", "value"),
     State("name-column-2", "value"), State("email-column-2", "value")]
)

def send_emails(n_clicks, sender_name, sender_email, sender_password, company_name, email_subject, email_template, 
                 filtered_data, name_col_1, email_col_1, name_col_2, email_col_2):
    if not n_clicks or not filtered_data:
        return "⏳ Apply filters and click 'Send Emails' to start."
    
    df = pd.DataFrame(filtered_data)
    status_messages = []
    
    for _, row in df.iterrows():
        # Name 1 → Email 1
        if name_col_1 and email_col_1 and row.get(email_col_1):
            email_body = email_template.format(
                employee_name=row.get(name_col_1, "Employee"),
                company_name=company_name,
                designation=row.get("Designation", ""),
                sender_name=sender_name
            )
            status_messages.append(send_email(sender_email, sender_password, row[email_col_1], email_subject, email_body))
        
        # Name 2 → Email 2
        if name_col_2 and email_col_2 and row.get(email_col_2):
            email_body = email_template.format(
                employee_name=row.get(name_col_2, "Employee"),
                company_name=company_name,
                designation=row.get("Designation", ""),
                sender_name=sender_name
            )
            status_messages.append(send_email(sender_email, sender_password, row[email_col_2], email_subject, email_body))
    
    return html.Div([html.P(status) for status in status_messages])


# Run App
if __name__ == "__main__":
    app.run_server(host="0.0.0.0", port=10000, debug=False)

