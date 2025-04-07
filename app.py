import os
import webbrowser
import time
import dash
from dash import dcc, html, dash_table, Input, Output, State, ctx
import pandas as pd
import dash_bootstrap_components as dbc
import io
import base64
import smtplib
import xlsxwriter
from email.message import EmailMessage

# ✅ List of valid license keys
VALID_KEYS = {"INN123", "XYZ123", "MNO123"}  # Add more keys here

# ✅ Path to store the license key
LICENSE_FILE = "license.txt"

# ✅ Function to check the stored license key
def check_license():
    if os.path.exists(LICENSE_FILE):
        with open(LICENSE_FILE, "r") as file:
            saved_key = file.read().strip()
            return saved_key in VALID_KEYS  # ✅ Check if key is valid
    return False

# ✅ Function to save the license key
def save_license(key):
    with open(LICENSE_FILE, "w") as file:
        file.write(key)


# Initialize Dash App
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css"])

# ✅ License Key Layout (Shown only if key is missing)
# ✅ License Key Layout (Shown only if key is missing)
license_layout = dbc.Container([
    dcc.Location(id="url", refresh=True),  # Keeps track of the URL
    html.Div(id="redirect-script"),  # ✅ This will hold the JavaScript refresh script
    dbc.Row(
        dbc.Col(
            html.Div([
                html.H2("Enter License Key", className="text-center mb-4"),
                dbc.Input(id="license-key", type="text", placeholder="Enter your License Key", className="mb-2"),
                dbc.Button("Activate", id="activate-btn", color="success", className="btn-block"),
                html.Div(id="license-status", className="mt-2 text-center", style={"fontWeight": "bold"})
            ], style={
                "textAlign": "center", "padding": "20px", "width": "50%", "margin": "auto",
                "border": "1px solid #ddd", "borderRadius": "8px", "boxShadow": "2px 2px 10px rgba(0,0,0,0.1)"
            }),
            width=12
        )
    )
], fluid=True)




# Application Layout
app_layout = dbc.Container([
        dbc.Row(
        dbc.Col(
            html.Div(
                "Email Automation Dashboard",
                style={
                    "fontSize": "36px",  # Large font size
                    "textAlign": "center",  # Center align text
                    "padding": "20px",  # Add some padding
                    "backgroundColor": "#28d077",  # Blue background
                    "color": "white",  # White text color
                    "borderRadius": "10px",  # Rounded corners
                    "boxShadow": "2px 2px 10px rgba(0, 0, 0, 0.2)",  # Shadow effect
                    "fontWeight": "bold"  # Make text bold
                }
            ),
            width=12  # Full width
        )
    ),
    # Upload Section
    dbc.Row(
        dbc.Col(
            html.Div([
                dcc.Upload(
                    id="upload-excel",
                    children=html.Button("Upload Excel File", className="btn btn-primary"),
                    multiple=False
                ),
                html.Div(id="uploaded-file-name", style={
                    "textAlign": "center", "marginTop": "10px", "fontWeight": "bold"
                })
            ], style={
                "textAlign": "center", "padding": "20px"
            }),
            width=12
        )
    ),
    
    # Dynamic Filters Section
        dbc.Row(
            dbc.Col(
                html.Div(
                    html.Button("Add Filter", id="add-filter-btn", className="btn btn-info mt-2"),
                    style={
                        "textAlign": "center", "marginTop": "10px", "marginBottom": "10px", "padding": "10px",
                        "backgroundColor": "#f8f9fa", "border": "1px solid #ddd", "borderRadius": "8px",
                        "boxShadow": "2px 2px 10px rgba(0, 0, 0, 0.1)", "width": "50%", "marginLeft": "auto", "marginRight": "auto"
                    }
                ),
                width=12
            )
        ),
        html.Div(id="filters-container", children=[], style={"textAlign": "center", "width": "30%", "margin": "auto"}),  # Holds dynamic filters



    # Apply Filter Button
        dbc.Row(
        dbc.Col(
            html.Div(
                html.Button("Apply Filters", id="apply-filters-btn", className="btn btn-success mt-2"),
                style={
                    "textAlign": "center", "marginTop": "10px", "padding": "10px",
                    "backgroundColor": "#f8f9fa", "border": "1px solid #ddd", "borderRadius": "8px",
                    "boxShadow": "2px 2px 10px rgba(0, 0, 0, 0.1)", "width": "30%", "marginLeft": "auto", "marginRight": "auto"
                }
            ),
            width=20
        )
    ),

    # Table to Display Filtered Data
    html.Hr(),
    dbc.Row(
        dbc.Col(
            dash_table.DataTable(
                id="filtered-table",
                columns=[],
                data=[],
                page_size=10,
                style_table={"overflowX": "auto", "width": "80%", "margin": "auto", "border": "1px solid #ddd", "borderRadius": "8px", "boxShadow": "2px 2px 10px rgba(0, 0, 0, 0.1)", "padding": "10px", "backgroundColor": "#ffffff"},
                style_header={"backgroundColor": "#007BFF", "color": "white", "fontWeight": "bold", "textAlign": "center"},
                style_cell={"textAlign": "center", "padding": "8px", "border": "1px solid #ddd"},
                style_data={"backgroundColor": "#f8f9fa", "color": "#333"}
            ),
            width=12
        )
    ),
    # Download Filtered Data
    dbc.Row(
        dbc.Col(
            html.Div(
                html.Button("Download Filtered Data", id="download-btn", className="btn btn-warning mt-2"),
                style={
                    "textAlign": "center", "marginTop": "10px", "padding": "10px",
                    "backgroundColor": "#f8f9fa", "border": "1px solid #ddd", "borderRadius": "8px",
                    "boxShadow": "2px 2px 10px rgba(0, 0, 0, 0.1)", "width": "30%", "marginLeft": "auto", "marginRight": "auto"
                }
            ),
            width=12
        )
    ),
    dcc.Download(id="download-dataframe-xlsx"),


    html.Hr(),

    # Select Columns for Emailing
    dbc.Row(
            dbc.Col(
                html.Div(
                    [
                        html.Label("Select Name and Email Columns (Pairs)", style={
                            "fontSize": "20px", "fontWeight": "bold", "textAlign": "center", "color": "#333"
                        }),
                        dbc.Row([
                            dbc.Col(dcc.Dropdown(id="name-column-1", placeholder="Select Name Column", style={"width": "80%", "margin": "auto"}), width=6),
                            dbc.Col(dcc.Dropdown(id="email-column-1", placeholder="Select Email Column", style={"width": "80%", "margin": "auto"}), width=6)
                        ], className="mb-3", style={"justifyContent": "center"}),
                        dbc.Row([
                            dbc.Col(dcc.Dropdown(id="name-column-2", placeholder="Select Name Column", style={"width": "80%", "margin": "auto"}), width=6),
                            dbc.Col(dcc.Dropdown(id="email-column-2", placeholder="Select Email Column", style={"width": "80%", "margin": "auto"}), width=6)
                        ], className="mb-3", style={"justifyContent": "center"})
                    ],
                    style={
                        "textAlign": "center", "padding": "15px", "backgroundColor": "#f8f9fa", "border": "1px solid #ddd",
                        "borderRadius": "8px", "boxShadow": "2px 2px 10px rgba(0, 0, 0, 0.1)", "width": "50%", "margin": "auto"
                    }
                ),
                width=12
            )
        ),

    # Email Sending Section
        # Email Sender Information Section
    dbc.Row(
        dbc.Col(
            html.Div(
                [
                    html.H2("Email Sender Information", style={
                        "fontSize": "22px", "fontWeight": "bold", "textAlign": "center", "color": "#333",
                        "padding": "10px", "backgroundColor": "#f8f9fa", "borderRadius": "8px"
                    }),
                    dbc.Row([
                        dbc.Col(
                            html.Div([
                                html.Label("Sender's Full Name", style={"fontWeight": "bold"}),
                                dbc.Input(id="sender-name", type="text", placeholder="Enter Full Name", className="form-control")
                            ]), width=6
                        ),
                        dbc.Col(
                            html.Div([
                                html.Label("Sender's Email Address", style={"fontWeight": "bold"}),
                                dbc.Input(id="sender-email", type="email", placeholder="Enter Email Address", className="form-control")
                            ]), width=6
                        )
                    ], className="mb-3", style={"padding": "10px"}),
                    dbc.Row([
                        dbc.Col(
                            html.Div([
                                html.Label([
                                    "App Password (for SMTP) ",
                                    html.I(className="fas fa-info-circle", id="password-info-icon", style={"cursor": "pointer", "color": "#007BFF", "marginLeft": "5px"})
                                ], style={"fontWeight": "bold"}),
                                dbc.Tooltip(
                                    "An app password is a 16-digit passcode that gives less secure apps or devices permission to access your Google Account. You can generate it in your Google account security settings.",
                                    target="password-info-icon",
                                    placement="right"
                                ),
                                dbc.InputGroup([
                                    dbc.Input(id="sender-password", type="password", placeholder="Enter App Password", className="form-control"),
                                    dbc.Button("👁", id="toggle-password", n_clicks=0, color="secondary", outline=True)
                                ])
                            ]), width=6
                        ),
                        dbc.Col(
                            html.Div([
                                html.Label("Company Name", style={"fontWeight": "bold"}),
                                dbc.Input(id="company-name", type="text", placeholder="Enter Company Name", className="form-control")
                            ]), width=6
                        )
                    ], className="mb-3", style={"padding": "10px"})
                ],
                style={
                    "textAlign": "center", "padding": "15px", "backgroundColor": "#ffffff", "border": "1px solid #ddd",
                    "borderRadius": "8px", "boxShadow": "2px 2px 10px rgba(0, 0, 0, 0.1)", "width": "50%", "margin": "auto"
                }
            ),
            width=12
        )
    ),
        # Email Subject & Template Section
    dbc.Row(
        dbc.Col(
            html.Div(
                [
                    html.H2("Email Content Configuration", style={
                        "fontSize": "22px", "fontWeight": "bold", "textAlign": "center", "color": "#333",
                        "padding": "10px", "backgroundColor": "#f8f9fa", "borderRadius": "8px"
                    }),
                    
                    dbc.Row([
                        dbc.Col(
                            html.Div([
                                html.Label("Email Subject", style={"fontWeight": "bold"}),
                                dbc.Input(id="email-subject", type="text", placeholder="Enter Email Subject", className="form-control")
                            ]), width=12
                        )
                    ], className="mb-3", style={"padding": "10px"}),
                    
                    dbc.Row([
                        dbc.Col(
                            html.Div([
                                html.Label("Email Template", style={"fontWeight": "bold"}),
                                dbc.Textarea(
                                    id="email-template",
                                    placeholder="Enter email template with {employee_name}, {company_name}, {designation}, {sender_name}",
                                    className="form-control",
                                    rows=6,
                                    value="""
Dear {employee_name},

Greetings from {company_name}!

We are reaching out to you regarding your role as {designation}. 
Please review the attached document and let us know if you need any further details.

Best regards,  
{sender_name}  
{company_name}
"""
                                )
                            ]), width=12
                        )
                    ], className="mb-3", style={"padding": "10px"}),
                    
                    dbc.Row(
                        dbc.Col(
                            html.Div([
                                html.Button("Send Emails", id="send-email", className="btn btn-success", style={"width": "100%", "fontSize": "18px", "padding": "10px"})
                            ], style={"textAlign": "center", "marginTop": "10px"}),
                            width=12
                        )
                    ),
                    
                    html.Div(id="email-status", className="mt-3", style={"textAlign": "center", "fontSize": "16px", "fontWeight": "bold", "color": "#28a745"})
                ],
                style={
                    "textAlign": "center", "padding": "15px", "backgroundColor": "#ffffff", "border": "1px solid #ddd",
                    "borderRadius": "8px", "boxShadow": "2px 2px 10px rgba(0, 0, 0, 0.1)", "width": "60%", "margin": "auto"
                }
            ),
            width=12
        )
    )
], fluid=True)

# ✅ Show the correct layout (License Key Prompt or Main App)
if check_license():
    app.layout = app_layout
else:
    app.layout = license_layout

# ✅ Callback to verify and save the license key
@app.callback(
    [Output("license-status", "children"),
     Output("license-status", "style"),
     Output("activate-btn", "disabled"),
     Output("redirect-script", "children")],  # ✅ JavaScript trigger for page refresh
    Input("activate-btn", "n_clicks"),
    State("license-key", "value"),
)
def activate_license(n_clicks, entered_key):
    if n_clicks and entered_key:
        if entered_key in VALID_KEYS:
            save_license(entered_key)  # ✅ Save the key
            return (
                "✅ License Activated! Restarting...",
                {"color": "green"},
                True,
                html.Script("window.location.reload();")  # ✅ Forces a full page reload
            )
        else:
            return (
                "❌ Invalid License Key! Try Again.",
                {"color": "red"},
                False,
                dash.no_update  # ❌ Prevents page reload if the key is wrong
            )
    return "", {}, False, dash.no_update


# Function to Read Excel File
def parse_contents(contents):
    content_type, content_string = contents.split(",")
    decoded = base64.b64decode(content_string)
    return pd.read_excel(io.BytesIO(decoded))

# Callback to display uploaded file name
@app.callback(
    Output("uploaded-file-name", "children"),
    Input("upload-excel", "filename")
)
def update_filename(filename):
    if filename:
        return f"Uploaded File: {filename}"
    return ""

# Callback to Toggle Password Visibility
@app.callback(
    Output("sender-password", "type"),
    Input("toggle-password", "n_clicks"),
    State("sender-password", "type")
)
def toggle_password_visibility(n_clicks, password_type):
    return "text" if password_type == "password" else "password"

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

