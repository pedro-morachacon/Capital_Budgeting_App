import json

import gspread
from flask import Flask, abort, jsonify, render_template, request
from google.oauth2.service_account import Credentials

# from googleapiclient.discovery import build

app = Flask(__name__)


def convert_to_A1_notation(start_row, start_col, end_row, end_col):
    def num_to_col_letters(num):
        letters = ""
        while num > 0:
            num, remainder = divmod(num - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters

    start_letter = num_to_col_letters(start_col)
    end_letter = num_to_col_letters(end_col)
    return f"{start_letter}{start_row}:{end_letter}{end_row}"


def rateCheck(rate):
    if rate >= 1:
        rate *= 0.01
    return round(rate, 4)


@app.route("/")
def root():
    # get db connection
    return render_template("index.html")


@app.route("/calculate", methods=["POST"])
def calculate():
    data = request.get_json()  # Assuming JSON data is sent from the frontend
    # equip.update(data)  # Update the equip dictionary with received data
    result = {"status": "success", "message": "Data received and saved successfully!"}
    equipment_name = data.get("equipment_name")
    equipment_cost = float(data.get("equipment_cost"))
    modifications_cost = float(data.get("modifications_cost"))
    installation_cost = float(data.get("installation_cost"))
    registration_cost = float(data.get("registration_cost"))
    sales_tax_amount = float(data.get("sales_tax_amount"))
    equipment_life = int(data.get("equipment_life"))
    terminal_use_years = int(data.get("terminal_use_years"))
    terminal_residual_value = float(data.get("terminal_residual_value"))
    direct_revenue = float(data.get("direct_revenue"))
    operator_wages = float(data.get("operator_wages"))
    fuel_cost = float(data.get("fuel_cost"))
    equipment_maintenance = float(data.get("equipment_maintenance"))
    other_expenses = float(data.get("other_expenses"))
    corp_tax_rate = rateCheck(float(data.get("corp_tax_rate")))
    cons_index_rate = rateCheck(float(data.get("cons_index_rate")))
    discount_rate = rateCheck(float(data.get("discount_rate")))

    # width of worksheet to add
    newSheetWidth = 7 + terminal_use_years

    # Last cashflow Column number in worksheet to add paste range
    LastCashflowColumn = 4 + terminal_use_years

    # Last PV Column number in worksheet to add paste range
    LastPVColumn = 5 + terminal_use_years

    # scope = [
    #     "https://www.googleapis.com/auth/spreadsheets",
    #     "https://www.googleapis.com/auth/drive",
    # ]
    scope = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = Credentials.from_service_account_file("creds.json", scopes=scope)

    client = gspread.authorize(creds)
    # service = build('sheets', 'v4', credentials=creds)

    # [print(client.list_spreadsheet_files(title=None, folder_id=None))]

    # [{'id': '1hCaX_ivJWCxFM-gCxCnDtLGA3ecvY1cFO637MBCBv54', 'name': 'CapitalBudgeting', 'createdTime'::17:02.422Z'}, {'id': '1AWnmXUlOas6yNY9MCB-AwwVK4qp6fmbhQqBPlQwvv8M', 'name': 'CapitalBudgeting', '2024-03-19T13:07:01.665Z'}, {'id': '1NF13rWClaOl5B5hJrEvXPHIlGWv77FwOajEDQ8DrFsA', 'name': 'CapitodifiedTime': '2024-03-19T13:05:16.176Z'}, {'id': '1DVi9WAp6Bj-R5KVzmYlU1PmL1fEKOkA7fD-xcP49SE8', 6T14:05:58.700Z', 'modifiedTime': '2024-03-19T12:50:55.105Z'}]

    sp = client.open("CapitalBudgeting")
    sh = sp.worksheet("CapitalBudget")
    # sh.batch_clear([convert_to_A1_notation(1, 7, 36, sh.column_count + 1)])
    sh.copy_range(
        convert_to_A1_notation(3, 3, 35, 3),
        convert_to_A1_notation(3, 7, 35, sh.column_count),
        paste_type="PASTE_NORMAL",
        paste_orientation="NORMAL",
    )
    sh.copy_range(
        convert_to_A1_notation(3, 6, 25, 6),
        convert_to_A1_notation(3, 7, 25, LastCashflowColumn),
        paste_type="PASTE_NORMAL",
        paste_orientation="NORMAL",
    )
    sh.copy_range(
        convert_to_A1_notation(29, 6, 35, 6),
        convert_to_A1_notation(29, 7, 35, LastPVColumn),
        paste_type="PASTE_NORMAL",
        paste_orientation="NORMAL",
    )

    sh.batch_clear([convert_to_A1_notation(1, LastPVColumn, 25, LastPVColumn)])

    print(f"Capital Budgeting for {equipment_name} Example")
    #  Page Title
    # --------------------------
    sh.batch_update(
        [
            {
                "range": "A1",
                "values": [[f"Capital Budgeting for {equipment_name} Example"]],
            },
            {
                "range": "B4:B9",
                "values": [
                    [equipment_life],
                    [terminal_use_years],
                    [terminal_residual_value],
                    [discount_rate],
                    [cons_index_rate],
                    [corp_tax_rate],
                ],
            },
            {
                "range": "B13:B17",
                "values": [
                    [equipment_cost],
                    [modifications_cost],
                    [installation_cost],
                    [sales_tax_amount],
                    [registration_cost],
                ],
            },
            {"range": "E4", "values": [[direct_revenue]]},
            {
                "range": "E7:E10",
                "values": [
                    [operator_wages],
                    [fuel_cost],
                    [equipment_maintenance],
                    [other_expenses],
                ],
            },
        ]
    )
    result = {"status": "success", "message": "Data received and saved successfully!"}
    return jsonify(result)


# To run flask function.
if __name__ == "__main__":
    app.run(debug=True)
