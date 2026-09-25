import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os
import subprocess
import ollama
import json
from datetime import datetime

expense_words = [
    "bought",
    "buy",
    "buying",
    "purchased",
    "purchase",
    "paid",
    "pay",
    "paying",
    "spent",
    "spend",
    "spending",
    "invested",
    "investing",
    "investment",
    "obtained",
    "obtaining",
    "obtain",
    "financed",
    "financing",
    "to finance",
    "repurchased",
    "repurchasing",
    "charge",
    "charges",
    "charged"
]

income_words = [
    "got paid"
    "sold",
    "sell",
    "selling",
    "received",
    "receive",
    "receiving",
    "earned",
    "earn",
    "earning",
    "collected",
    "collect",
    "collecting",
    "distribute",
    "distributed",
    "distributing",
    "deposit",
    "depositing",
    "deposited",
    "given",
    "gave me"
]

# The spreadsheet Phillip is currently using
current_file = "accounting.xlsx"

ai_name = "Phillip"

template_options = [
    "Quarter Accounting",
    "Personal Taxes",
    "Personal Budget",
    "Sales Tracking",
    "Checking",
    "Custom"
]

template_columns = {

    "Quarter Accounting": [
        "Date",
        "Vendor/Client",
        "Description",
        "Quantity",
        "Category",
        "Type",
        "Amount"
    ],

    "Personal Taxes": [
        "Date",
        "Payee",
        "Description",
        "Tax Category",
        "Deductible",
        "Amount"
    ],

    "Personal Budget": [
        "Date",
        "Description",
        "Category",
        "Type",
        "Amount"
    ],

    "Sales Tracking": [
        "Date",
        "Customer",
        "Product",
        "Quantity",
        "Unit Price",
        "Total"
    ],

    "Checking": [
        "Date",
        "Description",
        "Category",
        "Direction",
        "Payment Method",
        "Debit",
        "Credit",
        "Cash",
        "Total"
    ],

    "Custom": [
        "Date",
        "Description",
        "Category",
        "Amount"
    ]
}

template_ai_instructions = {

    "Quarter Accounting": """
Return exactly these fields:
vendor, description, quantity, category, type, amount

Rules:
- vendor = company, person, or customer involved
- description = what was purchased, sold, or paid for
- quantity = number of items, or 1 if not stated
- category = appropriate accounting category
- type = ONLY "Income" or "Expense"
- amount = numerical amount only
- If a numerical value is not provided, return 0, never "none" or "null"
""",

    "Personal Taxes": """
Return exactly these fields:
payee, description, tax_category, deductible, amount

Rules:
- payee = person or company paid
- description = what the transaction was for
- tax_category = appropriate tax category
- deductible = ONLY true or false
- amount = numerical amount only
- If a numerical value is not provided, return 0, never "none" or "null"
""",

    "Personal Budget": """
Return exactly these fields:
description, category, type, amount

Rules:
- description = what the transaction was for
- category = appropriate personal budget category
- type = ONLY "Income" or "Expense"
- amount = numerical amount only
- If a numerical value is not provided, return 0, never "none" or "null"
""",

    "Sales Tracking": """
Return exactly these fields:
customer, product, quantity, unit_price, total

Rules:
- customer = customer who purchased the product
- product = product or service sold
- quantity = number sold
- unit_price = price per item
- total = total sale amount
- If a numerical value is not provided, return 0, never "none" or "null"
""",

    "Checking": """
Return exactly these fields:
description, category, direction, debit, credit, balance

Rules:
- description = description of the transaction
- category = appropriate checking-account category
- direction = ONLY "Income" or "Expense"
- payment method = ONLY "Debit", "Credit" or "Cash"

For Income transactions:
- paycheck, wages, salary, or direct deposit from an employer = "Paycheck"
- money earned from a business = "Business Income"
- interest earned from a bank account = "Interest"
- dividends or investment earnings = "Investment Income"
- tax refunds, returned purchases, or other refunds = "Refund"
- money received as a gift = "Gift"
- transfers from another account = "Transfer"
- other unidentified income = "Other Income"

For Expense transactions:
- groceries = "Groceries"
- rent or mortgage = "Housing"
- utilities = "Utilities"
- restaurants or eating out = "Dining"
- fuel or gas = "Transportation"
- purchases of clothing = "Clothing"
- medical or healthcare expenses = "Healthcare"
- entertainment = "Entertainment"
- subscriptions = "Subscriptions"
- other unidentified expenses = "Other Expense"

- debit = label debit if said money entered or exitted the debit card, otherwise 0
- credit = label credit if said money entered or exitted the credit card, otherwise 0
- cash = numerical cash amount involved, otherwise 0
- total = added activity from 'debit', 'credit', and 'cash', income=+ and expense=-, otherwise 0
- If a numerical value is not provided, return 0, never "none" or "null"
""",

    "Custom": """
Return exactly these fields:
description, category, amount

Rules:
- description = description of the transaction
- category = appropriate category
- amount = numerical amount only
- If a numerical value is not provided, return 0, never "none" or "null"
"""
}

def create_new_spreadsheet():
    global current_file

    filename = new_file_entry.get().strip()

    if not filename:
        messagebox.showwarning(
            "Phillip",
            "Please enter a name for the spreadsheet."
        )
        return

    if not filename.endswith(".xlsx"):
        filename += ".xlsx"

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Account Log"

    # Get the selected template
    template = selected_template.get()

    # Get the columns for that template
    columns = template_columns[template]

    # Create the headers
    sheet.append(columns)

    workbook.save(filename)

    # Make the new spreadsheet Phillip's current spreadsheet
    current_file = os.path.abspath(filename)

    messagebox.showinfo(
        "Phillip",
        f"New spreadsheet created:\n\n{filename}"
    )

    new_file_entry.delete(0, tk.END)

    status_label.config(
        text=f"Current spreadsheet: {filename}",
        font=("Times New Roman", 11)
    )

    # Refresh spreadsheet list
    refresh_spreadsheet_list()

def safe_float(value, default=0):
    if value is None:
        return default

    if isinstance(value, str):
        value = value.strip().lower()

        if value in ["none", "null", "n/a", "na", ""]:
            return default

        value = value.replace("$", "").replace(",", "")

    try:
        return float(value)
    except (ValueError, TypeError):
        return default

def detect_payment_method(transaction, ai_method="Cash"):
    text = transaction.lower()

    debit_terms = [
        "debit",
        "debit card",
        "bank card",
        "checking card",
        "greenlight",
        "direct"
    ]

    credit_terms = [
        "credit",
        "credit card",
        "visa",
        "mastercard",
        "amex",
        "american express",
        "discover",
        "check"
    ]

    cash_terms = [
        "cash",
        "in cash",
        "paid cash",
        "cash payment",
        "with change",
        "dollar bills",
        "coins"
    ]

    if any(term in text for term in debit_terms):
        return "Debit"

    if any(term in text for term in credit_terms):
        return "Credit"

    if any(term in text for term in cash_terms):
        return "Cash"

    if ai_method in ["Debit", "Credit", "Cash"]:
        return ai_method

    return "Cash"

def add_transaction():
    global current_file

    transaction = entry.get().strip()

    if not transaction:
        messagebox.showwarning(
            "Phillip",
            "Please enter a transaction."
        )
        return

    template = selected_template.get()

    try:
        response = ollama.chat(
            model="llama3.2:3b",
            messages=[
                {
                    "role": "system",
                    "content": f"""
You are Phillip, an accounting assistant.

The user has selected the spreadsheet template:
{template}

Analyze the user's transaction according to the rules for that template.

{template_ai_instructions[template]}

Return ONLY valid JSON.

Do not include explanations outside the JSON.
"""
                },
                {
                    "role": "user",
                    "content": transaction
                }
            ],
            format="json"
        )

        ai_text = response["message"]["content"]

        # Convert the JSON response into Python data
        data = json.loads(ai_text)

    except Exception as e:
        messagebox.showerror(
            "Phillip Error",
            f"Phillip could not process the transaction:\n\n{e}"
        )
        return

    # Automatically get today's date
    date = datetime.now().strftime("%m-%d-%Y")

    try:

        # ---------------------------------
        # QUARTER ACCOUNTING
        # ---------------------------------

        if template == "Quarter Accounting":

            vendor = data.get("vendor", "N/A")
            description = data.get("description", "N/A")
            quantity = data.get("quantity", 1)
            category = data.get("category", "Uncategorized")
            amount = safe_float(data.get("amount", 0))

            transaction_lower = transaction.lower()

            if any(word in transaction_lower for word in expense_words):
                transaction_type = "Expense"

            elif any(word in transaction_lower for word in income_words):
                transaction_type = "Income"

            else:
                transaction_type = data.get("type", "Expense")

            if transaction_type not in ["Income", "Expense"]:
                transaction_type = "Expense"

            row = [
                date,
                vendor,
                description,
                quantity,
                category,
                transaction_type,
                amount
            ]

            currency_columns = [7]


        # ---------------------------------
        # PERSONAL TAXES
        # ---------------------------------

        elif template == "Personal Taxes":

            payee = data.get("payee", "N/A")
            description = data.get("description", "N/A")
            tax_category = data.get(
                "tax_category",
                "Uncategorized"
            )
            deductible = data.get("deductible", False)
            amount = safe_float(data.get("amount", 0))

            row = [
                date,
                payee,
                description,
                tax_category,
                deductible,
                amount
            ]

            currency_columns = [6]


        # ---------------------------------
        # PERSONAL BUDGET
        # ---------------------------------

        elif template == "Personal Budget":

            description = data.get("description", "N/A")
            category = data.get(
                "category",
                "Uncategorized"
            )
            transaction_type = data.get(
                "type",
                "Expense"
            )
            amount = safe_float(data.get("amount", 0))

            if transaction_type not in ["Income", "Expense"]:
                transaction_type = "Expense"

            row = [
                date,
                description,
                category,
                transaction_type,
                amount
            ]

            currency_columns = [5]


        # ---------------------------------
        # SALES TRACKING
        # ---------------------------------

        elif template == "Sales Tracking":

            customer = data.get("customer", "N/A")
            product = data.get("product", "N/A")
            quantity = data.get("quantity", 1)

            unit_price = safe_float(data.get("unit_price", 0))
        

            total = safe_float(data.get("total", 0))
            

            row = [
                date,
                customer,
                product,
                quantity,
                unit_price,
                total
            ]

            currency_columns = [5, 6]


        # ---------------------------------
        # CHECKING
        # ---------------------------------

        elif template == "Checking":

            description = data.get(
                "description",
                "N/A"
            )

            category = data.get(
                "category",
                "Uncategorized"
            )

            direction = data.get(
                "direction",
                "expense"
            )

            payment_method = detect_payment_method(
                transaction,
                data.get("payment_method", "Cash")
            )

            debit = safe_float(
                data.get("debit", 0)
            )

            credit = safe_float(
                data.get("credit", 0)
            )

            cash = safe_float(
                data.get("cash", 0)
            )

            activity = debit + credit + cash

            if payment_method == "Debit":
                debit = activity
                credit = 0
                cash = 0

            elif payment_method == "Credit":
                credit = activity
                debit = 0
                cash = 0

            elif payment_method == "Cash":
                cash = activity
                debit = 0
                credit = 0
                        
            if direction not in ["Income", "Expense"]:
                direction = "Expense"

            # ---------------------------------
            # CALCULATE TRANSACTION ACTIVITY
            # ---------------------------------

            activity = debit + credit + cash

            # ---------------------------------
            # GET PREVIOUS TOTAL
            # ---------------------------------

            if os.path.exists(current_file):

                existing_workbook = load_workbook(current_file)
                existing_sheet = existing_workbook["Account Log"]

                if existing_sheet.max_row > 1:

                    previous_total = safe_float(
                        existing_sheet.cell(
                            row=existing_sheet.max_row,
                            column=9
                        ).value
                    )

                else:

                    previous_total = 0

                existing_workbook.close()

            else:

                previous_total = 0

            # ---------------------------------
            # UPDATE RUNNING TOTAL
            # ---------------------------------

            if direction == "Income":

                total = previous_total + activity

            else:

                total = previous_total - activity

            # ---------------------------------
            # CREATE CHECKING ROW
            # ---------------------------------


            row = [
                date,
                description,
                category,
                direction,
                payment_method,
                debit,
                credit,
                cash,
                total
            ]

            currency_columns = [6, 7, 8, 9]


        # ---------------------------------
        # CUSTOM
        # ---------------------------------

        elif template == "Custom":

            description = data.get(
                "description",
                "N/A"
            )

            category = data.get(
                "category",
                "Uncategorized"
            )

            amount = safe_float(data.get("amount", 0))
            

            row = [
                date,
                description,
                category,
                amount
            ]

            currency_columns = [4]


        else:

            raise ValueError(
                f"Unknown spreadsheet template: {template}"
            )

    except Exception as e:

        messagebox.showerror(
            "Phillip Error",
            f"Phillip could not organize the transaction:\n\n{e}"
        )

        return


    # ---------------------------------
    # OPEN THE SPREADSHEET
    # ---------------------------------

    if os.path.exists(current_file):

        workbook = load_workbook(current_file)
        sheet = workbook["Account Log"]

    else:

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Account Log"

        # Create the correct headers
        sheet.append(
            template_columns[template]
        )


    # ---------------------------------
    # ADD THE TRANSACTION
    # ---------------------------------

    sheet.append(row)


    # ---------------------------------
    # FORMAT CURRENCY CELLS
    # ---------------------------------

    for column in currency_columns:

        amount_cell = sheet.cell(
            row=sheet.max_row,
            column=column
        )

        amount_cell.number_format = '$#,##0.00'


    # Save the spreadsheet
    workbook.save(current_file)


    # Open the spreadsheet
    subprocess.run(
        ["open", current_file]
    )


    status_label.config(
        text=f"Data successfully added to {os.path.basename(current_file)}!",
        font=("Times New Roman", 11)
    )

    entry.delete(0, tk.END)
    
# -------------------------
# SEARCH SPREADSHEETS
# -------------------------

def refresh_spreadsheet_list():
    spreadsheet_list.delete(0, tk.END)

    search_text = search_entry.get().lower()

    folder = os.path.dirname(os.path.abspath(__file__))

    for filename in os.listdir(folder):
        if filename.endswith(".xlsx"):
            if search_text in filename.lower():
                spreadsheet_list.insert(tk.END, filename)


def search_spreadsheets(event=None):
    refresh_spreadsheet_list()


def select_spreadsheet():
    global current_file

    selection = spreadsheet_list.curselection()

    if not selection:
        messagebox.showwarning(
            "Phillip",
            "Please select a spreadsheet."
        )
        return

    filename = spreadsheet_list.get(selection[0])

    current_file = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        filename
    )

    status_label.config(
        text=f"Current spreadsheet: {filename}",
        font=("Times New Roman", 11)
    )

def delete_spreadsheet():
    global current_file

    selection = spreadsheet_list.curselection()

    if not selection:
        messagebox.showwarning(
            "Phillip",
            "Please select a spreadsheet to delete."
        )
        return

    filename = spreadsheet_list.get(selection[0])

    confirm = messagebox.askyesno(
        "Delete Spreadsheet",
        f"Are you sure you want to permanently delete:\n\n{filename}?"
    )

    if not confirm:
        return

    filepath = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        filename
    )

    try:
        os.remove(filepath)

        # If the deleted spreadsheet was the current spreadsheet,
        # switch back to the default spreadsheet name
        if os.path.abspath(current_file) == os.path.abspath(filepath):
            current_file = "accounting.xlsx"

        refresh_spreadsheet_list()

        status_label.config(
            text=f"Deleted spreadsheet: {filename}",
            font=("Times New Roman", 11)
        )

    except Exception as e:
        messagebox.showerror(
            "Phillip Error",
            f"Phillip could not delete the spreadsheet:\n\n{e}"
        )

# -------------------------
# WINDOW
# -------------------------

window = tk.Tk()

window.title("Phillip Accounting")
window.geometry("640x690")


title = tk.Label(
    window,
    text="PHILLIP ACCOUNTING",
    font=("Times New Roman", 23, "bold")
)

title.pack(pady=20)


selected_template = tk.StringVar(master=window)
selected_template.set("Quarter Accounting")

template_label = tk.Label(
    window,
    text="Select Spreadsheet Template:",
    font=("Times New Roman", 12)
)

template_label.pack(pady=(15, 5))


template_menu = tk.OptionMenu(
    window,
    selected_template,
    *template_options
)

template_menu.config(
    font=("Times New Roman", 11),
    width=25
)

template_menu.pack(pady=5)


instructions = tk.Label(
    window,
    text="Input New Data Here:",
    font=("Times New Roman", 15)
)

instructions.pack()


entry = tk.Entry(
    window,
    width=45,
    font=("Courier New", 12),
    bg="#F8F6F0",
    fg="black",
    highlightbackground="#808080",
    highlightcolor="#808080",    
)

entry.pack(pady=15)


submit_button = tk.Button(
    window,
    text="Enter Data",
    font=("Times New Roman", 13, "bold"),
    bg="#FFD700",
    fg="black",
    activebackground="#FFD700",
    activeforeground="black",
    highlightbackground="#FFD700",
    highlightcolor="#FFD700",    
    command=add_transaction
)

submit_button.pack(pady=10)


# -------------------------
# SPREADSHEET SEARCH
# -------------------------

search_label = tk.Label(
    window,
    text="Search Existing Spreadsheets:",
    font=("Times New Roman", 11)
)

search_label.pack(pady=(15, 5))


search_entry = tk.Entry(
    window,
    width=40,
    font=("Courier New", 12),
    bg="#F8F6F0",
    fg="black",
    highlightbackground="#808080",
    highlightcolor="#808080",    
)

search_entry.pack(pady=5)

search_entry.bind(
    "<KeyRelease>",
    search_spreadsheets
)


spreadsheet_list = tk.Listbox(
    window,
    width=35,
    height=3,
    font=("Courier New", 11)
)

spreadsheet_list.pack(pady=5)


select_button = tk.Button(
    window,
    text="Select Spreadsheet",
    font=("Times New Roman", 11),
    bg="white",
    fg="black",
    command=select_spreadsheet
)

select_button.pack(pady=5)

delete_button = tk.Button(
    window,
    text="Delete Spreadsheet",
    font=("Times New Roman", 11),
    bg="#B71C1C",
    fg="black",
    activebackground="#8B0000",
    activeforeground="black",
    highlightbackground="#B71C1C",
    highlightcolor="#B71C1C",
    command=delete_spreadsheet
)

delete_button.pack(pady=5)

# -------------------------
# NEW SPREADSHEET
# -------------------------

status_label = tk.Label(
    window,
    text=f"Current spreadsheet: {current_file}",
    font=("Times New Roman", 11)
)

status_label.pack(pady=10)


new_file_entry = tk.Entry(
    window,
    width=40,
    font=("Courier New", 12),
    bg="#F8F6F0",
    fg="black",
    highlightbackground="#808080",
    highlightcolor="#808080",    
)

new_file_entry.pack(pady=5)


new_file_button = tk.Button(
    window,
    text="Create New Spreadsheet",
    font=("Courier New", 12),
    command=create_new_spreadsheet
)

new_file_button.pack(pady=5)


# -------------------------
# FOOTER
# -------------------------

footer = tk.Label(
    window,
    text="Powered by DeBeaul Tech and Solutions",
    font=("Times New Roman", 8)
)

footer.pack(pady=10)


# Load spreadsheets when Phillip starts
refresh_spreadsheet_list()


window.mainloop()
