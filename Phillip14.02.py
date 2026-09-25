import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.chart import PieChart, LineChart, Reference
import os
import subprocess
import sys
import math
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
    "charged",
    "donated",
    "donation",
    "donating",
    "refunding",
    "refunded"
]

income_words = [
    "got paid",
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
    "gave me",
    "refund"
]

# Planetary gravity values (m/s²)
planet_gravity = {
    "Earth": 9.81,
    "Moon / Luna": 1.62,
    "Mercury": 3.70,
    "Venus": 8.87,
    "Mars": 3.71,
    "Jupiter": 24.79,
    "Saturn": 10.44,
    "Uranus": 8.69,
    "Neptune": 11.15,
    "Pluto": 0.62,
    "Sun": 274.00,
    "Proxima Centauri b": 11.00,
    "Halley's Comet": 0.001,
    "Miller's Planet": 12.75,
    "51 Pegasi b / Dimidium": 7.9,
    "WASP-76b": 6.4,
    "WASP-121b": 9.33,
    "WASP-107b": 2.7, 
    "Titan": 1.35,
    "Phobos": 0.0057,
    "Deimos": 0.003,
    "Io": 1.796,
    "Ganymede": 1.428,
    "Europa": 1.315,
    "Callisto": 1.235,
    "Rhea": 0.264,
    "Dione": 0.232,
    "Iapetus": 0.223,
    "Tethys": 0.145,
    "Enceladus": 0.113,
    "Mimas": 0.064,
    "Titania": 0.367,
    "Oberon": 0.346,
    "Ariel": 0.267,
    "Umbriel": 0.228,
    "Miranda": 0.079,
    "Triton": 0.779,
    "Charon": 0.288,
    "Naiad": 0.0012,
    "Thalassa": 0.0016,
    "Despina": 0.0036,
    "Galatea": 0.0044,
    "Larissa": 0.016,
    "Hippocamp": 0.0003,
    "Proteus": 0.07,
    "Nereid": 0.013,
    "Halimede": 0.0014,
    "Sao": 0.0011,
    "S/2002 N 5": 0.0007,
    "Laomedeia": 0.0011,
    "Psamathe": 0.0007,
    "Neso": 0.0014,
    "S/2021 N 1": 0.0004,
    "Bianca": 0.016,
    "Cressida": 0.026,
    "Desdemona": 0.023,
    "Juliet": 0.031,
    "Portia": 0.046,
    "Rosalind": 0.046,
    "Cupid": 0.003,
    "Belinda": 0.028,
    "Perdita": 0.009,
    "Puck": 0.028,
    "Mab": 0.002,
    "Francisco": 0.003,
    "Caliban": 0.02,
    "Stephano": 0.009,
    "Trinculo": 0.005,
    "Sycorax": 0.047,
    "Margaret": 0.005,
    "Prospero": 0.015,
    "Setebos": 0.014,
    "Ferdinand": 0.003,
    "S/2023 U 1": 0.001,
    "Surtur": 0.0001,
    "Pan": 0.006,
    "Daphnis": 0.002,
    "Atlas": 0.002,
    "Prometheus": 0.006,
    "Pandora": 0.005,
    "Epimetheus": 0.011,
    "Janus": 0.014,
    "Methone": 0.0003,
    "Pallene": 0.0003,
    "Telesto": 0.004,
    "Helene": 0.005,
    "Polydeuces": 0.0003,
    "Hyperion": 0.02,
    "Phoebe": 0.05,
    "Skoll": 0.001,
    "Mundilfari": 0.001,
    "Suttungr": 0.001,
    "Metis": 0.016,
    "Adrastea": 0.01,
    "Amalthea": 0.02,
    "Thebe": 0.013,
    "Valetudo": 0.0003,
    "Carpo": 0.0003,
    "Hydra": 0.051,
    "Nix": 0.043,
    "Kerberos": 0.015,
    "Styx": 0.011,
    "Hi'laka": 0.067,
    "Namaka": 0.012,
    "Dysnomia": 0.08,
    "Makemake 2": 0.012,
    "Ceres": 0.27,
    "Haumea": 0.4,
    "Makemake 1": 0.5,
    "Eris": 0.82,
    "Gonggong": 0.3,
    "Quaoar": 0.2,
    "Sedna": 0.3,
    "Kepler-22b": 23.5,
    "LHS 1140 b": 18.4,
    "TOI-700": 17.6,
    "K2-18b": 12.1,
    "HR 5183 b": 65.00,
    "KELT-9b": 20.8
}

selected_planet = "Earth"

def rounded_button(
    parent,
    text,
    command,
    bg=("white"),
    fg=("black"),
    width=220,
    height=45,
    font=("Arial", 11, "bold"),
    corner_radius=12
):
    canvas = tk.Canvas(
        parent,
        width=width,
        height=height,
        highlightthickness=0,
        bg=parent.cget("bg")
    )

    r = corner_radius

    # Main rectangle
    canvas.create_rectangle(
        r,
        0,
        width - r,
        height,
        fill=bg,
        outline=bg
    )

    canvas.create_rectangle(
        0,
        r,
        width,
        height - r,
        fill=bg,
        outline=bg
    )

    # Four rounded corners
    canvas.create_arc(
        0,
        0,
        r * 2,
        r * 2,
        start=90,
        extent=90,
        fill=bg,
        outline=bg
    )

    canvas.create_arc(
        width - r * 2,
        0,
        width,
        r * 2,
        start=0,
        extent=90,
        fill=bg,
        outline=bg
    )

    canvas.create_arc(
        0,
        height - r * 2,
        r * 2,
        height,
        start=180,
        extent=90,
        fill=bg,
        outline=bg
    )

    canvas.create_arc(
        width - r * 2,
        height - r * 2,
        width,
        height,
        start=270,
        extent=90,
        fill=bg,
        outline=bg
    )

    # Button text
    canvas.create_text(
        width / 2,
        height / 2,
        text=text,
        fill=fg,
        font=font
    )

    # Click
    canvas.bind(
        "<Button-1>",
        lambda event: command()
    )

    return canvas

def launch_newton(v0, angle, gravity, height, planet):
    newton_path = "/Users/jaxondegen/KinematicsAI/Newton3.py"

    print("Newton path:", newton_path)
    print("Newton exists:", os.path.exists(newton_path))

    subprocess.Popen([
        sys.executable,
        newton_path,
        str(v0),
        str(angle),
        str(gravity),
        str(height),
        str(planet)
    ])

    print("Phillip location:", os.path.dirname(os.path.abspath(__file__)))

def open_planet_selector():
    planet_window = tk.Toplevel(window)
    planet_window.title("Select a Celestial Entity")
    planet_window.geometry("270x200")
    planet_window.resizable(False, False)
    planet_window.configure(bg="#1F123E")

    #########################
    #Font and Wording
    tk.Label(
        planet_window,
        text="Select a Celestial Entity",
        font=("Menlo", 12, "bold"),
        fg="#0096FF",
        bg="#1F123E"
    ).pack(pady=15)

    selected_planet_var = tk.StringVar(value=selected_planet)

    planet_menu = tk.OptionMenu(
        planet_window,
        selected_planet_var,
        *planet_gravity.keys()
    )

    planet_menu.config(
        width=20,
        font=("Monaco", 10)
    )

    planet_menu.pack(pady=9)

    gravity_label = tk.Label(
        planet_window,
        text="Gravity: 9.81 m/s²",
        font=("Menlo", 12),
        bg="#1F123E"
    )

    gravity_label.pack(pady=7)
    

    def update_gravity(*args):
        global selected_planet

        planet = selected_planet_var.get()
        gravity = planet_gravity[planet]

        gravity_label.config(
            text=f"Gravity: {gravity} m/s²"
        )

    selected_planet_var.trace_add("write", update_gravity)
        
    def apply_planet():
        global selected_planet

        selected_planet = selected_planet_var.get()

        planet_window.destroy()


    rounded_button(
        planet_window,
        text="Apply Gravity",
        command=apply_planet,
        bg="#F5FEFD",
        fg="#0096FF",
        font=("Menlo", 11, "bold"),
        width=130,
        height=30,
        corner_radius=12
    ).pack(pady=8)

# The spreadsheet Phillip is currently using
DATA_FOLDER = os.path.expanduser("~/Documents/Phillip Accounting")

os.makedirs(DATA_FOLDER, exist_ok=True)

current_file = os.path.join("Welcome.xlsx")

ai_name = "Phillip"

def check_ai_engine():
    try:
        import ollama
        ollama.list()
        return True
    except Exception:
        return False

template_options = [
    "Quarter Accounting",
   # "Personal Taxes",
   # "Personal Budget",
    "Sales Tracking",
    "Checking",
    "1D Kinematics Calculator",
    "2D Kinematics Calculator",
   # "Custom"
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

   # "Personal Taxes": [
   #     "Date",
   #     "Payee",
   #     "Description",
   #     "Tax Category",
   #     "Deductible",
   #     "Amount"
   # ],

   # "Personal Budget": [
   #     "Date",
   #     "Description",
   #     "Category",
   #     "Type",
   #     "Amount"
   # ],

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


    "1D Kinematics Calculator": [
        "Trial",
        "Object",
        "X or Y",
        "Initial Position",
        "End Position",
        "Initial Velocity",
        "End Velocity",
        "Acceleration",
        "Displacement",
        "Time"
    ],

    "2D Kinematics Calculator": [
        "Trial",
        "Object",
        "Initial Height",
        "Initial Position",
        "Max Height",
        "End Position",
        "Displacement",
        "Acceleration",
        "Vertical Velocity",
        "Horizontal Velocity",
        "Angle",
        "Speed",
        "Air Time",
        "Celestial Body"
    ],


   # "Custom": [
   #     "Date",
   #     "Description",
   #     "Category",
   #     "Amount"
   # ]
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
description, category, direction, payment_method, debit, credit, cash

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
- Venmo, Cash App, PayPal, Zelle, Apple Pay, Apple Cash, Google Wallet, Google Pay, Samsung Pay, Chime, Revolut, Snapcash, Meta Pay, or Splitwise transactions = "Digital Wallets"
- retirement benefits, disability benefits, stipends, unemployment benefits, social security benefits, and government stiumulus = "Benefits"
- Child Support, Alimony, and other lawsuit claims = "Law Claim"
- other unidentified income = "Other Income"

For Expense transactions:
- groceries = "Groceries"
- rent, mortgage, or purchase of house or apartment = "Housing"
- purchase of luxury car, boat, atv, jetski, jet, plane, etc. = "Luxury Vehicle"
- purchase of a regular car, motorcycle, or bike = "Standard Vehicle"
- books, tuition, or school supplies = "Academic Costs"
- utilities = "Utilities"
- restaurants or eating out = "Dining"
- fuel, gas, plane ticket, bus ticket, train ticket, cruise ticket, or ferry ticket = "Transportation"
- hotel fees or car rental = "Vacation"
- purchases of clothing = "Clothing"
- medical or healthcare expenses = "Healthcare"
- purchases online = "Digital Purchases"
- entertainment = "Entertainment"
- subscriptions = "Subscriptions"
- taxes = "Taxes"
- donation = "Donation"
- Venmo, Cash App, PayPal, Zelle, Apple Pay, Apple Cash, Google Wallet, Google Pay, Samsung Pay, Chime, Revolut, Snapcash, Meta Pay, or Splitwise transactions = "Digital Wallets"
- other unidentified expenses = "Other Expense"

- debit = numerical amount involved with a debit card, otherwise 0
- credit = numerical amount involved with a credit card, otherwise 0
- cash = numerical amount involved with cash, otherwise 0
- Do not calculate total
- If a numerical value is not provided, return 0, never "none" or "null"
""",

    "1D Kinematics Calculator": """
Return exactly these fields:

trial
object
axis
initial_position
end_position
initial_velocity
end_velocity
acceleration
displacement
time

Rules:

- trial = trial number if explicitly stated, otherwise 1
- object = object being analyzed
- axis = ONLY "X" or "Y"

POSITION:
- initial_position = starting position in meters
- end_position = final position in meters
- If the user says "starts at 10 meters", that means initial_position = 10
- If the user says "ends at 50 meters", that means end_position = 50
- Do NOT confuse position with displacement

DISPLACEMENT:
- displacement = change in position
- displacement = end_position - initial_position
- If the user gives a distance traveled or displacement directly, put it in displacement
- Do NOT put displacement into end_position unless the user explicitly says it is the final position

VELOCITY:
- initial_velocity = starting velocity in m/s
- end_velocity = final velocity in m/s

ACCELERATION:
- acceleration = acceleration in m/s²
- Negative acceleration is allowed

TIME:
- time = elapsed time in seconds

IMPORTANT:
- Return null for values that are not provided or cannot be determined directly from the information given.
- Do not perform physics calculations.
- Convert units to meters, m/s, m/s², and seconds.
- Return only valid JSON.
""",

"2D Kinematics Calculator": """
Return exactly these fields:

trial
object
initial_height
initial_position
max_height
end_position
displacement
acceleration
vertical_velocity
horizontal_velocity
angle
speed
air_time

Rules:

TRIAL:
- trial = trial number if explicitly stated, otherwise 1

OBJECT:
- object = object being analyzed

POSITION:
- initial_height = starting vertical position in meters
- initial_position = starting horizontal position in meters
- If the user does not mention an initial height, return null
- If the user does not mention an initial horizontal position, return null
- Do not confuse height with horizontal position

MAX HEIGHT:
- max_height = maximum vertical height in meters
- Only return it if explicitly provided
- Do not calculate it

END POSITION:
- end_position = final horizontal position in meters
- Do not confuse end position with maximum height
- Do not put displacement into end_position unless the user explicitly describes it as the final position

DISPLACEMENT:
- displacement = horizontal change in position in meters
- If the user explicitly gives a displacement, put it here
- Do not confuse displacement with initial position or end position

ACCELERATION:
- acceleration = acceleration in m/s²
- Negative acceleration is allowed
- For ordinary projectile motion near Earth, the acceleration may be provided as -9.8 m/s²

VELOCITY:
- vertical_velocity = vertical component of velocity in m/s
- horizontal_velocity = horizontal component of velocity in m/s
- Do not combine the two components into one value

ANGLE:
- angle = launch angle in degrees
- Return only the numerical angle
- Do not convert the angle to radians

SPEED:
- speed = initial launch speed in m/s
- Return null if the speed is not explicitly provided

AIR TIME:
- air_time = total time in the air in seconds
- Return null if it is not explicitly provided

IMPORTANT:
- Return null for values that are not provided or cannot be determined directly from the information given.
- Do not perform physics calculations.
- Convert units to meters, m/s, m/s², and seconds.
- Return only valid JSON.
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

    filepath = os.path.join(DATA_FOLDER, filename)
    workbook.save(filepath)

    # Make the new spreadsheet Phillip's current spreadsheet
    current_file = filepath

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

def calculate_1d_kinematics(data):
    """
    Solve a 1D kinematics problem.

    Standard equations:
        vf = vi + a*t
        d = vi*t + 0.5*a*t^2
        vf^2 = vi^2 + 2*a*d
        d = ((vi + vf) / 2) * t
        xf = xi + d
    """

    def get_value(name):
        value = data.get(name)

        if value is None:
            return None

        if isinstance(value, str):
            value = value.strip()

            if value.lower() in ["none", "null", ""]:
                return None

        try:
            return float(value)
        except (ValueError, TypeError):
            return None

    xi = get_value("initial_position")
    xf = get_value("end_position")
    vi = get_value("initial_velocity")
    vf = get_value("end_velocity")
    a = get_value("acceleration")
    d = get_value("displacement")
    t = get_value("time")

    # -----------------------------------------
    # POSITION → DISPLACEMENT
    # -----------------------------------------

    if xi is None:
        xi = 0

    if vi is None:
        vi = 0

    if d is None and a is not None and t is not None:
        d = (vi * t) + (0.5 * a * (t ** 2))

    if xf is None and d is not None:
        xf = xi + d

    if d is None and xf is not None:
        d = xf - xi
        
    # -----------------------------------------
    # VELOCITY
    # vf = vi + at
    # -----------------------------------------

    if vf is None and vi is not None and a is not None and t is not None:
        vf = vi + (a * t)

    # -----------------------------------------
    # INITIAL VELOCITY
    # vi = vf - at
    # -----------------------------------------

    if vi is None and vf is not None and a is not None and t is not None:
        vi = vf - (a * t)

    # -----------------------------------------
    # ACCELERATION
    # a = (vf - vi) / t
    # -----------------------------------------

    if a is None and vi is not None and vf is not None and t not in [None, 0]:
        a = (vf - vi) / t

    # -----------------------------------------
    # TIME
    # t = (vf - vi) / a
    # -----------------------------------------

    if t is None and vi is not None and vf is not None and a not in [None, 0]:
        t = (vf - vi) / a

    # -----------------------------------------
    # DISPLACEMENT
    # d = vi*t + 1/2*a*t²
    # -----------------------------------------

    if d is None and vi is not None and a is not None and t is not None:
        d = (vi * t) + (0.5 * a * (t ** 2))

    # -----------------------------------------
    # DISPLACEMENT
    # d = ((vi + vf) / 2)t
    # -----------------------------------------

    if d is None and vi is not None and vf is not None and t is not None:
        d = ((vi + vf) / 2) * t

    # -----------------------------------------
    # FINAL VELOCITY
    # vf² = vi² + 2ad
    # -----------------------------------------

    if vf is None and vi is not None and a is not None and d is not None:

        value = (vi ** 2) + (2 * a * d)

        if value >= 0:
            vf = value ** 0.5

    # -----------------------------------------
    # INITIAL VELOCITY
    # vi² = vf² - 2ad
    # -----------------------------------------

    if vi is None and vf is not None and a is not None and d is not None:

        value = (vf ** 2) - (2 * a * d)

        if value >= 0:
            vi = value ** 0.5

    # -----------------------------------------
    # FINAL POSITION
    # xf = xi + d
    # -----------------------------------------

    if xf is None and xi is not None and d is not None:
        xf = xi + d

    # -----------------------------------------
    # INITIAL POSITION
    # xi = xf - d
    # -----------------------------------------

    if xi is None and xf is not None and d is not None:
        xi = xf - d

    return {
        "initial_position": xi,
        "end_position": xf,
        "initial_velocity": vi,
        "end_velocity": vf,
        "acceleration": a,
        "displacement": d,
        "time": t
    }

def calculate_2d_kinematics(data, planet="Earth"):
    """
    Solve a basic 2D projectile-motion problem.

    The AI extracts the given information.
    Python performs the physics calculations.
    """

    import math

    def get_value(name):
        value = data.get(name)

        if value is None:
            return None

        if isinstance(value, str):
            value = value.strip()

            if value.lower() in ["none", "null", ""]:
                return None

        try:
            return float(value)
        except (ValueError, TypeError):
            return None

    initial_height = get_value("initial_height")
    initial_position = get_value("initial_position")
    max_height = get_value("max_height")
    end_position = get_value("end_position")
    displacement = get_value("displacement")
    acceleration = get_value("acceleration")
    vertical_velocity = get_value("vertical_velocity")
    horizontal_velocity = get_value("horizontal_velocity")
    angle = get_value("angle")
    speed = get_value("speed")
    air_time = get_value("air_time")

    # Default initial positions
    if initial_height is None:
        initial_height = 0

    if initial_position is None:
        initial_position = 0

    # Standard Earth gravity
    acceleration = -planet_gravity.get(planet, 9.81)

    # Calculate speed from velocity components
    if speed is None and vertical_velocity is not None and horizontal_velocity is not None:
        speed = math.sqrt(
            (vertical_velocity ** 2) +
            (horizontal_velocity ** 2)
        )

    # Calculate displacement from horizontal velocity and air time
    if displacement is None and horizontal_velocity is not None and air_time is not None:
        displacement = horizontal_velocity * air_time

    # Calculate end position
    if end_position is None and displacement is not None:
        end_position = initial_position + displacement

    # Calculate air time if the projectile lands at height 0
    if air_time is None and vertical_velocity is not None:

        # Solve:
        # y = y0 + vy*t + 0.5*a*t²

        a = 0.5 * acceleration
        b = vertical_velocity
        c = initial_height

        discriminant = (b ** 2) - (4 * a * c)

        if discriminant >= 0 and a != 0:

            t1 = (-b + math.sqrt(discriminant)) / (2 * a)
            t2 = (-b - math.sqrt(discriminant)) / (2 * a)

            positive_times = [
                t for t in [t1, t2]
                if t >= 0
            ]

            if positive_times:
                air_time = max(positive_times)

    # Calculate displacement once air time is known
    if displacement is None and horizontal_velocity is not None and air_time is not None:
        displacement = horizontal_velocity * air_time

    # Calculate end position once displacement is known
    if end_position is None and displacement is not None:
        end_position = initial_position + displacement

    # Calculate maximum height
    if max_height is None and vertical_velocity is not None:

        if acceleration < 0:
            max_height = (
                initial_height
                + (vertical_velocity ** 2) / (-2 * acceleration)
            )

    # Calculate angle
    # angle = arctan(height / half displacement)
    if angle is None and max_height is not None and displacement is not None:

        vertical_rise = max_height - initial_height
        half_displacement = displacement / 2

        if half_displacement != 0:
            angle = math.degrees(
                math.atan(
                    vertical_rise / half_displacement
                )
            )

        return {
            "initial_height": initial_height,
            "initial_position": initial_position,
            "max_height": max_height,
            "end_position": end_position,
            "displacement": displacement,
            "acceleration": acceleration,
            "vertical_velocity": vertical_velocity,
            "horizontal_velocity": horizontal_velocity,
            "angle": angle,
            "speed": speed,
            "air_time": air_time,
            "celestial_body": planet
        }

def detect_payment_method(transaction, ai_method="Cash"):
    text = transaction.lower()

    debit_terms = [
        "debit",
        "debit card",
        "bank card",
        "checking card",
        "greenlight",
        "direct",
        "check"
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

def automatic_visualizations(template):
    """
    Phillip automatically maintains the Checking dashboard.

    Checking Dashboard:
    1. Spending by Category pie chart
    2. Debit vs Credit vs Cash pie chart
    3. Running Total line graph with Debit/Credit/Cash activity
    """

    if not current_file or not os.path.exists(current_file):
        print("No current spreadsheet found.")
        return

    try:
        workbook = load_workbook(current_file)

        if "Account Log" not in workbook.sheetnames:
            print("Account Log sheet not found.")
            workbook.close()
            return

        sheet = workbook["Account Log"]

        # Remove the old Dashboard
        if "Dashboard" in workbook.sheetnames:
            del workbook["Dashboard"]

        dashboard = workbook.create_sheet("Dashboard")

        print("===================================")
        print("PHILLIP AUTOMATIC VISUALIZATIONS")
        print("Template:", template)
        print("Spreadsheet:", current_file)
        print("Rows:", sheet.max_row)
        print("===================================")

        # ==================================================
        # CHECKING
        # ==================================================

        if template == "Checking":

            headers = [
                cell.value
                for cell in sheet[1]
            ]

            print("Checking headers:")
            print(headers)

            required_columns = [
                "Date",
                "Description",
                "Category",
                "Direction",
                "Payment Method",
                "Debit",
                "Credit",
                "Cash",
                "Total"
            ]

            if not all(
                column in headers
                for column in required_columns
            ):

                print(
                    "Checking spreadsheet is missing required columns."
                )

                workbook.close()
                return

            # Find columns
            date_index = headers.index("Date") + 1
            category_index = headers.index("Category") + 1
            direction_index = headers.index("Direction") + 1
            payment_method_index = headers.index("Payment Method") + 1
            debit_index = headers.index("Debit") + 1
            credit_index = headers.index("Credit") + 1
            cash_index = headers.index("Cash") + 1
            total_index = headers.index("Total") + 1

            # ==================================================
            # 1. SPENDING BY CATEGORY
            # ==================================================

            category_totals = {}

            for row in range(2, sheet.max_row + 1):

                direction = str(
                    sheet.cell(
                        row=row,
                        column=direction_index
                    ).value or ""
                ).strip()

                category = str(
                    sheet.cell(
                        row=row,
                        column=category_index
                    ).value or "Other Expense"
                ).strip()

                debit = safe_float(
                    sheet.cell(
                        row=row,
                        column=debit_index
                    ).value
                )

                credit = safe_float(
                    sheet.cell(
                        row=row,
                        column=credit_index
                    ).value
                )

                cash = safe_float(
                    sheet.cell(
                        row=row,
                        column=cash_index
                    ).value
                )

                activity = debit + credit + cash

                # Only expenses belong in the spending chart
                if direction == "Expense":

                    if category not in category_totals:
                        category_totals[category] = 0

                    category_totals[category] += activity

            # Put category data on Dashboard
            dashboard["A1"] = "Category"
            dashboard["B1"] = "Amount Spent"

            category_row = 2

            for category, amount in category_totals.items():

                dashboard.cell(
                    row=category_row,
                    column=1
                ).value = category

                dashboard.cell(
                    row=category_row,
                    column=2
                ).value = amount

                dashboard.cell(
                    row=category_row,
                    column=2
                ).number_format = '$#,##0.00'

                category_row += 1

            # Create category pie chart
            if category_totals:

                category_pie = PieChart()

                category_labels = Reference(
                    dashboard,
                    min_col=1,
                    min_row=2,
                    max_row=category_row - 1
                )

                category_data = Reference(
                    dashboard,
                    min_col=2,
                    min_row=1,
                    max_row=category_row - 1
                )

                category_pie.add_data(
                    category_data,
                    titles_from_data=True
                )

                category_pie.set_categories(
                    category_labels
                )

                category_pie.title = "Spending by Category"

                category_pie.height = 8
                category_pie.width = 12

                dashboard.add_chart(
                    category_pie,
                    "D2"
                )

                print("CATEGORY PIE CHART CREATED")

            # ==================================================
            # 2. DEBIT VS CREDIT VS CASH
            # ==================================================

            debit_total = 0
            credit_total = 0
            cash_total = 0

            for row in range(2, sheet.max_row + 1):

                debit_total += safe_float(
                    sheet.cell(
                        row=row,
                        column=debit_index
                    ).value
                )

                credit_total += safe_float(
                    sheet.cell(
                        row=row,
                        column=credit_index
                    ).value
                )

                cash_total += safe_float(
                    sheet.cell(
                        row=row,
                        column=cash_index
                    ).value
                )

            payment_start_row = 1

            dashboard["D1"] = "Payment Method"
            dashboard["E1"] = "Amount"

            dashboard["D2"] = "Debit"
            dashboard["E2"] = debit_total

            dashboard["D3"] = "Credit"
            dashboard["E3"] = credit_total

            dashboard["D4"] = "Cash"
            dashboard["E4"] = cash_total

            for row in range(2, 5):

                dashboard.cell(
                    row=row,
                    column=5
                ).number_format = '$#,##0.00'

            payment_pie = PieChart()

            payment_labels = Reference(
                dashboard,
                min_col=4,
                min_row=2,
                max_row=4
            )

            payment_data = Reference(
                dashboard,
                min_col=5,
                min_row=1,
                max_row=4
            )

            payment_pie.add_data(
                payment_data,
                titles_from_data=True
            )

            payment_pie.set_categories(
                payment_labels
            )

            payment_pie.title = (
                "Debit vs Credit vs Cash"
            )

            payment_pie.height = 8
            payment_pie.width = 12

            dashboard.add_chart(
                payment_pie,
                "L2"
            )

            print("PAYMENT METHOD PIE CHART CREATED")

            # ==================================================
            # 3. TRANSACTION HISTORY
            # ==================================================

            dashboard["A15"] = "Order"
            dashboard["B15"] = "Date"
            dashboard["C15"] = "Debit"
            dashboard["D15"] = "Credit"
            dashboard["E15"] = "Cash"
            dashboard["F15"] = "Total"

            history_row = 16
            transaction_order = 1

            for row in range(2, sheet.max_row + 1):

                dashboard.cell(
                    row=history_row,
                    column=1
                ).value = transaction_order

                dashboard.cell(
                    row=history_row,
                    column=2
                ).value = sheet.cell(
                    row=row,
                    column=date_index
                ).value

                dashboard.cell(
                    row=history_row,
                    column=3
                ).value = safe_float(
                    sheet.cell(
                        row=row,
                        column=debit_index
                    ).value
                )

                dashboard.cell(
                    row=history_row,
                    column=4
                ).value = safe_float(
                    sheet.cell(
                        row=row,
                        column=credit_index
                    ).value
                )

                dashboard.cell(
                    row=history_row,
                    column=5
                ).value = safe_float(
                    sheet.cell(
                        row=row,
                        column=cash_index
                    ).value
                )

                dashboard.cell(
                    row=history_row,
                    column=6
                ).value = safe_float(
                    sheet.cell(
                        row=row,
                        column=total_index
                    ).value
                )

                # Currency formatting
                for column in [3, 4, 5, 6]:

                    dashboard.cell(
                        row=history_row,
                        column=column
                    ).number_format = '$#,##0.00'

                history_row += 1
                transaction_order += 1

            # ==================================================
            # LINE GRAPH
            # ==================================================

            if sheet.max_row >= 2:

                line = LineChart()

                # Debit, Credit, Cash and Total
                line_data = Reference(
                    dashboard,
                    min_col=3,
                    max_col=6,
                    min_row=15,
                    max_row=history_row - 1
                )

                # Transaction order
                transaction_numbers = Reference(
                    dashboard,
                    min_col=1,
                    min_row=16,
                    max_row=history_row - 1
                )

                line.add_data(
                    line_data,
                    titles_from_data=True
                )

                line.set_categories(
                    transaction_numbers
                )

                line.title = (
                    "Running Total and Payment Activity"
                )

                line.y_axis.title = "Amount"
                line.x_axis.title = "Transaction Order"

                line.height = 12
                line.width = 22

                dashboard.add_chart(
                    line,
                    "L15"
                )

                print("TRANSACTION HISTORY LINE GRAPH CREATED")

        # ==================================================
        # DASHBOARD FORMATTING
        # ==================================================

        dashboard.column_dimensions["A"].width = 22
        dashboard.column_dimensions["B"].width = 18
        dashboard.column_dimensions["C"].width = 15
        dashboard.column_dimensions["D"].width = 18
        dashboard.column_dimensions["E"].width = 18
        dashboard.column_dimensions["F"].width = 18
        dashboard.column_dimensions["L"].width = 18

        # ==================================================
        # SAVE
        # ==================================================

        workbook.save(current_file)

        print("Dashboard saved successfully.")

        workbook.close()

        print("===================================")

    except Exception as e:

        print(
            "Automatic visualization error:",
            e
        )
        
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
        print("Phillip is attempting to contact Ollama...")
        print("Ollama model: llama3.2:3b")

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
        # 1D KINEMATICS CALCULATOR
        # ---------------------------------

        elif template == "1D Kinematics Calculator":

            if os.path.exists(current_file):

                existing_workbook = load_workbook(current_file)
                existing_sheet = existing_workbook["Account Log"]

                # Header is row 1
                #Therefore the first trial is 1.
                trial = existing_sheet.max_row

                existing_workbook.close()

            else:

                trial = 1
                
            object_name = data.get("object", "Object")
            axis = data.get("axis", "X")

            if axis not in ["X", "Y"]:
                axis = "X"

            results = calculate_1d_kinematics(data)

            row = [
                f"#{trial}",
                object_name,
                axis,
                f'{results["initial_position"]} m'
                    if results["initial_position"] is not None else "",
                f'{results["end_position"]} m'
                    if results["end_position"] is not None else "",
                f'{results["initial_velocity"]} m/s'
                    if results["initial_velocity"] is not None else "",
                f'{results["end_velocity"]} m/s'
                    if results["end_velocity"] is not None else "",
                f'{results["acceleration"]} m/s²'
                    if results["acceleration"] is not None else "",
                f'{results["displacement"]} m'
                    if results["displacement"] is not None else "",
                f'{results["time"]} s'
                    if results["time"] is not None else ""
            ]
            currency_columns = []

        # ---------------------------------
        # 2D KINEMATICS CALCULATOR
        # ---------------------------------

        elif template == "2D Kinematics Calculator":

            if os.path.exists(current_file):

                existing_workbook = load_workbook(current_file)
                existing_sheet = existing_workbook["Account Log"]

                # Row 1 is the header.
                # Therefore max_row gives the next trial number.
                trial = existing_sheet.max_row

                existing_workbook.close()

            else:

                trial = 1

            object_name = data.get("object", "Object")

            results = calculate_2d_kinematics(data, selected_planet)

            print("Selected planet:", selected_planet)
            print("Calculated acceleration:", results["acceleration"])

            row = [
                f"#{trial}",
                object_name,
                f'{results["initial_height"]} m'
                    if results["initial_height"] is not None else "",
                f'{results["initial_position"]} m'
                    if results["initial_position"] is not None else "",
                f'{results["max_height"]} m'
                    if results["max_height"] is not None else "",
                f'{results["end_position"]} m'
                    if results["end_position"] is not None else "",
                f'{results["displacement"]} m'
                    if results["displacement"] is not None else "",
                f'{results["acceleration"]} m/s²'
                    if results["acceleration"] is not None else "",
                f'{results["vertical_velocity"]} m/s'
                    if results["vertical_velocity"] is not None else "",
                f'{results["horizontal_velocity"]} m/s'
                    if results["horizontal_velocity"] is not None else "",
                f'{results["angle"]}°'
                    if results["angle"] is not None else "",
                f'{results["speed"]} m/s'
                    if results["speed"] is not None else "",
                f'{results["air_time"]} s'
                    if results["air_time"] is not None else "",
                results["celestial_body"]
            ]

          # Launch Newton only for the 2D Kinematics Calculator
            if results["speed"] is not None and results["angle"] is not None:
                launch_newton(
                    results["speed"],
                    results["angle"],
                    abs(results["acceleration"]),
                    results["initial_height"],
                    selected_planet
                )

            currency_columns = []

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

    print("ABOUT TO CREATE VISUALIZATIONS")
          
    # Phillip automatically analyzes the updated
    # spreadsheet and maintains its visualizations.
    automatic_visualizations(template)

    print("FINISHED CREATING VISUALIZATIONS")

    # Open the spreadsheet
    subprocess.run([
        "open",
        "-a",
        "Numbers",
        current_file
    ])


    status_label.config(
        text=f"Data successfully added to {os.path.basename(current_file)}!",
        font=("Times New Roman", 11)
    )

    entry.delete(0, tk.END)

    # ---------------------------------
    # FORMAT 1D KINEMATICS UNITS
    # ---------------------------------

    if template == "1D Kinematics Calculator":

        current_row = sheet.max_row

        #Trial: numbers
        sheet.cell(
            row=current_row,
            column=1
        ).number_format = '"#"0'

        # Position: meters
        sheet.cell(
            row=current_row,
            column=4
        ).number_format = '0.00" m"'

        sheet.cell(
            row=current_row,
            column=5
        ).number_format = '0.00" m"'

        # Velocity: meters per second
        sheet.cell(
            row=current_row,
            column=6
        ).number_format = '0.00" m/s"'

        sheet.cell(
            row=current_row,
            column=7
        ).number_format = '0.00" m/s"'

        # Acceleration: meters per second squared
        sheet.cell(
            row=current_row,
            column=8
        ).number_format = '0.00" m/s²"'

        # Displacement: meters
        sheet.cell(
            row=current_row,
            column=9
        ).number_format = '0.00" m"'

        # Time: seconds
        sheet.cell(
            row=current_row,
            column=10
        ).number_format = '0.00" s"'
    
# -------------------------
# SEARCH SPREADSHEETS
# -------------------------

def edit_transaction(results_list):
    selected = results_list.curselection()

    if not selected:
        messagebox.showinfo(
            "Edit Transaction",
            "Select a transaction from the search results first."
        )
        return

    selected_transaction = results_list.get(selected[0])

    row_number = int(
        selected_transaction.split(":")[0].replace("Row ", "")
    )

    if not os.path.exists(current_file):
        messagebox.showinfo(
            "Edit Transaction",
            "No spreadsheet is currently open."
        )
        return

    workbook = load_workbook(current_file)
    sheet = workbook.active

    headers = [
        cell.value
        for cell in sheet[1]
    ]

    values = [
        sheet.cell(row=row_number, column=column).value
        for column in range(1, len(headers) + 1)
    ]

    workbook.close()

    edit_window = tk.Toplevel(window)
    edit_window.title("Edit Transaction")
    edit_window.geometry("500x730")
    edit_window.resizable(False, True)

    tk.Label(
        edit_window,
        text="Edit Transaction",
        font=("Times New Roman", 20, "bold"),
        fg="white"
    ).pack(pady=15)

    fields_frame = tk.Frame(edit_window)
    fields_frame.pack(pady=5)

    entries = []

    for header, value in zip(headers, values):

        tk.Label(
            fields_frame,
            text=str(header),
            font=("Times New Roman", 12, "bold")
        ).pack(pady=(6, 2))

        entry = tk.Entry(
            fields_frame,
            width=40,
            font=("Courier New", 11),
            bg="#F8F6F0",
            fg="black"
        )

        entry.insert(
            0,
            "" if value is None else str(value)
        )

        entry.pack(pady=2)

        entries.append(entry)

    def save_edit():

        workbook = load_workbook(current_file)
        sheet = workbook.active

        for column, entry in enumerate(entries, start=1):
            sheet.cell(
                row=row_number,
                column=column
            ).value = entry.get()

        workbook.save(current_file)
        workbook.close()

        messagebox.showinfo(
            "Edit Transaction",
            "Transaction updated successfully."
        )

        edit_window.destroy()

    rounded_button(
        edit_window,
        text="Save Changes",
        command=save_edit,
        bg="#168FE3",
        fg="white",
        font=("Times New Roman", 12, "bold"),
        width=130,
        height=32,
        corner_radius=12
    ).pack(pady=15)

def delete_transaction(results_list):
    selected = results_list.curselection()

    if not selected:
        messagebox.showinfo(
            "Delete Transaction",
            "Select a transaction from the search results first."
        )
        return

    selected_transaction = results_list.get(selected[0])

    row_number = int(
        selected_transaction.split(":")[0].replace("Row ", "")
    )

    confirm = messagebox.askyesno(
        "Delete Transaction",
        f"Are you sure you want to delete this transaction?\n\n"
        f"{selected_transaction}"
    )

    if not confirm:
        return

    if not os.path.exists(current_file):
        messagebox.showinfo(
            "Delete Transaction",
            "No spreadsheet is currently open."
        )
        return

    workbook = load_workbook(current_file)
    sheet = workbook.active

    sheet.delete_rows(row_number, 1)

    workbook.save(current_file)
    workbook.close()

    results_list.delete(selected[0])

    messagebox.showinfo(
        "Delete Transaction",
        "Transaction deleted successfully."
    )
    
def open_search_window():
    search_window = tk.Toplevel(window)
    search_window.title("Search Transactions")
    search_window.geometry("610x425")
    search_window.resizable(False, True)
    search_window.configure(bg="#343534")


    tk.Label(
        search_window,
        text="Search Transactions",
        font=("Times New Roman", 23, "bold"),
        fg="white",
        bg="#343534"
    ).pack(pady=18)

    tk.Label(
        search_window,
        text=("Type Search Here:"),
        font=("Times New Roman", 13),
        fg="white",
        bg="#343534"
    ).pack(pady=5)

    search_entry = tk.Entry(
        search_window,
        width=27,
        font=("Courier New", 12),
        bg="#F8F6F0",
        fg="black"
    )

    search_entry.pack(pady=3)
    search_entry.focus_set()

    results_list = tk.Listbox(
        search_window,
        width=85,
        height=5,
        font=("Courier New", 8),
        bg="#b1f2ff",
        fg="black"
    )
    results_list.pack(pady=5)

    transaction_count_label = tk.Label(
        search_window,
        text="0 transactions found",
        font=("Times New Roman", 12),
        fg="white",
        bg="#343534"
    )
    transaction_count_label.pack(pady=5)

    def search_transactions():
        results_list.delete(0, tk.END)

        search_text = search_entry.get().strip().lower()

        if not search_text:
            results_list.insert(tk.END, "Please enter something to search for!")
            return

        if not os.path.exists(current_file):
            results_list.insert(tk.END, "No spreadsheet is currently open.")
            return

        workbook = load_workbook(current_file, data_only=True)
        sheet = workbook.active

        found = False

        for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            row_text = " | ".join(
                str(value) for value in row if value is not None
            )

            if search_text in row_text.lower():
                results_list.insert(
                    tk.END,
                    f"Row {row_number}: {row_text}"
                )
                found = True

        transaction_count_label.config(
            text=f"{results_list.size()} transactions found!"
        )

        workbook.close()

        if not found:
            results_list.insert(tk.END, "No matching transactions found.")

    clear_button = rounded_button(
        search_window,
        text="Clear",
        command=lambda: (
            search_entry.delete(0, tk.END),
            transaction_count_label.config(text="0 transactions found"),
            results_list.delete(0, tk.END)
        ),
        bg="black",
        fg="white",
        font=("Times New Roman", 13),
        width=70,
        height=25,
        corner_radius=12,
    )

    clear_button.place(relx=0.85, rely=0.28, anchor="ne")


    
    rounded_button(
        search_window,
        text="Search",
        command=search_transactions,
        bg="#168FE3",
        fg="white",
        font=("Times New Roman", 15),
        width=90,
        height=30,
        corner_radius=12
    ).pack(pady=5)

    search_entry.bind("<Return>", lambda event: search_transactions())

    rounded_button(
        search_window,
        text="View Spreadsheet",
        command=lambda: subprocess.run([
            "open",
            "-a",
            "Numbers",
            current_file
        ]),
        bg="#F5FEFD",
        fg="#168FE3",
        font=("Times New Roman", 11),
        width=130,
        height=30,
        corner_radius=12
    ).pack(pady=5)

    rounded_button(
        search_window,
        text="Edit Transaction",
        command=lambda: edit_transaction(results_list),
        bg="#F5FEFD",
        fg="#168FE3",
        font=("Times New Roman", 11),
        width=130,
        height=30,
        corner_radius=12
    ).pack(pady=5)

    rounded_button(
        search_window,
        text="Delete Transaction",
        command=lambda: delete_transaction(results_list),
        bg="#B71C1C",
        fg="white",
        font=("Times New Roman", 11),
        width=140,
        height=30,
        corner_radius=12
    ).pack(pady=5)

#---------------------------------------------
#   MAIN WINDOW
#---------------------------------------------

def refresh_spreadsheet_list():
    spreadsheet_list.delete(0, tk.END)

    search_text = search_entry.get().lower()

    folder = os.path.expanduser("~/Documents/Phillip Accounting")

    print("SPREADSHEET LIST REFRESH")
    print("Looking in:", folder)
    print("Folder exists:", os.path.exists(folder))

    if not os.path.exists(folder):
        return
    
    files = os.listdir(folder)

    print("Files found:", files)

    for filename in os.listdir(folder):
        if filename.endswith(".xlsx"):
            print("Adding spreadsheet:", filename)

            if search_text in filename.lower():
                spreadsheet_list.insert(tk.END, filename)

    print("LIST REFRESH COMPLETE")

def search_spreadsheets(event=None):
    refresh_spreadsheet_list()


def select_spreadsheet():
    global current_file

    selection = spreadsheet_list.curselection()

    if not selection:
        status_label.config(
            text=f"Current spreadsheet: None",
            font=("Times New Roman", 11)
        )

        return

    filename = spreadsheet_list.get(selection[0])

    current_file = os.path.join(
        DATA_FOLDER,
        filename
    )

    last_updated = datetime.fromtimestamp(
        os.path.getmtime(current_file)
    ).strftime("%B %d, %Y at %I:%M %p")

    workbook = load_workbook(current_file, data_only=True)
    sheet = workbook.active

    transaction_count = max(sheet.max_row - 1, 0)

    workbook.close()

    status_label.config(
        text=(
            f"Current spreadsheet: {filename} | "
            f"Last Updated: {last_updated} | "
            f"Transactions: {transaction_count}"
        ),
        font=("Times New Roman", 10),
        height=1
    )

def show_help():
    help_window = tk.Toplevel(window)
    help_window.title("Phillip Accounting - Help")
    help_window.geometry("580x600")

    title = tk.Label(
        help_window,
        text="PHILLIP ACCOUNTING - HELP",
        font=("Times New Roman", 22, "bold")
    )
    title.pack(pady=20)

    instructions = """
Getting Started:

1. Select a spreadsheet template.
2. Enter your accounting information naturally.
3. Click "Enter Data".
4. Phillip will interpret the transaction and add it
   to the current spreadsheet.

Transaction Search:

(ENSURE SPELLING IS CORRECT! Caps do not matter.)

[Search methods]:    
-Search by date: mm-dd-yyyy
-Search by description: Ex. "Walmart", "Target", etc.
-Search by category: Ex. "Groceries", "Taxes", "Refunds", etc.
-Search by direction: "Income" or "Expense" ONLY
-Search by payment method: "Cash", "Debit", or "Credit" ONLY
-Search by cost: Ex. 10 (Do NOT include '$')

    Then, press "Search" to view raw data. Click on the desired transaction,
    and press "View Spreadsheet" to see what spreadsheet it is located in.
    Press "Edit" to edit the individual cells. Press "Delete Transaction" to
    delete the whole row. Pressing "Clear" only clears the search bar, not
    the results box. After making an edit or clearing, press enter to refresh
    the results box.

Spreadsheet Templates:

(Select Template Before Creating New Spreadsheet or Inputing Data)

    Quarter Accounting:
        General accounting records.

    Sales Tracking:
        Tracks customers, products, quantities and sales.

    Checking:
        Tracks income, expenses, payment methods and running balance.

    1D Kinematics Calculator:
        Calculates all physics components of a 1D Kinematics equation.
            (Give at least two components of informations to calculate.)

    2D Kinematics Calculator:
        Calculates all physics components of a 2D Kinematics equation on
        different moons, planets, or other celestial entities. [Select an
        entity before inputing data!]
            (Give at least two components of informations to calculate.)
        

Managing Spreadsheets:

    Use "Search Existing Spreadsheets" to find spreadsheets. Press "Enter", on
    on your keyboard, after clearing the search box to return to the results
    menu.

    Select a spreadsheet and click "Select Spreadsheet"
    to make it the current spreadsheet.

    Use "Create New Spreadsheet" to create a new workbook.

    Use "Delete Spreadsheet" to remove a spreadsheet.



Transaction Examples:

    "I spent $45 at Walmart with my debit card."

    "I received $2,000 from my job."

    "I paid $30 cash for groceries."

Specfic Words for Different Templates:

    -ONLY use 'refund' for Checking Templates
    -ONLY use 'refunded' or 'refunding' for Sales Tracking or
     Quarter Accounting Templates

-----------------------------------------------------------------------

Phillip will use the information you provide to
categorize and record the transaction.

ALWAYS review automatically generated accounting
information for accuracy.

                        "Enjoy!"
                     -Jax R. Degen
                        

"""

    text = tk.Text(
        help_window,
        font=("Times New Roman", 13),
        wrap="word"
    )

    text.insert("1.0", instructions)
    text.config(state="disabled")
    text.pack(fill="both", expand=True, padx=20, pady=10)

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
        DATA_FOLDER,
        filename
    )

    try:
        os.remove(filepath)

        # If the deleted spreadsheet was the current spreadsheet,
        # switch back to the default spreadsheet name
        if os.path.abspath(current_file) == os.path.abspath(filepath):
            current_file = "Welcome.xlsx"

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
window.geometry("610x687")
window.configure(bg="#343534")
window.resizable(False,False)
window.pack_propagate(False)
window.grid_propagate(False)

search_button = rounded_button(
    window,
    text="Search Transactions",
    font=("Times New Roman", 11),
    bg="#656764",
    fg="white",
    command=open_search_window,
    height=25,
    width=140,
    corner_radius=12
)

search_button.place(relx=0.24, rely=0.01, anchor="ne")

help_button = tk.Canvas(
    window,
    width=32,
    height=32,
    bg="#343534",
    highlightthickness=0
)

help_button.place(relx=0.98, rely=0.01, anchor="ne")

help_button.create_oval(
    2,
    2,
    30,
    30,
    fill="#0066CC",
    outline="#87CEEB",
    width=2
)

help_button.create_text(
    16,
    16,
    text="?",
    font=("Times New Roman", 17),
    fill="white"
)

help_button.bind(
    "<Button-1>",
    lambda event: show_help()
)

title = tk.Label(
    window,
    text="PHILLIP ACCOUNTING",
    font=("Times New Roman", 23, "bold"),
    bg="#343534"
)

title.pack(pady=32)


selected_template = tk.StringVar(master=window)
selected_template.set("Quarter Accounting")

def template_changed(*args):
    if selected_template.get() == "2D Kinematics Calculator":
        open_planet_selector()

selected_template.trace_add("write", template_changed)

template_label = tk.Label(
    window,
    text="Select Spreadsheet Template:",
    font=("Times New Roman", 12),
    bg="#343534"
)

template_label.pack(pady=(8, 4))


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
    font=("Times New Roman", 15),
    bg="#343534"
)

instructions.pack()


entry = tk.Entry(
    window,
    width=43,
    font=("Courier New", 12),
    bg="#F8F6F0",
    fg="black",
    highlightbackground="#808080",
    highlightcolor="#808080",    
)

entry.pack(pady=5)
entry.focus_set()

clear_button = rounded_button(
    window,
    text="Clear",
    command=lambda: entry.delete(0, tk.END),
    bg="black",
    fg="white",
    font=("Times New Roman", 13),
    width=70,
    height=25,
    corner_radius=12,
)

clear_button.place(relx=0.99, rely=0.31, anchor="ne")


submit_button = rounded_button(
    window,
    text="Enter Data",
    font=("Times New Roman", 14, "bold"),
    bg="#FFD700",
    fg="black",
    command=add_transaction,
    height=30,
    width=120,
    corner_radius=12
)

submit_button.pack(pady=3)


# -------------------------
# SPREADSHEET SEARCH
# -------------------------

search_label = tk.Label(
    window,
    text="Search Existing Spreadsheets:",
    font=("Times New Roman", 11),
    bg="#343534"
)

search_label.pack(pady=(15, 5))

clear_button = rounded_button(
    window,
    text="Clear",
    command=lambda: search_entry.delete(0, tk.END),
    bg="black",
    fg="white",
    font=("Times New Roman", 13),
    width=70,
    height=25,
    corner_radius=12,
)

clear_button.place(relx=0.96, rely=0.48, anchor="ne")


search_entry = tk.Entry(
    window,
    width=40,
    font=("Courier New", 12),
    bg="#F8F6F0",
    fg="black",
    highlightbackground="#808080",
    highlightcolor="#808080",    
)

search_entry.pack(pady=3)

search_entry.bind(
    "<KeyRelease>",
    search_spreadsheets
)


spreadsheet_list = tk.Listbox(
    window,
    width=35,
    height=3,
    font=("Courier New", 12),
    bg="#429BDD",
    fg="black"
)

spreadsheet_list.pack(pady=2)


select_button = rounded_button(
    window,
    text="Select Spreadsheet",
    font=("Times New Roman", 11),
    bg="white",
    fg="black",
    command=select_spreadsheet,
    height=28,
    width=135,
    corner_radius=12
)

select_button.pack(pady=2)

delete_button = rounded_button(
    window,
    text="Delete Spreadsheet",
    font=("Times New Roman", 11),
    bg="#B71C1C",
    fg="white",
    command=delete_spreadsheet,
    height=28,
    width=135,
    corner_radius=12
)

delete_button.pack(pady=7)

# -------------------------
# NEW SPREADSHEET
# -------------------------

status_label = tk.Label(
    window,
    text=f"Current spreadsheet: {current_file}",
    font=("Times New Roman", 11),
    bg="#343534"
)

status_label.pack(pady=7)

clear_button = rounded_button(
    window,
    text="Clear",
    command=lambda: new_file_entry.delete(0, tk.END),
    bg="black",
    fg="white",
    font=("Times New Roman", 13),
    width=70,
    height=25,
    corner_radius=12,
)

clear_button.place(relx=0.96, rely=0.79, anchor="ne")


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


new_file_button = rounded_button(
    window,
    text="Create New Spreadsheet",
    font=("Courier New", 12),
    command=create_new_spreadsheet,
    width=230,
    height=30,
    corner_radius=13
)

new_file_button.pack(pady=5)

# -------------------------
# FOOTER
# -------------------------

footer = tk.Label(
    window,
    text="Powered by DeBeaul Tech and Solutions",
    font=("Times New Roman", 10),
    bg="#343534"
)

footer.pack(pady=5)

footer = tk.Label(
    window,
    text="Version 14.04",
    font=("Arial", 8),
    bg="#343534"
)

footer.pack(pady=0)

footer = tk.Label(
    window,
    text="Premium Edition",
    font=("Arial", 9),
    bg="#343534"
)

footer.pack(pady=0)

# Load spreadsheets when Phillip starts
refresh_spreadsheet_list()

window.mainloop()
