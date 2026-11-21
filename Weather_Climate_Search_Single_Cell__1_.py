# -*- coding: utf-8 -*-
"""
Converted from IPYNB to PY
"""

# %% [code] Cell 1
# Weather Climate Search - interactive program (Python + Pandas + NumPy)
#
# Reads the 'Weather Data' sheet of Weather_Analysis_Results.xlsx and lets you
# search cities, check temperature / rainfall / wind, see the climate of a city
# and compare cities.
#
# How to run: just run this single cell. If the Excel file is not found, a box
# opens (in Google Colab) so you can upload Weather_Analysis_Results.xlsx.
#
# Pandas needs the openpyxl package to read .xlsx files (pip install openpyxl).
#
# Everything shown is calculated from the Excel file. The only extra column is
# "Condition", a simple label this program creates from rainfall, temperature
# and wind (the rules are in the SETTINGS section below, you can change them).

import os
import numpy as np
import pandas as pd

# ----------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------
FILE_NAME = "Weather_Analysis_Results.xlsx"
SHEET_NAME = "Weather Data"

# Rules for the "Condition" label (change them if you like)
HEAVY_RAIN_MM = 25      # rainfall at or above this  -> Heavy Rain
MODERATE_RAIN_MM = 5    # rainfall at or above this  -> Moderate Rain
HOT_DAY_MAX_C = 35      # max temperature at or above this -> Hot
WINDY_KMH = 15          # wind speed at or above this -> Windy
WARM_AVG_C = 28         # average temperature at or above this -> Warm
COOL_AVG_C = 20         # average temperature below this -> Cool

NUMBER_COLUMNS = ["Max_Temperature_C", "Min_Temperature_C",
                  "Average_Temperature_C", "Rainfall_mm", "Wind_Speed_kmh"]

TABLE_COLUMNS = ["Date", "City", "Max_Temperature_C", "Min_Temperature_C",
                 "Average_Temperature_C", "Rainfall_mm", "Wind_Speed_kmh",
                 "Condition"]

HEADERS = {
    "Date": "Date", "City": "City",
    "Max_Temperature_C": "Max Temp (C)",
    "Min_Temperature_C": "Min Temp (C)",
    "Average_Temperature_C": "Avg Temp (C)",
    "Rainfall_mm": "Rain (mm)",
    "Wind_Speed_kmh": "Wind (km/h)",
    "Condition": "Condition",
}

CITY_TABLE_HEADERS = {
    "avg_temp": "Avg Temp (C)", "highest": "Highest Temp (C)",
    "lowest": "Lowest Temp (C)", "total_rain": "Total Rain (mm)",
    "rainy_days": "Rainy Days", "avg_wind": "Avg Wind (km/h)",
    "max_wind": "Max Wind (km/h)",
}


# ----------------------------------------------------------------------
# LOADING AND CLEANING THE EXCEL DATA
# ----------------------------------------------------------------------
def is_excel(name):
    low = name.lower()
    return low.endswith((".xlsx", ".xls")) and not name.startswith("~$")


def find_file():
    """Find the Excel file: known folders first, then a Colab upload, then ask for the path."""
    folders = []
    try:
        folders.append(os.path.dirname(os.path.abspath(__file__)))
    except NameError:                      # happens inside notebooks
        pass
    folders += [os.getcwd(), "/content", "/content/drive/MyDrive",
                os.path.expanduser("~/Downloads"), os.path.expanduser("~")]
    folders = [f for i, f in enumerate(folders) if os.path.isdir(f) and f not in folders[:i]]

    # 1. exact file name
    for folder in folders:
        path = os.path.join(folder, FILE_NAME)
        if os.path.exists(path):
            return path

    # 2. similar name, e.g. "Weather_Analysis_Results (1).xlsx" (Colab adds (1) to duplicates)
    for folder in folders:
        for name in sorted(os.listdir(folder)):
            if is_excel(name) and "weather" in name.lower():
                print(f"Using '{name}' found in {folder}")
                return os.path.join(folder, name)

    # 3. inside Google Colab: open the upload dialog
    try:
        from google.colab import files
        print(f"{FILE_NAME} was not found. Choose the file in the upload box below...")
        uploaded = files.upload()
        for name in uploaded:
            if is_excel(name):
                return os.path.abspath(name)
        print("No Excel file was uploaded.")
    except ImportError:
        pass

    # 4. ask for the path
    print(f"Could not find {FILE_NAME}. Folders checked: " + ", ".join(folders))
    while True:
        text = input("Type the full path of the Excel file or its folder (Enter = quit): ").strip()
        text = text.strip('"').strip("'")
        if text == "":
            raise SystemExit("No file given. Program stopped.")
        if os.path.isdir(text):
            text = os.path.join(text, FILE_NAME)
        if os.path.exists(text):
            return text
        print("That file does not exist. Try again.")


def load_data(path):
    """Read the weather sheet, clean it and add the Condition column."""
    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    start_rows = len(df)

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["City"] = df["City"].astype(str).str.strip().str.title()
    for column in NUMBER_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df = df.dropna(subset=["Date"] + NUMBER_COLUMNS)
    df = df.drop_duplicates(subset=["Date", "City"])
    df = df.sort_values(["City", "Date"]).reset_index(drop=True)

    df["Month"] = df["Date"].dt.month_name()
    df["Month_No"] = df["Date"].dt.month

    # Simple daily label made from the numbers
    rain = df["Rainfall_mm"].values
    conditions = [
        rain >= HEAVY_RAIN_MM,
        rain >= MODERATE_RAIN_MM,
        rain > 0,
        df["Max_Temperature_C"].values >= HOT_DAY_MAX_C,
        df["Wind_Speed_kmh"].values >= WINDY_KMH,
        df["Average_Temperature_C"].values >= WARM_AVG_C,
        df["Average_Temperature_C"].values >= COOL_AVG_C,
    ]
    labels = ["Heavy Rain", "Moderate Rain", "Light Rain",
              "Hot", "Windy", "Warm", "Pleasant"]
    df["Condition"] = np.select(conditions, labels, default="Cool")

    removed = start_rows - len(df)
    print(f"Loaded {len(df)} records from '{os.path.basename(path)}'"
          f" ({removed} bad or duplicate rows removed).")
    return df


# ----------------------------------------------------------------------
# INPUT HELPERS
# ----------------------------------------------------------------------
def resolve_city(text):
    """Turn what the user typed (name, part of a name, or list number) into a city."""
    text = text.strip().lower()
    if text.isdigit() and 1 <= int(text) <= len(cities):
        return cities[int(text) - 1]
    for city in cities:
        if city.lower() == text:
            return city
    if text != "":
        starts = [c for c in cities if c.lower().startswith(text)]
        if len(starts) == 1:
            return starts[0]
    return None


def print_city_list():
    print("Cities: " + ", ".join(f"{i}. {c}" for i, c in enumerate(cities, start=1)))


def ask_city(allow_all=False):
    """Ask for one city. Returns the city name, 'ALL' (if allowed) or None (cancel)."""
    print_city_list()
    if allow_all:
        prompt = "Enter city name or number (Enter = all cities): "
    else:
        prompt = "Enter city name or number (Enter = cancel): "
    while True:
        text = input(prompt).strip()
        if text == "":
            return "ALL" if allow_all else None
        city = resolve_city(text)
        if city is not None:
            return city
        print("City not found. Please try again.")


def ask_many_cities():
    """Ask for two or more cities (comma separated) or ALL."""
    print_city_list()
    text = input("Enter two or more cities separated by commas (or ALL): ").strip()
    if text.lower() == "all":
        return list(cities)
    chosen = []
    for part in text.split(","):
        if part.strip() == "":
            continue
        city = resolve_city(part)
        if city is None:
            print(f"'{part.strip()}' not found, skipping it.")
        elif city not in chosen:
            chosen.append(city)
    return chosen


def ask_number(prompt):
    while True:
        text = input(prompt).strip()
        try:
            return float(text)
        except ValueError:
            print("Please enter a valid number.")


def ask_date(prompt="Enter date (YYYY-MM-DD, e.g. 2026-11-05): "):
    while True:
        text = input(prompt).strip()
        for date_format in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
            try:
                return pd.to_datetime(text, format=date_format)
            except ValueError:
                pass
        print("Invalid date. Use YYYY-MM-DD (or DD-MM-YYYY).")


def ask_month():
    """Ask for a month (number or name). Only months that exist in the data are accepted."""
    available = data[["Month_No", "Month"]].drop_duplicates().sort_values("Month_No")
    names = list(available["Month"])
    print("Months in the data: " + ", ".join(names))
    while True:
        text = input("Enter month (name or number, Enter = cancel): ").strip().lower()
        if text == "":
            return None
        for number, name in zip(available["Month_No"], available["Month"]):
            if text == str(number) or (len(text) >= 3 and name.lower().startswith(text)):
                return number
        print("No data for that month. Choose from:", ", ".join(names))


def rows_for(city):
    """All rows for a city, or every row when city is 'ALL'."""
    if city == "ALL":
        return data
    return data[data["City"] == city]


# ----------------------------------------------------------------------
# DISPLAY HELPERS
# ----------------------------------------------------------------------
def show_table(rows, limit=25):
    """Print records as a simple text table."""
    if rows.empty:
        print("No records found.")
        return
    total = len(rows)
    print(f"\nRecords found: {total}")
    count = total
    if total > limit:
        answer = input(f"How many rows to show? (Enter = {limit}, 'all' = everything): ").strip().lower()
        if answer == "all":
            count = total
        elif answer.isdigit() and int(answer) > 0:
            count = int(answer)
        else:
            count = limit
    out = rows[TABLE_COLUMNS].head(count).copy()
    out["Date"] = out["Date"].dt.strftime("%Y-%m-%d")
    out = out.rename(columns=HEADERS)
    print(out.to_string(index=False))
    if total > count:
        print(f"... showing {count} of {total} records.")


def city_table(rows):
    """One summary row per city (used by compare and rank)."""
    table = rows.groupby("City").agg(
        avg_temp=("Average_Temperature_C", "mean"),
        highest=("Max_Temperature_C", "max"),
        lowest=("Min_Temperature_C", "min"),
        total_rain=("Rainfall_mm", "sum"),
        rainy_days=("Rainfall_mm", lambda s: int((s > 0).sum())),
        avg_wind=("Wind_Speed_kmh", "mean"),
        max_wind=("Wind_Speed_kmh", "max"),
    )
    return table.round(2)


def climate_description(c):
    """A short plain-English description made from the numbers."""
    avg_temp = c["Average_Temperature_C"].mean()
    rainy_share = (c["Rainfall_mm"] > 0).mean() * 100
    avg_wind = c["Wind_Speed_kmh"].mean()

    temp_word = np.select([avg_temp >= 30, avg_temp >= 26, avg_temp >= 20],
                          ["Hot", "Warm", "Mild"], default="Cool")
    rain_word = np.select([rainy_share == 0, rainy_share <= 20, rainy_share <= 40],
                          ["dry", "mostly dry", "occasionally rainy"], default="often rainy")
    wind_word = np.select([avg_wind >= 12, avg_wind >= 8],
                          ["breezy", "moderate winds"], default="light winds")
    return f"{temp_word} and {rain_word}, with {wind_word}"


# ----------------------------------------------------------------------
# MENU ACTIONS
# ----------------------------------------------------------------------
def action_list_cities():
    print("\nAvailable cities:")
    for i, city in enumerate(cities, start=1):
        days = int((data["City"] == city).sum())
        print(f" {i}. {city}  ({days} days of data)")
    print(f"\nData covers {data['Date'].min():%Y-%m-%d} to {data['Date'].max():%Y-%m-%d}.")


def action_climate_overview():
    city = ask_city()
    if city is None:
        return
    c = data[data["City"] == city]

    hottest = c.loc[c["Max_Temperature_C"].idxmax()]
    coldest = c.loc[c["Min_Temperature_C"].idxmin()]
    wettest = c.loc[c["Rainfall_mm"].idxmax()]
    windiest = c.loc[c["Wind_Speed_kmh"].idxmax()]
    rainy_days = int((c["Rainfall_mm"] > 0).sum())

    print(f"\n=== CLIMATE OVERVIEW: {city.upper()} ===")
    print(f"Period : {c['Date'].min():%Y-%m-%d} to {c['Date'].max():%Y-%m-%d} ({len(c)} days)")
    print(f"Climate: {climate_description(c)}")
    print("(description is generated by this program from the numbers)\n")

    lines = [
        ("Average temperature (C)", round(c["Average_Temperature_C"].mean(), 2)),
        ("Average of daily max (C)", round(c["Max_Temperature_C"].mean(), 2)),
        ("Average of daily min (C)", round(c["Min_Temperature_C"].mean(), 2)),
        ("Highest temperature (C)", f"{hottest['Max_Temperature_C']} on {hottest['Date']:%Y-%m-%d}"),
        ("Lowest temperature (C)", f"{coldest['Min_Temperature_C']} on {coldest['Date']:%Y-%m-%d}"),
        ("Total rainfall (mm)", round(c["Rainfall_mm"].sum(), 1)),
        ("Rainy days", f"{rainy_days} of {len(c)}"),
        ("Wettest day", f"{wettest['Rainfall_mm']} mm on {wettest['Date']:%Y-%m-%d}"),
        ("Average wind speed (km/h)", round(c["Wind_Speed_kmh"].mean(), 2)),
        ("Windiest day", f"{windiest['Wind_Speed_kmh']} km/h on {windiest['Date']:%Y-%m-%d}"),
    ]
    for label, value in lines:
        print(f"{label:<28}: {value}")

    print("\n--- Days by condition ---")
    counts = c["Condition"].value_counts()
    print(counts.rename_axis("Condition").reset_index(name="Days").to_string(index=False))

    print("\n--- Month by month ---")
    monthly = c.groupby("Month_No").agg(
        Month=("Month", "first"),
        Days=("Date", "count"),
        Avg_Temp=("Average_Temperature_C", "mean"),
        Highest=("Max_Temperature_C", "max"),
        Lowest=("Min_Temperature_C", "min"),
        Total_Rain=("Rainfall_mm", "sum"),
        Avg_Wind=("Wind_Speed_kmh", "mean"),
    ).round(2)
    print(monthly.to_string(index=False))


def action_all_records():
    city = ask_city()
    if city is not None:
        show_table(rows_for(city))


def action_city_and_date():
    city = ask_city()
    if city is None:
        return
    date = ask_date()
    result = data[(data["City"] == city) & (data["Date"] == date)]
    if result.empty:
        print(f"No record for {city} on {date:%Y-%m-%d}. "
              f"Data covers {data['Date'].min():%Y-%m-%d} to {data['Date'].max():%Y-%m-%d}.")
    else:
        show_table(result)


def action_city_and_month():
    city = ask_city()
    if city is None:
        return
    month = ask_month()
    if month is None:
        return
    show_table(data[(data["City"] == city) & (data["Month_No"] == month)])


def action_date_range():
    city = ask_city(allow_all=True)
    start = ask_date("Enter start date (YYYY-MM-DD): ")
    end = ask_date("Enter end date   (YYYY-MM-DD): ")
    if start > end:
        start, end = end, start
    rows = rows_for(city)
    show_table(rows[rows["Date"].between(start, end)])


def filter_by_range(column, label):
    city = ask_city(allow_all=True)
    low = ask_number(f"Enter minimum {label}: ")
    high = ask_number(f"Enter maximum {label}: ")
    if low > high:
        low, high = high, low
    rows = rows_for(city)
    show_table(rows[rows[column].between(low, high)])


def action_temperature_filter():
    print("Which temperature?  1 = Average   2 = Daily maximum   3 = Daily minimum")
    choice = input("Choice (Enter = 1): ").strip()
    column = {"2": "Max_Temperature_C", "3": "Min_Temperature_C"}.get(choice, "Average_Temperature_C")
    filter_by_range(column, "temperature (C)")


def action_rainfall_filter():
    filter_by_range("Rainfall_mm", "rainfall (mm)")


def action_wind_filter():
    filter_by_range("Wind_Speed_kmh", "wind speed (km/h)")


def action_condition_search():
    available = sorted(data["Condition"].unique())
    print("Conditions: " + ", ".join(f"{i}. {c}" for i, c in enumerate(available, start=1)))
    text = input("Enter condition name or number (Enter = cancel): ").strip().lower()
    if text == "":
        return
    chosen = None
    if text.isdigit() and 1 <= int(text) <= len(available):
        chosen = available[int(text) - 1]
    else:
        for name in available:
            if name.lower() == text:
                chosen = name
    if chosen is None:
        print("Condition not found.")
        return
    city = ask_city(allow_all=True)
    rows = rows_for(city)
    show_table(rows[rows["Condition"] == chosen])


def action_extreme_days():
    options = {
        "1": ("Hottest days", "Max_Temperature_C", False),
        "2": ("Coldest days", "Min_Temperature_C", True),
        "3": ("Wettest days", "Rainfall_mm", False),
        "4": ("Windiest days", "Wind_Speed_kmh", False),
    }
    print("1 = Hottest days   2 = Coldest days   3 = Wettest days   4 = Windiest days")
    choice = input("Choice (Enter = cancel): ").strip()
    if choice not in options:
        return
    title, column, ascending = options[choice]
    city = ask_city(allow_all=True)
    text = input("How many days to show? (Enter = 5): ").strip()
    count = int(text) if text.isdigit() and int(text) > 0 else 5

    rows = rows_for(city).sort_values(column, ascending=ascending).head(count)
    scope = "all cities" if city == "ALL" else city
    print(f"\n{title} - {scope}")
    show_table(rows, limit=max(count, 25))


def action_compare():
    chosen = ask_many_cities()
    if len(chosen) < 2:
        print("Please choose at least two valid cities.")
        return
    table = city_table(data[data["City"].isin(chosen)])
    table = table.rename(columns=CITY_TABLE_HEADERS)
    print("\n=== CITY COMPARISON ===")
    print(table.to_string())


def action_rank():
    options = {
        "1": ("Average temperature (hottest first)", "avg_temp", False),
        "2": ("Highest temperature (hottest first)", "highest", False),
        "3": ("Lowest temperature (coldest first)", "lowest", True),
        "4": ("Total rainfall (wettest first)", "total_rain", False),
        "5": ("Average wind speed (windiest first)", "avg_wind", False),
    }
    for key, (title, _, _) in options.items():
        print(f" {key} = {title}")
    choice = input("Choice (Enter = cancel): ").strip()
    if choice not in options:
        return
    title, column, ascending = options[choice]
    table = city_table(data).sort_values(column, ascending=ascending)
    table.insert(0, "Rank", np.arange(1, len(table) + 1))
    table = table.rename(columns=CITY_TABLE_HEADERS)
    print(f"\n=== CITIES RANKED BY: {title} ===")
    print(table.to_string())


# ----------------------------------------------------------------------
# MAIN MENU
# ----------------------------------------------------------------------
MENU = [
    ("1", "List all cities and data period", action_list_cities),
    ("2", "Climate overview of a city (temperature, rain, wind)", action_climate_overview),
    ("3", "Show all records of a city", action_all_records),
    ("4", "Search by city and date", action_city_and_date),
    ("5", "Search by city and month", action_city_and_month),
    ("6", "Search by date range", action_date_range),
    ("7", "Filter by temperature range", action_temperature_filter),
    ("8", "Filter by rainfall range", action_rainfall_filter),
    ("9", "Filter by wind speed range", action_wind_filter),
    ("10", "Search by weather condition (Rain, Hot, Windy ...)", action_condition_search),
    ("11", "Hottest / coldest / wettest / windiest days", action_extreme_days),
    ("12", "Compare cities", action_compare),
    ("13", "Rank all cities", action_rank),
]


def show_menu():
    print("\n" + "=" * 62)
    print("                    WEATHER SEARCH")
    print("=" * 62)
    for key, text, _ in MENU:
        print(f" {key:>2}. {text}")
    print("  0. Exit")
    print("=" * 62)


data = load_data(find_file())
cities = sorted(data["City"].unique())

try:
    while True:
        show_menu()
        choice = input("Enter your choice (0-13): ").strip()
        if choice == "0":
            print("Goodbye!")
            break
        for key, _, action in MENU:
            if choice == key:
                action()
                break
        else:
            print("Invalid choice. Please enter a number from 0 to 13.")
            continue
        input("\nPress Enter to return to the menu...")
except (KeyboardInterrupt, EOFError):
    print("\nProgram stopped.")
