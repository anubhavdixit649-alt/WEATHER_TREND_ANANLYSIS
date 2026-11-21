# 🌦️ Weather Climate Search

A Python-based interactive weather data analysis project that allows users to search, filter, compare, and analyse weather records for different cities using an Excel dataset.

## About the Project

Weather Climate Search uses Python, Pandas, and NumPy to process weather records and provide a menu-driven search system. Users can explore city-wise climate summaries and find records based on dates, temperature, rainfall, wind speed, and weather conditions.

## Features

The program provides 13 menu options:

1. List all cities and the available data period
2. View the climate overview of a city
3. Display all records for a city
4. Search by city and date
5. Search by city and month
6. Search by date range
7. Filter by temperature range
8. Filter by rainfall range
9. Filter by wind speed range
10. Search by weather condition
11. Find the hottest, coldest, wettest, and windiest days
12. Compare cities
13. Rank all cities

City selection supports a full city name, a unique prefix, or a city number from the displayed list.

## Technologies Used

- **Python** – Main programming language
- **Pandas** – Reading, cleaning, filtering, and analysing tabular data
- **NumPy** – Numerical operations and condition classification
- **openpyxl** – Excel file support for Pandas
- **Excel (.xlsx)** – Input dataset

## Dataset

The project uses the workbook `Weather_Analysis_Results.xlsx`, specifically its `Weather Data` worksheet.

- **Records:** 100
- **Cities:** Bengaluru, Delhi, Jaipur, Mumbai, and Prayagraj
- **Date range:** 27 October 2026 to 15 November 2026

### Dataset columns

- `Date`
- `City`
- `Max_Temperature_C`
- `Min_Temperature_C`
- `Average_Temperature_C`
- `Rainfall_mm`
- `Wind_Speed_kmh`
- `Month`

## Weather Condition Rules

The program creates a `Condition` label using these rules, checked in the order shown:

| Condition | Rule |
|---|---|
| Heavy Rain | Rainfall ≥ 25 mm |
| Moderate Rain | Rainfall ≥ 5 mm |
| Light Rain | Rainfall > 0 mm |
| Hot | Maximum temperature ≥ 35°C |
| Windy | Wind speed ≥ 15 km/h |
| Warm | Average temperature ≥ 28°C |
| Pleasant | Average temperature ≥ 20°C |
| Cool | None of the above conditions apply |

**Note:** Because the rules are checked in order, a rainfall condition takes priority over temperature and wind conditions.

## Getting Started

### 1. Clone the repository

Replace `YOUR-USERNAME` with your GitHub username.

```bash
git clone https://github.com/YOUR-USERNAME/Weather-Climate-Search.git
cd Weather-Climate-Search
```

### 2. Install dependencies

```bash
pip install pandas numpy openpyxl
```

### 3. Add the dataset

Keep `Weather_Analysis_Results.xlsx` in the same folder as the Python file. If you run the program in Google Colab and the workbook is not found, the program can prompt you to upload it.

### 4. Run the program

```bash
python "weather_climate_search_single_cell_(1).py"
```

Follow the on-screen menu and enter the option number for the search or analysis you want to perform.

## Project Structure

```text
Weather-Climate-Search/
├── weather_climate_search_single_cell_(1).py
├── Weather_Analysis_Results.xlsx
└── README.md
```

## Limitations

- The program analyses the records in the supplied Excel workbook; it does not fetch live weather data.
- It does not predict future weather.
- Weather condition labels are assigned using the project's predefined threshold rules.

## Possible Future Improvements

- Connect to a live weather API.
- Add more cities and historical records.
- Add charts and data visualisations.
- Build a web-based interface.
- Explore machine-learning methods for weather forecasting.

## Purpose

This project was developed for educational purposes to practise Python programming, data cleaning, data filtering, and basic data analysis.

## Author

**Your Name**

---

If you find this project useful, feel free to star the repository.
