# Sweep Data Integration

## Project Overview
This project processes activity data from multiple sources to clean, transform, and calculate CO2e emissions using predefined emission factors, with results stored in CSV files for further analysis or reporting. Additionally, it includes extracting emission data from the Sweep API for the year 2022, aggregating the data by Facility, and generating a bar chart to visualize the distribution of emissions.

## Table of Contents
- [Project Overview](#project-overview)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [How to Use](#how-to-use)
- [Output Files](#output-files)


## Project Structure
The project directory is structured as follows:

```bash
Sweep_Data_Integration
├── project_1
│   └── data
│   └── docs
│   └── output
│   └── script.py
├── project_2
│   ├── output
│   └── script.py
├── README.md
└── requirements.txt
```

## Requirements

The project requires the following Python libraries:

- pandas
- numpy
- openpyxl
- requests
- matplotlib

## How to Use

### Clone the Repository:

```bash
git clone https://github.com/omaruus99/Sweep-Data-Integration
```
### Install Required Libraries:
```bash
cd Sweep-Data-Integration
pip install -r requirements.txt
```

### Run the script for project 1
```bash
cd project_1 
python script.py
```

### Run the script for project 2
```bash
cd project_2 
python script.py
```
For security reasons, the API key is not hard-coded. You will need to enter your API key when running the script.


## Output Files
### For project 1
After running the script, the following cleaned CSV files will be available in `project_1/output` :

- Procurement_Castel_Cleaned.csv
- CONCUR_2023_Cars_Inc_Cleaned.csv
- Energy_data_Cleaned.csv

The generated README.md file will be saved in the `project_1/docs`.

### For project 2
The graph showing the distribution of emissions by Facility for the year 2022 is saved in the `project_2/output`.