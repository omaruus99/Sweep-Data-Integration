import pandas as pd
import numpy as np
import openpyxl

# Create the emission factors dictionary
emission_factors = {
    259795: 1700,
    6553911: 0.137567911,
    6281066: 1.07412,
    6281843: 0.656101,
    6626527: 0.78669561,
    262740: 0.305,
    5418889: 0.149,
    5424149: 0.00592,
    262477: 0.406,
    3360917: 0.18387
}

# Load the Excel data
file_path = "./data/activity_data_sweep-input.xlsx"
procurement_data = pd.read_excel(file_path, sheet_name="Procurement Castel")
concur_data = pd.read_excel(file_path, sheet_name="CONCUR 2023 Cars Inc", header=1)
energy_data = pd.read_excel(file_path, sheet_name="Energy data", header=1)
ef_data = pd.read_excel(file_path, sheet_name="EF")


""" Cleaning Procurement Castel data """
# Find the emission factor ID for aluminum
aluminum_ef_id = ef_data[ef_data['Emission Factor Name'].str.contains('aluminum', case=False, na=False)]['Emission Factor ID'].values[0]

# Replace missing values in the 'Emission Factor ID' column for aluminum
procurement_data.loc[procurement_data['Material'].str.contains('Aluminium', case=False, na=False) & (procurement_data['Emission Factor ID'] == '-'), 'Emission Factor ID'] = aluminum_ef_id

# Convert values from tonnes (t) to kilograms (kg)
procurement_data.loc[procurement_data['Unit'] == 't', 'Quantity in kg'] = procurement_data.loc[procurement_data['Unit'] == 't', 'Quantity in kg'].astype(float) * 1000
procurement_data.loc[procurement_data['Unit'] == 't', 'Unit'] = 'kg'

# Replace commas with dots in the "Quantity" column
procurement_data['Quantity in kg'] = procurement_data['Quantity in kg'].astype(str).str.replace(',', '.')
procurement_data['Quantity in kg'] = pd.to_numeric(procurement_data['Quantity in kg'], errors='coerce')

# Round values in the "Quantity" column to 2 decimal places
procurement_data['Quantity in kg'] = procurement_data['Quantity in kg']
#.round(2)

""" Cleaning CONCUR 2023 Cars Inc data """
# Replace IDs in the 'Emission Factor ID' column based on the 'Transport' column
emission_factor_mapping = {
    'Air': 262740,
    'Road': 5418889,
    'Train': 5424149
}
concur_data['Emission Factor ID'] = concur_data['Transport'].map(emission_factor_mapping)

# Replace all values in the 'Unit' column with 'p.km' if they are not already 'p.km'
concur_data['Unit'] = concur_data['Unit'].apply(lambda x: 'p.km' if x != 'p.km' else x)

""" Cleaning Energy Data """
# Replace missing values in the '%missing' column with 0 if the status is "Complete"
energy_data.loc[energy_data['Status'] == 'Complete', '%missing'] = energy_data['%missing'].fillna(0)

# Fix units: replace with kWh if different
energy_data['Unit'] = energy_data['Unit'].apply(lambda x: 'kWh' if x != 'kWh' else x)

# Replace the outlier value with the average
outlier_value = 105560000000000000
filtered_data = energy_data[
    (energy_data['Country'] == 'France') &
    (energy_data['Location'] == 'FR Offices') &
    (energy_data['Type'] == 'Gas')
]
filtered_data_no_outlier = filtered_data[filtered_data['Quantity'] != outlier_value]
average_consumption = filtered_data_no_outlier['Quantity'].mean()
energy_data.loc[
    (energy_data['Country'] == 'France') &
    (energy_data['Location'] == 'FR Offices') &
    (energy_data['Type'] == 'Gas') &
    (energy_data['Quantity'] == outlier_value), 'Quantity'
] = average_consumption

""" Mapping """
def mapping(data, is_procurement=False):
    # For procurement_data, use 'Quantity in kg', otherwise use 'Quantity'
    quantity_column = 'Quantity in kg' if is_procurement else 'Quantity'
    
    # Ensure the correct 'Quantity' column and 'Emission Factor ID' are numeric
    data[quantity_column] = pd.to_numeric(data[quantity_column], errors='coerce')
    data['Emission Factor ID'] = pd.to_numeric(data['Emission Factor ID'], errors='coerce')

    # Create the CO2e column by mapping values
    data['CO2e'] = data.apply(lambda row: row[quantity_column] * emission_factors.get(row['Emission Factor ID'], 0) if pd.notnull(row['Emission Factor ID']) else np.nan, axis=1)
    
    return data

procurement_data = mapping(procurement_data, is_procurement=True)
concur_data = mapping(concur_data)
energy_data = mapping(energy_data)

""" Data validation """
# Function to check for duplicates, missing values, and negative values
def validate_data(df, filename):
    duplicates = df[df.duplicated(keep=False)]
    if not duplicates.empty:
        df = df.drop_duplicates()
    
    if df.isnull().any().any():
        print(f"Warning: Missing values found in {filename}")
    
    if (df.select_dtypes(include=[np.number]) < 0).any().any():
        print(f"Warning: Negative values found in {filename}")
    
    return df  

procurement_data = validate_data(procurement_data, 'Procurement Castel_Cleaned.csv')
concur_data = validate_data(concur_data, 'CONCUR 2023 Cars Inc_Cleaned.csv')
energy_data = validate_data(energy_data, 'Energy data_cleaned.csv')

# Save the cleaned and validated DataFrames to CSV files
procurement_data.to_csv('./output/Procurement Castel_Cleaned.csv', index=False)
concur_data.to_csv('./output/CONCUR 2023 Cars Inc_Cleaned.csv', index=False)
energy_data.to_csv('./output/Energy data_cleaned.csv', index=False)

print("CSV files have been created successfully. Please check the 'output' folder")

""" README.md generation """
def generate_readme():
    with open("./docs/README.md", "w") as f:
        f.write("# Data Cleaning and CO2e Calculation for Emission Data\n\n")
        f.write("## Project Overview\n")
        f.write("This project involves cleaning and transforming activity data from three different sheets within an Excel file: Procurement Castel, CONCUR 2023 Cars Inc, and Energy data. Emission factors are applied to calculate CO2e emissions for each record. The result is then saved as separate CSV files.\n\n")
        
        f.write("## Steps Followed\n\n")

        f.write("### 1. Create Emission Factors Dictionary\n")
        f.write("A dictionary was created that maps Emission Factor ID to the corresponding emission factor values. This dictionary is used to calculate CO2e emissions for each entry.\n\n")
        
        f.write("### 2. Load Excel Data\n")
        f.write("The Excel file `activity_data_sweep-input.xlsx` is loaded, and the following sheets are extracted:\n")
        f.write("- Procurement Castel\n")
        f.write("- CONCUR 2023 Cars Inc\n")
        f.write("- Energy data\n")
        f.write("- EF (Emission Factors)\n\n")

        f.write("### 3. Procurement Castel Data Cleaning (1st sheet)\n")
        f.write("#### a. Replace Missing Aluminum Emission Factor ID\n")
        f.write("The correct Emission Factor ID for aluminum is identified from the EF sheet. Missing Emission Factor ID values in the Procurement Castel sheet where the material is aluminum are replaced with this ID.\n\n")

        f.write("#### b. Convert Units\n")
        f.write("Values in tonnes (`t`) are converted to kilograms (`kg`) by multiplying them by 1000.\n\n")

        f.write("#### c. Clean Quantity Column\n")
        f.write("Commas are replaced with dots in the Quantity column for consistency. The Quantity column is then converted to numeric format and rounded to two decimal places.\n\n")
        
        f.write("### 4. CONCUR 2023 Cars Inc Data Cleaning (2nd sheet)\n")
        f.write("#### a. Map Emission Factor ID Based on Transport Mode\n")
        f.write("The Emission Factor ID is replaced based on the transport mode (Air, Road, Train) using a predefined dictionary mapping.\n\n")

        f.write("#### b. Standardize Unit Column\n")
        f.write("All entries in the Unit column are set to `p.km` if they are not already `p.km`.\n\n")

        f.write("### 5. Energy Data Cleaning (3rd sheet)\n")
        f.write("#### a. Handle Missing Values\n")
        f.write("Missing values in the `%missing` column are replaced with 0 for records where the Status column indicates `Complete`.\n\n")

        f.write("#### b. Fix Units\n")
        f.write("The Unit column is corrected so that all entries are set to `kWh`.\n\n")

        f.write("#### c. Replace Outliers\n")
        f.write("An outlier value in the Quantity column (105560000000000000) is identified and replaced with the average consumption for similar records (filtered by Country, Location, and Type).\n\n")

        f.write("I chose to replace the extreme value. Here's the detailed explanation of this approach:\n\n")
        f.write("- **Country**: In this case, the country is 'France'. I assumed that offices located within the same country, particularly in France, would have relatively similar energy consumption patterns due to uniform energy consumption standards, strict environmental regulations, and a relatively consistent climate. This homogeneity supports the use of the average to correct an outlier in this context.\n\n")
        f.write("- **Location**: I limited the records to those from 'FR Offices', which refers to offices located in France. Offices within the same company and country are expected to have comparable energy consumption patterns, as they share similar infrastructure and standardized energy usage habits.\n\n")
        f.write("- **Type**: The outlier value was related to gas consumption, which is why I filtered by 'Gas'. Gas and electricity consumption follow different dynamics, so it was important to isolate gas-related records to ensure an accurate calculation.\n\n")
        
        f.write("### 6. Emission Factors Mapping and CO2e Calculation\n")
        f.write("A `mapping` function was created to calculate the CO2e emissions by multiplying the Quantity by the corresponding emission factor from the dictionary.\n\n")

        f.write("### 7. Data Validation\n")
        f.write("After the CO2e calculation, a validation step was performed to ensure data integrity. The following checks were applied to each dataset before saving the CSV files:\n\n")

        f.write("#### a. Duplicate Rows\n")
        f.write("Each dataset was checked for duplicate rows. If duplicates were found, they were automatically removed (6 duplicates detected in the CONCUR 2023 Cars Inc_Cleaned file).\n\n")

        f.write("#### b. Missing Values\n")
        f.write("The datasets were scanned for missing values, especially in critical columns like `Quantity`, `Emission Factor ID` and `CO2e`. If any missing values were detected.\n\n")

        f.write("#### c. Negative Values\n")
        f.write("The `Quantity` and `CO2e` columns were checked for negative values, as these are not valid in the context of emission calculations.\n\n")

        f.write("### 8. Save Cleaned Data to CSV\n")
        f.write("The cleaned and transformed datasets were saved as CSV files:\n")
        f.write("- Procurement Castel_Cleaned.csv\n")
        f.write("- CONCUR 2023 Cars Inc_Cleaned.csv\n")
        f.write("- Energy data_Cleaned.csv\n\n")

        f.write("### 9. Completion\n")
        f.write("The CSV files were generated and are ready for further analysis or reporting.\n")

generate_readme()

print("README.md has been generated successfully. Please check the 'docs' folder")