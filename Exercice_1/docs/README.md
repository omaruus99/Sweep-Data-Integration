# Data Cleaning and CO2e Calculation for Emission Data

## Project Overview
This project involves cleaning and transforming activity data from three different sheets within an Excel file: Procurement Castel, CONCUR 2023 Cars Inc, and Energy data. Emission factors are applied to calculate CO2e emissions for each record. The result is then saved as separate CSV files.

## Steps Followed

### 1. Create Emission Factors Dictionary
A dictionary was created that maps Emission Factor ID to the corresponding emission factor values. This dictionary is used to calculate CO2e emissions for each entry.

### 2. Load Excel Data
The Excel file `activity_data_sweep-input.xlsx` is loaded, and the following sheets are extracted:
- Procurement Castel
- CONCUR 2023 Cars Inc
- Energy data
- EF (Emission Factors)

### 3. Procurement Castel Data Cleaning (1st sheet)
#### a. Replace Missing Aluminum Emission Factor ID
The correct Emission Factor ID for aluminum is identified from the EF sheet. Missing Emission Factor ID values in the Procurement Castel sheet where the material is aluminum are replaced with this ID.

#### b. Convert Units
Values in tonnes (`t`) are converted to kilograms (`kg`) by multiplying them by 1000.

#### c. Clean Quantity Column
Commas are replaced with dots in the Quantity column for consistency. The Quantity column is then converted to numeric format and rounded to two decimal places.

### 4. CONCUR 2023 Cars Inc Data Cleaning (2nd sheet)
#### a. Map Emission Factor ID Based on Transport Mode
The Emission Factor ID is replaced based on the transport mode (Air, Road, Train) using a predefined dictionary mapping.

#### b. Standardize Unit Column
All entries in the Unit column are set to `p.km` if they are not already `p.km`.

### 5. Energy Data Cleaning (3rd sheet)
#### a. Handle Missing Values
Missing values in the `%missing` column are replaced with 0 for records where the Status column indicates `Complete`.

#### b. Fix Units
The Unit column is corrected so that all entries are set to `kWh`.

#### c. Replace Outliers
An outlier value in the Quantity column (105560000000000000) is identified and replaced with the average consumption for similar records (filtered by Country, Location, and Type).

I chose to replace the extreme value. Here's the detailed explanation of this approach:

- **Country**: In this case, the country is 'France'. I assumed that offices located within the same country, particularly in France, would have relatively similar energy consumption patterns due to uniform energy consumption standards, strict environmental regulations, and a relatively consistent climate. This homogeneity supports the use of the average to correct an outlier in this context.

- **Location**: I limited the records to those from 'FR Offices', which refers to offices located in France. Offices within the same company and country are expected to have comparable energy consumption patterns, as they share similar infrastructure and standardized energy usage habits.

- **Type**: The outlier value was related to gas consumption, which is why I filtered by 'Gas'. Gas and electricity consumption follow different dynamics, so it was important to isolate gas-related records to ensure an accurate calculation.

### 6. Emission Factors Mapping and CO2e Calculation
A `mapping` function was created to calculate the CO2e emissions by multiplying the Quantity by the corresponding emission factor from the dictionary.

### 7. Data Validation
After the CO2e calculation, a validation step was performed to ensure data integrity. The following checks were applied to each dataset before saving the CSV files:

#### a. Duplicate Rows
Each dataset was checked for duplicate rows. If duplicates were found, they were automatically removed (6 duplicates detected in the CONCUR 2023 Cars Inc_Cleaned file).

#### b. Missing Values
The datasets were scanned for missing values, especially in critical columns like `Quantity`, `Emission Factor ID` and `CO2e`. If any missing values were detected.

#### c. Negative Values
The `Quantity` and `CO2e` columns were checked for negative values, as these are not valid in the context of emission calculations.

### 8. Save Cleaned Data to CSV
The cleaned and transformed datasets were saved as CSV files:
- Procurement Castel_Cleaned.csv
- CONCUR 2023 Cars Inc_Cleaned.csv
- Energy data_Cleaned.csv

### 9. Completion
The CSV files were generated and are ready for further analysis or reporting.
