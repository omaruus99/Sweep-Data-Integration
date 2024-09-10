import requests
import json
import matplotlib.pyplot as plt

API_KEY = input("Please enter your API_KEY : ")

# API URL for retrieving emission measurements
url = 'https://api.sweep.net/api/v1/measurements'

# Request parameters for the year 2022
params = {
    'start_date': '2022-01-01',
    'end_date': '2022-12-31',
}

# Headers for authentication, including the API Key
headers = {
    'X-Api-Key': API_KEY
}

# Perform the GET request to the API to retrieve the data
response = requests.get(url, headers=headers, params=params)

# Check if the request was successful (HTTP status code 200)
if response.status_code == 200:
    data = response.json()
else:
    print(f"Erreur {response.status_code}: {response.text}")

# Extract and aggregate emissions by Facility
emissions_by_facility = {}

# Iterate through each measurement in the received data
for measurement in data['measurements']:
    facility = measurement['customerData']['Facility']
    emissions = measurement['resultValue']
    
    if facility in emissions_by_facility:
        emissions_by_facility[facility] += emissions
    else:
        emissions_by_facility[facility] = emissions

# Sort the data in descending order of emissions
sorted_data = sorted(emissions_by_facility.items(), key=lambda x: x[1], reverse=True)
facilities_sorted = [item[0] for item in sorted_data]
emissions_sorted = [item[1] for item in sorted_data]

# Create a bar chart
max_index = 0
colors = ['skyblue'] * len(emissions_sorted)
colors[max_index] = 'red'  # The bar with the maximum emissions will be red (To show the most emissive Facility in 2022)

plt.figure(figsize=(12, 8))
bars = plt.bar(facilities_sorted, emissions_sorted, color=colors)
plt.title('Emission Distribution by Facility in 2022 (Sorted in Descending Order)')
plt.xlabel('Facility')
plt.ylabel('Emissions (tons of CO2e)')
plt.xticks(rotation=0)  

# Add value labels on top of each bar in the chart
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width() / 2, height, f'{height:.2f}', ha='center', va='bottom')

# Save and display the chart
plt.savefig('./output/emissions_by_facility.png', format='png')
plt.tight_layout()
plt.show()
