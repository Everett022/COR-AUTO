import csv
import json

# Input and output file paths
csv_file = 'RPT386.csv'
json_file = 'RPT386.json'

# Read CSV and write to JSON
with open(csv_file, 'r') as csv_f:
    csv_reader = csv.DictReader(csv_f)
    data = [row for row in csv_reader]

with open(json_file, 'w') as json_f:
    json.dump(data, json_f, indent=4)

print(f"CSV converted to JSON and saved as {json_file}")