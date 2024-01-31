import pandas as pd

json_file_path = 'https://raw.githubusercontent.com/candythink/Thai-Listed-Companies/main/ThaiListedCompanies.json'

# Read the JSON Lines file into a pandas DataFrame
df = pd.read_json(json_file_path, lines=True)