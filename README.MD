# SET and MAI Listed Companies Data Cleaning and Export

This Python script is designed to clean and export data from the "Profile of Listed Companies in SET and MAI" Excel file obtained from the Financial & Economic Data Bureau (https://www.fe-data.or.th). The script cleans the data and exports it to CSV and JSON formats.

## Prerequisites

Before using this script, make sure you have the following:

- Python 3.x installed on your system.
- The required Python packages (`pandas` and `re`) installed. You can install them using `pip`:

    ```bash
    pip install pandas
    ```

## Usage

1. Download the "Profile of Listed Companies in SET and MAI" Excel file (`PP_001.xlsx`) from [Financial & Economic Data Bureau](https://www.fe-data.or.th).

2. Place the downloaded Excel file in the same directory as this script.

3. Open the script and configure the following variables according to your data:

   - `file_path`: The filename of the Excel file you downloaded.
   - `sheet_name`: The name of the sheet within the Excel file where your data is located.
   - `header_text`: The text to identify the header row of your data.

4. Run the script using a Python interpreter.

    ```bash
    python script.py
    ```

5. The script will clean the data and export it to two formats:

   - `ListedCompanies.csv`: A CSV file with the cleaned data.
   - `ListedCompanies.json`: A JSON file with the cleaned data.

6. You can now use the exported data in your data analysis or other applications.

## Example

Here's an example of how to use the script:

```python
file_path = 'PP_001.xlsx'
sheet_name = 'profile'
header_text = 'Security Symbol'

cleaned_data = read_and_clean_excel(file_path, sheet_name, header_text)

if cleaned_data is not None:
    # Export to Excel
    export_to_csv(cleaned_data, 'ListedCompanies.csv')

    # Export to JSON
    export_to_json(cleaned_data, 'ListedCompanies.json')
```

## Using the Exported JSON Data

After running the script, you will have a JSON file named `ListedCompanies.json` containing the cleaned data. This JSON file can be easily imported into a NoSQL database like MongoDB. Here's an example of how you can do it:

1. Ensure you have MongoDB installed and running on your system.

2. Use a MongoDB client or command-line tool to import the data. For instance, you can use the `mongoimport` command to import the JSON data into a MongoDB collection:

   ```bash
   mongoimport --db your-database-name --collection your-collection-name --file ListedCompanies.json --jsonArray
    ```
    Replace your-database-name with the name of your MongoDB database and your-collection-name with the desired collection name.

3. Your data is now available in MongoDB, and you can perform queries and analysis as needed.

Feel free to explore and analyze the data further using MongoDB or any other NoSQL database of your choice.

![mondodb](https://github.com/candythink/Listed-Companies-Data-Cleaning/blob/withPic/pics/mongodb.png?raw=true)

## Contributing

Feel free to contribute to this project by creating pull requests or opening issues. Your feedback and contributions are highly appreciated.

## License

This project is licensed under the MIT License - see the LICENSE file for details.