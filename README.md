# Buzzsprout-to-Excel
Python script to fetch and analyze podcast episode details and statistics from Buzzsprout's API and export them to an Excel file.

---

### README.md

# Buzzsprout API Integration with Python

This Python script allows you to fetch podcast episode details and stats from Buzzsprout's API and save them into an Excel file.

## Prerequisites

- Python 3.x
- `requests` package
- `pandas` package
- `openpyxl` package

## Installation

Install the required packages if you haven't already:

```bash
pip install requests pandas openpyxl
```

## Configuration

1. **API Token**: Replace `YOUR_API_HERE` in the `API_TOKEN` variable with your Buzzsprout API token.
   
2. **Excel File Path**: Update the `EXCEL_FILE_PATH` variable with the full path to where you want to save the Excel file.

## Usage

1. Clone the repository or download the script to your local machine.

2. Open a terminal and navigate to the directory where the script is located.

3. Run the script:

   ```bash
   python your_script_name.py
   ```

4. If everything is set up correctly, the script will fetch podcast episode details and statistics and save them into an Excel file.

## Error Handling

The script has built-in error handling for:

- API request failures
- Invalid JSON responses

## Contributing

Feel free to fork the repository and submit pull requests.

## License

MIT License

---
