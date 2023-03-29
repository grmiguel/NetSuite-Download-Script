import requests
import os
import json
from bs4 import BeautifulSoup
import csv
import datetime  # Import the datetime module

# -------------------------------------------
# FILE AND PARAMETER HANDLING
# -------------------------------------------

def log_activity(log_file_path, message):
    with open(log_file_path, 'a') as log_file:
        log_file.write(f"{datetime.datetime.now()} - {message}\n")

def read_report_params(file_path):
    """
    Read report parameters from the JSON file.

    :param file_path: str, path to the JSON file containing report parameters.
    :return: list of dictionaries, report parameters for each report.
    """
    with open(file_path, 'r') as file:
        report_params = json.load(file)

    return report_params

# -------------------------------------------
# HTML PARSING
# -------------------------------------------

def parse_html_table(html):
    """
    Parse HTML table and return a list of lists.
    Each list within the main list represents a row in the table.

    :param html: str, the HTML code containing the table to be parsed.
    :return: list of lists, the parsed table data.
    """
    # Create a BeautifulSoup object to parse HTML
    soup = BeautifulSoup(html, 'html.parser')

    # Find the table element
    table = soup.find('table')

    # Validate if a table was found
    if table is None:
        raise ValueError("No table element found in the provided HTML")

    # Find all the rows in the table
    rows = table.find_all('tr')

    # Initialize an empty list to store table data
    data = []

    def is_number(value):
        try:
            float(value)
            return True
        except ValueError:
            return False

    # Iterate over each row and extract the data from its columns
    for row in rows:
        cols = row.find_all('td')
        # Extract the text from each column and remove leading/trailing white space
        cols = [col.text.strip() for col in cols]

        # Remove the equal symbol at the beginning of any cell value if it's followed by a number or a negative number
        cols = [col[1:] if col.startswith('=') and is_number(col[1:]) else col for col in cols]

        # Append the row data to the main list
        data.append(cols)

    return data

# -------------------------------------------
# CSV HANDLING
# -------------------------------------------

def save_table_data_to_csv(table_data, file_path):
    """
    Save table data to a CSV file.

    :param table_data: list of lists, table data to be saved.
    :param file_path: str, path to the CSV file to be saved.
    """
    with open(file_path, "w", encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(table_data)


# -------------------------------------------
# MAIN FUNCTION
# -------------------------------------------

def main():
    # Get the script's parent directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.abspath(os.path.join(script_dir, os.pardir))
    reports_file = os.path.join(script_dir, '..', 'netsuite_credentials.json')

    # Add a path for the log.txt file in the parent folder
    log_file_path = os.path.join(parent_dir, 'log.txt')

    # Read report parameters from the JSON file
    reports_params = read_report_params(reports_file)

    # Fetch and save each report

    headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Safari/537.36"
    }

    with requests.Session() as session:
        session.headers.update(headers)

        for report_params in reports_params:
            compid = report_params['compid']
            report_name = report_params.pop('name')
            url = f"https://{compid}.app.netsuite.com/app/reporting/webquery.nl"
            url_params = {k: v for k, v in report_params.items() if k != 'name'}

            try:
                # Send the request to the report URL with parameters
                response = session.get(url, params=url_params, headers=headers)

                # Check if the request was successful, otherwise raise an exception
                response.raise_for_status()

                # Parse the HTML table
                try:
                    table_data = parse_html_table(response.content.decode('utf-8'))
                except ValueError as e:
                    error_message = f"Error parsing {report_name} report: {e}"
                    print(error_message)
                    log_activity(log_file_path, error_message)
                    continue

                # Save the table data to a CSV file in the parent directory
                csv_file_path = os.path.join(parent_dir, f"{report_name}.csv")
                save_table_data_to_csv(table_data, csv_file_path)
                success_message = f"Downloaded {report_name} report."
                print(success_message)
                log_activity(log_file_path, success_message)
            except requests.exceptions.RequestException as e:
                error_message = f"Error fetching {report_name} report: {e}"
                print(error_message)
                log_activity(log_file_path, error_message)

if __name__ == "__main__":
    main()

