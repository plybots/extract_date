import json
import pandas as pd
import os
from dateutil import parser
import pytz

def convert_date_format(date_string: str):
    try:
        date_object = parser.parse(date_string)
        converted_date = date_object.astimezone(pytz.timezone('Africa/Kampala')).strftime("%Y-%m-%dT%H:%M:%S.%f")
        converted_date = converted_date[:23] + converted_date[26:] + '+03:00' # Modify here to get the format with 3 decimal places
        return converted_date
    except ValueError:
        return date_string  # If date is not recognizable return as is.

def main():
    project_dir = os.getcwd() # Get the project directory
    file_path = os.path.join(project_dir, 'data.xlsx') # Get the complete file path
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"An error occurred: {e}")
        return

    data_dict = df.to_dict('records')
    dates = []

    # Initialize a counter
    counter_matched_rows = 0
    for record in data_dict:
        for key, value in record.items():
            if key == 'dataValues':
                data_values = json.loads(value)
                if data_values:
                    # Increment the counter when data values are found
                    counter_matched_rows += 1
                    date = data_values[0].get('value')
                    if date:
                        dates.append(convert_date_format(date))
                    else:
                        dates.append("")
                else:
                    dates.append("")

    df['date'] = dates

    output_file_path = os.path.join(project_dir, 'new_data.xlsx')
    df.to_excel(output_file_path, index=False)

if __name__ == "__main__":
    main()