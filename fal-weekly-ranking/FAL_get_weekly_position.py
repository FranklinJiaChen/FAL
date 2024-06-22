import csv
import sys

TITLE = 2
POINTS = 3

def create_dictionary_from_csv(key_value_file):
    """
    The CSV file should be in MAL's weekly Anime Rankings format sorted by Weekly Pts.
    With the header rows being
    Total Ranking, ±LW, Anime Title, Weekly Pts, Total Pts
    """
    dictionary = {}

    with open(key_value_file, 'r') as file:
        reader = csv.reader(file)
        next(reader)
        for line_number, row in enumerate(reader, start=1):
            key = row[TITLE]
            value_str = row[POINTS].replace(",", "")
            try:
                value = (line_number, int(value_str))
            except ValueError:
                value = (line_number, 0)
            dictionary.setdefault(key, value)

    return dictionary

def write_weeekly_position_csv(dictionary, output_file):
    """
    The CSV file will be sorted by the Anime Title and in the format
    Anime Title, Weekly Position, Total Pts
    """
    sorted_dict = sorted(dictionary.items(), key=lambda x: x[0])
    with open(output_file, 'w', newline='') as file:
        writer = csv.writer(file)
        for key, value in sorted_dict:
            writer.writerow([key, value[0], value[1]])

if __name__ == '__main__':
    """
    Usage: python FAL-MAL-data-analyzer.py <week>
    """
    week = sys.argv[1]
    key_value_file = f"input/week-{week}.csv"
    output_file = f"output/week-{week}.csv"
    result_dictionary = create_dictionary_from_csv(key_value_file)
    write_weeekly_position_csv(result_dictionary, output_file)
    print(f"CSV file '{output_file}' created successfully.")
