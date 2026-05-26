import csv
import json
import sys

def csv_to_json(csv_path, json_path):
    """Convert a CSV file to JSON."""
    try:
        with open(csv_path, newline="", encoding="utf-8") as csvfile:
            reader = csv.DictReader(csvfile)
            data = list(reader)

        with open(json_path, "w", encoding="utf-8") as jsonfile:
            json.dump(data, jsonfile, indent=4)

        print(f"Successfully converted '{csv_path}' → '{json_path}'")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python csv_to_json.py <input.csv> <output.json>")
    else:
        csv_to_json(sys.argv[1], sys.argv[2])