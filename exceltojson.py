import pandas as pd
import json
import sys
from pathlib import Path

def excel_to_json(excel_file: str, json_file: str = None) -> None:
    """
    Convert Excel file to JSON format
    
    Args:
        excel_file (str): Path to Excel file
        json_file (str): Path to output JSON file (optional)
    """
    try:
        # Check if file exists
        if not Path(excel_file).exists():
            print(f"Error: File {excel_file} not found!")
            return

        # Read Excel file
        print(f"Reading {excel_file}...")
        df = pd.read_excel(excel_file)
        
        # Convert to JSON
        json_data = df.to_json(orient='records', indent=2)
        
        # If json_file is not specified, use the same name as excel file
        if json_file is None:
            json_file = Path(excel_file).stem + '.json'
        
        # Write to JSON file
        with open(json_file, 'w', encoding='utf-8') as f:
            f.write(json_data)
        
        print(f"Successfully converted to {json_file}")
        
    except Exception as e:
        print(f"Error during conversion: {str(e)}")

def main():
    # If file is provided as command line argument
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
        json_file = sys.argv[2] if len(sys.argv) > 2 else None
        excel_to_json(excel_file, json_file)
    else:
        print("Please provide Excel file path as argument")
        print("Usage: python exceltojson.py <excel_file> [json_file]")

if __name__ == "__main__":
    main()