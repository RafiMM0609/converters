import json
import pandas as pd
import re
import os
import sys

def clean_latitude(lat):
    """Clean and validate latitude values. Valid range: -90 to 90"""
    if lat is None or lat == '' or str(lat).strip() == '':
        return '0.0'
    
    try:
        # Convert to string and clean
        lat_str = str(lat).strip()
        
        # Remove any non-numeric characters except minus sign and decimal point
        lat_clean = re.sub(r'[^\d\-\.]', '', lat_str)
        
        if not lat_clean or lat_clean == '-' or lat_clean == '.':
            return '0.0'
        
        # Convert to float for validation
        lat_float = float(lat_clean)
        
        # Validate latitude range (-90 to 90)
        if lat_float < -90:
            lat_float = -90.0
        elif lat_float > 90:
            lat_float = 90.0
            
        return f"{lat_float:.6f}"
    except (ValueError, TypeError):
        return '0.0'

def clean_ptkp(ptkp):
    """Clean and format PTKP values"""
    if ptkp is None or ptkp == '' or str(ptkp).strip() == '':
        return None
    
    try:
        # Convert to string and trim whitespace
        ptkp_str = str(ptkp).strip()
        
        if not ptkp_str:
            return None
        
        # Check if it contains TK or K followed by a number
        # Add "/" before the number if pattern matches TK2, K2, etc.
        ptkp_clean = re.sub(r'(TK|K)(\d)', r'\1/\2', ptkp_str, flags=re.IGNORECASE)
        
        return ptkp_clean
        
    except Exception:
        return None

def clean_longitude(lon):
    """Clean and validate longitude values. Valid range: -180 to 180"""
    if lon is None or lon == '' or str(lon).strip() == '':
        return '0.0'
    
    try:
        # Convert to string and clean
        lon_str = str(lon).strip()
        
        # Remove any non-numeric characters except minus sign and decimal point
        lon_clean = re.sub(r'[^\d\-\.]', '', lon_str)
        
        if not lon_clean or lon_clean == '-' or lon_clean == '.':
            return '0.0'
        
        # Convert to float for validation
        lon_float = float(lon_clean)
        
        # Validate longitude range (-180 to 180)
        if lon_float < -180:
            lon_float = -180.0
        elif lon_float > 180:
            lon_float = 180.0
            
        return f"{lon_float:.6f}"
        
    except (ValueError, TypeError):
        return '0.0'

def process_data(input_filename=None):
    """Process outlet data with improved error handling and validation"""
    if input_filename is None:
        input_file = 'template_isian_database_NEW_LXC.json'  # default filename
    else:
        input_file = input_filename
    
    # Generate output filename based on input filename
    if input_file.endswith('.json'):
        # Remove .json extension properly
        base_name = input_file[:-5]
        # Ensure we have a valid base name
        if not base_name or base_name.strip() == '':
            base_name = 'data'
        output_file = f"cleaned-{base_name}.xlsx"
    else:
        # For non-json files, just add the extension
        base_name = input_file
        if not base_name or base_name.strip() == '':
            base_name = 'data'
        output_file = f"cleaned-{base_name}.xlsx"
      # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        return
    
    # Validate output filename
    try:
        # Test if we can create the output file path
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
    except Exception as e:
        print(f"Error: Cannot create output directory: {e}")
        return
    
    try:
        # Load JSON data
        with open(input_file, 'r', encoding='utf-8') as file:
            data = json.load(file)
        
        if not isinstance(data, list):
            print("Error: JSON file should contain an array of objects.")
            return
        
        print(f"Processing {len(data)} records...")
        
        # Process each record
        for i, record in enumerate(data):
            if not isinstance(record, dict):
                print(f"Warning: Record {i+1} is not a valid object, skipping...")
                continue
                
            # Clean latitude and longitude
            if 'latitude' in record:
                original_lat = record['latitude']
                record['latitude'] = clean_latitude(record['latitude'])
                if original_lat and str(original_lat).strip() and record['latitude'] == '0.0':
                    print(f"Warning: Invalid latitude '{original_lat}' in record {i+1}, set to 0.0")
                    
            if 'longitude' in record:
                original_lon = record['longitude']
                record['longitude'] = clean_longitude(record['longitude'])
                if original_lon and str(original_lon).strip() and record['longitude'] == '0.0':
                    print(f"Warning: Invalid longitude '{original_lon}' in record {i+1}, set to 0.0")
              # Clean string fields
            if 'name' in record and record['name']:
                record['name'] = str(record['name']).strip()

            if 'phone' in record and record['phone']:
                record['phone'] = str(record['phone']).strip()

            if 'client' in record and record['client']:
                record['client'] = str(record['client']).strip()

            if 'outlet' in record and record['outlet']:
                record['outlet'] = str(record['outlet']).strip()

            if 'rekening' in record and record['rekening']:
                record['rekening'] = str(record['rekening']).strip()

            if 'kk' in record and record['kk']:
                record['kk'] = str(record['kk']).strip()
            
            if 'ptkp' in record:
                record['ptkp'] = clean_ptkp(record['ptkp'])
                
            if 'email' in record:
                if record['email'] is None or record['email'] == '' or str(record['email']).strip() == '':
                    record['email'] = None
                else:
                    email = str(record['email']).strip().lower()
                    # Basic email validation
                    if '@' in email and '.' in email:
                        record['email'] = email
                    else:
                        # Add @gmail.com if email format is invalid
                        record['email'] = email + '@gmail.com'
                        print(f"Warning: Invalid email format, added @gmail.com to '{email}' in record {i+1}")
        
        # Convert to DataFrame and save
        df = pd.DataFrame(data)
        df.to_excel(output_file, index=False)
        print(f"Data has been cleaned and saved to '{output_file}'")
        print(f"Total records processed: {len(df)}")
        
        # Display some statistics
        if 'latitude' in df.columns and 'longitude' in df.columns:
            valid_coords = df[(df['latitude'] != '0.0') & (df['longitude'] != '0.0')]
            print(f"Records with valid coordinates: {len(valid_coords)}")
        
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON format in '{input_file}': {e}")
    except Exception as e:
        print(f"Error processing data: {e}")

if __name__ == "__main__":
    # Check if filename is provided as command line argument
    if len(sys.argv) > 1:
        filename = sys.argv[1]
        print(f"Processing file: {filename}")
        process_data(filename)
    else:
        print("Usage: python datacleansing.py <json_filename>")
        print("Example: python datacleansing.py data.json")
        print("Using default file: Aqua haier.json")
        process_data()