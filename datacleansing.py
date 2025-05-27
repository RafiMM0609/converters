import json
import pandas as pd
import re

def clean_latitude(lat):
    lat = str(lat)
    lat = re.sub(r'[^\d\-]', '', lat)
    if not lat or (lat == '-'):
        return '0.0'
    is_negative = lat.startswith('-')
    if is_negative:
        lat = lat[1:]  # Remove minus sign for processing
    
    if not lat:
        return '0.0'
    first_digit = lat[0]
    rest_digits = lat[1:] if len(lat) > 1 else '0'
    cleaned_lat = f"{'-' if is_negative else ''}{first_digit}.{rest_digits}"
    return cleaned_lat

def clean_longitude(lon):
    lon = str(lon)
    lon = re.sub(r'[^\d\-]', '', lon)
    if not lon or (lon == '-'):
        return '0.0'
    is_negative = lon.startswith('-')
    if is_negative:
        lon = lon[1:]  # Remove minus sign for processing
    if not lon:
        return '0.0'
    if len(lon) < 2:
        return f"{'-' if is_negative else ''}{lon}.0"
    first_part = lon[:3] if lon[0] == '1' else lon[:2]
    rest_digits = lon[len(first_part):] if len(lon) > len(first_part) else '0'
    cleaned_lon = f"{'-' if is_negative else ''}{first_part}.{rest_digits}"
    
    return cleaned_lon

def process_data():
    with open('template_outlet.json', 'r') as file:
        data = json.load(file)
    for record in data:
        # Clean latitude and longitude
        record['latitude'] = clean_latitude(record['latitude'])
        record['longitude'] = clean_longitude(record['longitude'])
    df = pd.DataFrame(data)
    df.to_excel('cleaned_outlet_data.xlsx', index=False)
    print("Data has been cleaned and saved to cleaned_outlet_data.xlsx")

if __name__ == "__main__":
    process_data()