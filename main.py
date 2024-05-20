import pandas as pd
import googlemaps
import time
import os
from dotenv import load_dotenv
load_dotenv()

api_key = os.getenv('GOOGLE_MAPS_API_KEY')
gmaps = googlemaps.Client(key=api_key)
origin = "Your office location"

df = pd.read_excel('pg_list.xlsx')

addresses = df.iloc[:, 0]

filtered_data = []

for index, address in addresses.items():
    destination = str(address)
    
    results = gmaps.distance_matrix(origin, destination)
    
    try:
        distance_text = results['rows'][0]['elements'][0]['distance']['text']
        distance = float(distance_text.split()[0])
        print(distance)
        row = df.iloc[index].copy()
        row['Distance from Office (km)'] = distance
        filtered_data.append(row)
    except (KeyError, IndexError, ValueError) as e:
        print(f"Error processing address '{destination}': {e}")
    
    # time.sleep(1)

filtered_df = pd.DataFrame(filtered_data)
# Sort
filtered_df = filtered_df.sort_values(by='Distance from Office (km)', ascending=True)
#Save
filtered_df.to_excel('filtered_pg_list.xlsx', index=False)



