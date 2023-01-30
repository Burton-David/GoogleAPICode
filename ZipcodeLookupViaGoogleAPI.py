import pandas as pd
import requests


def get_zipcode(address, company_name, city):
    # check if address is present
    if pd.notna(address):
        # make request to Google Maps Geocoding API
        response = requests.get(
            f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key=YOUR_API_KEY")
        # parse the response
        data = response.json()
        # get the first result
        result = data['results'][0]
        # get the zipcode from the result
        zipcode = result['address_components'][-1]['long_name']
        return zipcode
    elif pd.notna(company_name) and pd.notna(city):
        # make request to Google Maps Geocoding API with company name and city
        response = requests.get(
            f"https://maps.googleapis.com/maps/api/geocode/json?address={company_name}, {city}&key=YOUR_API_KEY")
        # parse the response
        data = response.json()
        # get the first result
        result = data['results'][0]
        # get the zipcode from the result
        zipcode = result['address_components'][-1]['long_name']
        return zipcode
    else:
        # return None if no address or company name and city is present
        return None


# load the excel sheet into a pandas dataframe
df = pd.read_excel("filename.xlsx")

# add a new column for zipcode
df['ZIPCODE'] = df.apply(lambda row: get_zipcode(
    row['ADDRESS'], row['COMPANY NAME'], row['CITY']), axis=1)

# save the updated dataframe to a new excel file
df.to_excel("output.xlsx", index=False)
