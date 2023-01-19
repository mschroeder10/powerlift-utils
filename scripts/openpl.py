import argparse
import requests
import pandas as pd
import re
import pygsheets
import sys

BASE_URL = "https://www.openpowerlifting.org/u/" 
cols = ['Last Name','First Name','openpowerlifting link', 'instagram', 'raw squat','raw bench','raw DL','raw total', 'eq squat','eq bench','eq DL','eq total', 'notes']

def get_data(name):
    """
    scrape data from openpowerlifting
    input: firstnamelastname 
    """
    url = BASE_URL + name
    try:
        r = requests.get(url)
        df_list = pd.read_html(r.text) # this parses all the tables in webpages to a list
        df = df_list[0]
        df = df.dropna(axis=1, how='all')
        return df
    except ValueError as e:
        return pd.DataFrame()

def build_df(first, last):
    """
    convert openpowerlifting format to kate spreadsheet format
    """
    result = {}
    df = pd.DataFrame(columns = cols)
    result['First Name'] = first
    result['Last Name'] = last
    
    # no special characters for openpl link lol
    first = re.sub(r'[^a-zA-Z0-9]', '', first)
    last = re.sub(r'[^a-zA-Z0-9]', '', last)
    name = first.lower() + last.lower()
    
    openpl_df = get_data(name)
    if not openpl_df.empty:
        raw = openpl_df[openpl_df['Equip'] == 'Raw'].drop(columns = ['Equip', 'Dots'])
        eq = openpl_df[openpl_df['Equip'] == 'Single'].drop(columns = ['Equip', 'Dots'])

        # lol this could be better but whatever
        result['openpowerlifting link'] = BASE_URL + name
        
        if not raw.empty:
            result['raw squat'] = raw['Squat'].iloc[0]
            result['raw bench'] = raw['Bench'].iloc[0]
            result['raw DL'] = raw['Deadlift'].iloc[0]
            result['raw total'] = raw['Total'].iloc[0]
        if not eq.empty:
            result['eq squat'] = eq['Squat'].iloc[0]
            result['eq bench'] = eq['Bench'].iloc[0]
            result['eq DL'] = eq['Deadlift'].iloc[0]
            result['eq total'] = eq['Total'].iloc[0]
    df = df.append(result, ignore_index=True)
    return df

def main():
    parser = argparse.ArgumentParser(description='openpl to google sheets')
    parser.add_argument('--credentials', dest="credentials")
    parser.add_argument('--sheet', dest="sheet")
    args = parser.parse_args()
    # load roster
    xls = pd.ExcelFile('cnats2023.xlsx')
    roster_df = pd.read_excel(xls, 'Roster')
    roster_df = roster_df[["Last Name", "First Name"]]

    #populate our big sheet
    df = pd.DataFrame(columns = cols)
    for index, row in roster_df.iterrows():
        first = row['First Name']
        last = row['Last Name']
        df = df.append(build_df(first, last))
    
    df.to_excel('bestlifts.xlsx', index = False)

    # if we passed in google sheets credentials 
    if args.credentials and args.sheet:
        gc = pygsheets.authorize(service_file=args.credentials) # open sheet
        sh = gc.open(args.sheet)

        #select the best lifters sheet
        index = -1
        count = 0
        for s in sh:
            if s.title == "best lifts test":
                index = count
            count += 1
            
        if index >= 0:
            wks = sh[index]
            wks.set_dataframe(df,(1,1), copy_index=False, nan="")

if __name__ == "__main__":
    sys.exit(main())