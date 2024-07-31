import pandas as pd
from parsons.ngpvan.van import VAN
import os
from parsons.google.google_sheets import GoogleSheets
import json
from dotenv import load_dotenv

load_dotenv()

# pull in the data from google sheets
def get_worksheet(spreadsheet_id, worksheet):
    """
    Given a sheet fileID, return the worksheet columns and rows as a Parsons Table
    """
    credential_filename = 'googlecredential.json'
    credentials = json.load(open(credential_filename))
    sheets = GoogleSheets(google_keyfile_dict=credentials)
    result_table = sheets.get_worksheet(spreadsheet_id,worksheet)
    return result_table

# using the spreadsheetid and worksheet name, call the function to get the gs data
gs_data = get_worksheet(spreadsheet_id='1Zdd9ruj0Ur9_iOmY90T2EvOT2zm3NQlmrJ8FRTyLSQ8', worksheet='Sheet1')
print(gs_data)

df = pd.DataFrame(gs_data)
print(df)

# establish VAN connection
van_api_key = os.getenv('van_api_key')
van = VAN(api_key=van_api_key, db='MyCampaign')

# use the Parsons apply activist code function to upload the activist codes
# if you want to include a canvass result, set omit_contact to False
for index, row in df.iterrows():
    vanId = row['vanid']
    activist_code_id = row['activist_code_id']

    van.apply_activist_code(vanId, activist_code_id, omit_contact=True)

