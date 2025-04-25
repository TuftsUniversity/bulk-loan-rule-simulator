import os
import pandas as pd
import re
from flask import Blueprint, request, redirect, url_for, send_file, current_app, render_template
from werkzeug.utils import secure_filename
import io
from io import BytesIO
import zipfile
import dotenv
import os
from dotenv import load_dotenv
import json
import requests
import time
import pymarc as pym
import sys
sys.path.append(os.path.relpath('config/'))
import secrets_local
############################################################################
############################################################################
####
####    Title:  createManagedSetsForItemPolicyChange.py
####    Author: Henry Steele, Senion Systems Librarian, Library Technology Services, Tufts University
####
####    Purpose:
####        Ingest a file of old:new item policy mappings, and create a series of managed sets that contain all
####        of the items that should undergo this change
####        
####        The goal is that the managed set will have a name that completely describes what changes should be
####        effected on the associated items 
####    
####    Method:
####        - ingest file
####        - identify current item policy
####        - identify library
####        - create managed set query
####            - conditions expressed as XML (see below)
####            - name as "<current item_policy> - <locations> - to - <new item policy> and <current item policy> as item material type"

'''
<set>
    <name>Test set for loan rule simulator 8</name>
    <type desc="Logical">LOGICAL</type>
    <content desc="Physical items">ITEM</content>
     <query desc="Physical items where (Current library equals ((Tisch Library : Tisch DVD Collections)) AND Item policy equals &quot;Video DVD&quot;)">ITEM where ITEM (current_Location OUTER_EQUAL (SMFA : SMFAS) AND ITEM (itemPolicy OUTER_EQUAL "VideoDVD")</query>
</set>


OR

<set>
    <name>Test set for loan rule simulator 10</name>
    <type desc="Logical">LOGICAL</type>
    <content desc="Physical items">ITEM</content>
     <query desc="Physical items where (Current library equals ((Tisch Library : Tisch DVD Collections)) AND Item policy equals &quot;Video DVD&quot;)">ITEM where ITEM (current_Location OUTER_EQUAL (TISCH : TISCHDVD) AND ITEM (itemPolicy OUTER_EQUAL "VideoDVD")</query>
</set>

'''

from tkinter.filedialog import askopenfilename


inputFilename = askopenfilename(title="Select Excel file containing current item policy and new item policy to map to")

df = pd.read_excel(inputFilename, engine="openpyxl", dtype='str')
# Group by and join multiple locations into semicolon-separated fields
# Static parts of the XML structure
set_creation_body_literal_1 = "<set><name>"

set_creation_body_literal_2 = '''</name>
    <type desc="Logical">LOGICAL</type>
    <content desc="Physical items">ITEM</content>
    <query desc="where (Item policy equals &quot;'''

set_creation_body_literal_3 = '''&quot; AND Current library equals (('''

set_creation_body_literal_4 = ''' : All)))">
    ITEM where ITEM ((itemPolicy OUTER_EQUAL "'''

set_creation_body_literal_5 = '''"))</query>
</set>'''

headers = {'Content-Type': 'application/xml', 'Accept': 'application/json'}

log_file = open("Result of Creating Sets for Sandbox.csv", "w+")
log_file.write("Library, Current Item Policy, Loan Length, Set Name, Result\n")

for index, row in df.iterrows():
    item_policy = row['Item Policy']
    item_policy_code = row['Item Policy Code']
    library = row['Library']
    library_code = row['Library Code']
    loan_length = row['Loan Length']

    # Build the combined name for the set
    name = f"{item_policy} - {library} - to {loan_length} and Item Material Type {item_policy}"

    # Build the full body
    body = (
        set_creation_body_literal_1
        + name
        + set_creation_body_literal_2
        + item_policy
        + set_creation_body_literal_3
        + library
        + set_creation_body_literal_4
        + item_policy_code
        + set_creation_body_literal_5
    )


    print(secrets_local.alma_sandbox_set_api_url + secrets_local.alma_sandbox_configuration_api_key)
    print("\n")
    print(body)
    
    result = requests.post(secrets_local.alma_sandbox_set_api_url + secrets_local.alma_sandbox_configuration_api_key, headers=headers, data=body)

    if result.status_code == 200 or result.status_code == 201 or result.status_code == 202 or result.status_code == 203 or result.status_code == 204:
        
        log_file.write(library + "," + item_policy + "," + loan_length + "," + name + "," + str(result.status_code) + "\n")


    else:
        log_file.write(library + "," + item_policy + "," + loan_length + "," + name + "," + str(result.text) + "\n")


log_file.close()


        


    





