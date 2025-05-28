#!/usr/bin/env python3
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
from tkinter.filedialog import askopenfilename
############################################################################
############################################################################
####
####    Title:  gapAnalysis.py
####    Author: Henry Steele, Senion Systems Librarian, Library Technology Services, Tufts University
####    Purpose
####        ingest three files from bulk loan rule testing, loan rule report export, and mapping of old item
####        policies to new item policies, and apply a column in the two former reports that contain the new item policy
####        field that will be written to the old item policy
####        
####        Then analyze the bulk loan rule testing file to see if applying this new item policy, which usually
####        aims to simplify the organization of loan rules, will have any unexecpted consequences
####    Input:
####        Ingest three files via file picker
####         - Bulk_Checkout_Request_Results - Formatted.xlsx from the bulk loan rule tester
####         - loan rule report from getLoanRules.py
####         - mapping of old item policies to new item policies
####    Method:
####       - after ingesting 3 files, sort bulk loan rule tester exporter formatted by (in order)
####          - location
####          - new item policy
####          - user group
####       - then group these by those fields, iteratively for each group
####       - the old item policies in this group will likely differ within each of these, but what this script
####         seeks to identify is within these groups, are there more than one loan rule and TOU/request rule and TOU
####         because this would be unexpected.  It may not be unwanted, but at least should be identified


inputFilenameBulkTestingFormatted = askopenfilename(title="Select Excel file Bulk_Checkout_Request_Results - Formatted.xlsx")
inputFilenameLoanRuleinFulfillmentUnit = askopenfilename(title="Select Excel file containing loan rules for Fulfillment Unit you want to add new item policy to")
inputFilenameMapping = askopenfilename(title="Select Excel file containing mapping of old item policy to new item policy")

df_tester = pd.read_excel(inputFilenameBulkTestingFormatted, engine="openpyxl", dtype='str')
df_loan_rules = pd.read_excel(inputFilenameLoanRuleinFulfillmentUnit, engine="openpyxl", dtype='str')
df_mapping = pd.read_excel(inputFilenameMapping, engine="openpyxl", dtype='str')

df_tester = df_tester.sort_values(by=[])

''' 
make a new column in both sheets.  the easiest is going to be formatted.  loan rule report will rely on item policy being in separate column

the new column will apply the new item policy where old item policy matches item policy in other sheet

unique_old_item_policies = 
df_tester[df_tester[]]
df_tester['New Item Policy'] = 
'''