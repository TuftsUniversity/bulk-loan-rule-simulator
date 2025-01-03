# Bulk Loan Rule Simulator
- The goal of this script is to allow you to completely automate the testing of all possible fulfillment unit rule scenarios at your institution captured by your fulfillment unit rules (loan and request)
- this way after changes are made to item records that could affect loan rule behavior, such as changes to item policies, or item locations, every single scenario can be tested automatically so you can be sure no changes come as a result, or onyl desired changes

  
## Input
- gather your unique item policy and location combinations, by using a report suchn as this: 	/shared/Community/Reports/Institutions/Tufts University/	Item Policy and Item Location Combinations with At Least 1 Active Item
-   - Export as Excel
    - this file picks one examplar barcode from every unqiue item policy/location combination that has items
- Also create a user group Excel file
-   - we only have 4 or 5 meaningful groupings of users, and one can be picked to represent each
    - should have these columns
      - User Group
      - Primary Identifier
- put these in the following following folders
  - input/Item Policies and Locations
  - input/User Groups 

## Set Up
- change the secrets_local_example.py file to have real API keys etc, and change its name to secrets_local.py
## Operation

- This may take a while to run, so you can run the headless version of this
- `python3 drillFulfillmentConfigurationUtility-Headless.py`
- note at this point if you are changing from sandbox to prod you'll have to hardcode the changes to the call to the secrets file
  
