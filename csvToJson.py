import xlrd
from collections import OrderedDict
import simplejson as json

# This python script takes in an excel spread sheet formatted terms / descriptions
# And generates the glossary config typescript constant used by the glossary page
# To run this, open cmd, navigate to this folder, run "python csvToJson.py"

inputFile = 'Terms.xlsx'
outputFile = 'glossary-config.ts'

# Open the workbook and select the first worksheet
wb = xlrd.open_workbook(inputFile)
sh = wb.sheet_by_index(0)

# List to hold dictionaries
glossary_terms_list = []

# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    terms = OrderedDict()
    row_values = sh.row_values(rownum)
    terms['term'] = row_values[0]
    terms['description'] = row_values[1]
    glossary_terms_list.append(terms)

alphabet_dict = dict()

# Iterate over terms and catogorize them into A-Z subsections
for term in glossary_terms_list:
    first_letter = term['term'][0] 
    if first_letter not in alphabet_dict:
        alphabet_dict[first_letter] = []
    alphabet_dict[first_letter].append(term)

# Serialize the list of dicts to JSON
j = json.dumps(alphabet_dict)

# Write to file
with open(outputFile, 'w') as f:
    f.write("export const GLOSSARY_CONFIG = ")
    f.write(j)
    f.write(";")
