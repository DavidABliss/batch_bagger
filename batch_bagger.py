# -*- coding: utf-8 -*-

import bagit
import sys
import uuid
import csv
import os
import re
import argparse

parser = argparse.ArgumentParser(description='Bag a list of folders in a directory, substituting values from a CSV into a bag-info template')
parser.add_argument('-d', '--directory', help="directory containing folders to be bagged", action="store", dest="d", default=os.getcwd())
parser.add_argument('-b', '--baginfo', required=True, help="bag-info.txt file to be distributed into each bag", action="store", dest="b")
parser.add_argument('-c', '--csv', required=True, help="CSV file containing the names of folders to be bagged, and any values to be substituted in the bag-info file", action="store", dest="c")
parser.add_argument('-v', '--verbose', required=False, help="print the contents of each bag-info file as it is written", action="store_true", dest="v")

args = parser.parse_args()


# Require the users to provide a bags folder, a bag-info text template file, and a CSV spreadsheet
# At a minimum, the spreadsheet must list the bags in the bag folder. 
try:
    bagsDir = args.d
    bagInfo = args.b
    spreadSheet = args.c
except:
    sys.exit("Usage: python batch_bagger.py <path to bags folder> <path to bag-info txt template> <path to bag-info CSV or XLSX spreadsheet>")

# Allow users to add a -v switch, which will print each bag-info file as it is written
try:
    if args.v:
        verbose = True
except:
    verbose = False
    
# Define the text strings that will be recognized as bag-info field labels
fieldsList = ['Source-Organization', 'Organization-Address', 'Contact-Name', 'Contact-Phone', 'Contact-Email', 'External-Description', 
              'External-Identifier', 'Internal-Sender-Description', 'Internal-Sender-Identifier', 'Rights-Statement', 'Bag-Group-Identifier']

if spreadSheet.endswith('.csv'):
    with open(spreadSheet, 'r', encoding='utf-8') as bag_spreadsheet:
        spreadsheet_reader = csv.reader(bag_spreadsheet, delimiter=',', quotechar='"')
        
        # Store the column headers in the first line of the spreadsheet
        replaceFields = spreadsheet_reader.__next__()
        
        for entry in spreadsheet_reader: 
            # Iterate over the spreadsheet, skipping the column headers
            if spreadsheet_reader.line_num == 0:
                next
                
            # Store the contents of each row, and create a dictionary of column
            # headers and values that will replace placeholder text in the bag-info
            # template file
            rowList = entry
            replacementDict = {}
            for i in range(len(replaceFields)):
                replacementDict[replaceFields[i]] = rowList[i]
    
            # Iterate over the template file, assembling a bag-info dictionary for each entry
            baginfoDict = {}
            # Generate a UUID to be used as an External-Identifier later
            bagUUID = str(uuid.uuid4())
            with open(bagInfo, 'r', encoding='utf-8') as templateFile:
                for line in templateFile:
                    line = line.lstrip()
                    
                    # Skip blank lines, to avoid duplicating the final field in the 
                    # bag-info template
                    if line == '':
                        next
                        
                    # If the line contains a colon, check, to see if the line begins
                    # with one of the labels found in the fieldsList
                    if line.split(':')[0] in fieldsList:
                            # If the label is already a key in the bag-info dictionary, 
                            # append the new value to the existing value, to combine 
                            # values of repeated fields
                            if line.split(':')[0] in baginfoDict:
                                label = line.split(':')[0]
                                value = ':'.join(line.split(':')[1:])
                                value = value.lstrip()
                                baginfoDict[label] = baginfoDict[label] + ' | ' + value
                                
                            # If the label is not already in the dictionary, create a new
                            # dictionary entry with the label and value
                            else:
                                label = line.split(':')[0]
                                value = ':'.join(line.split(':')[1:])
                                value = value.lstrip()
                                baginfoDict[label] = value
                    
                    # If the text before the colon is not found in the fieldsList,
                    # it is a continuation of the previous line's text, and just
                    # happens to contain a colon. In this case, add the line to 
                    # the previous dictionary entry                
                    elif not line.split(':')[0] in fieldsList:
                        baginfoDict[label] = baginfoDict[label] + line
            
            # Add the UUID External-Identifier to the bag-info dictionary. If other
            # External-Identifier values are present, insert the UUID first.
            if 'External-Identifier' in baginfoDict:
                baginfoDict['External-Identifier'] = bagUUID + ' | ' + baginfoDict['External-Identifier']
            else:
                baginfoDict['External-Identifier'] = bagUUID
                
            # Iterate over the bag-info dictionary, replacing keywords with corresponding
            # values from the spreadsheet. Also remove trailing white space in each value,
            # replace line break characters with spaces, and replace double spaces with
            # single spaces.
            for key, value in baginfoDict.items():
                value = value.rstrip()
                value = value.replace('\n', ' ')
                value = value.replace('  ', ' ')
                
                # Look for text in the bag-info dictionary wrapped in double brackets,
                # matching the column headers in the spreadsheet. Replace that text with
                # the corresponding text in the spreadsheet. Update the bag-info dictionary
                # with the updated text.
                for label, replacementValue in replacementDict.items():
                    value = re.sub('\[\[' + label + '\]\]', replacementValue, value, flags=re.IGNORECASE)
                    baginfoDict[key] = value
            
            # Set the UUIDs.txt file path
            outPath = os.path.join(bagsDir, 'UUIDs.csv')
            
            # Name the bag being created
            print('Bagging: ' + rowList[0])
            
            # If the user has added a -v switch, print the bag-info dictionary
            if verbose:
                for key in baginfoDict:
                    print(key + ': ' + baginfoDict[key])
                    
            # Create the bag at the bag directory, using the first column of the spreadsheet                
            bagPath = os.path.join(bagsDir, rowList[0])
            bag = bagit.make_bag(bagPath, baginfoDict, checksum=['sha256']) # Specify a sha256 hash algorithm
            
            # After the bag has been written, write the UUIDs.txt file
            with open(outPath, 'a+', encoding='utf-8', newline='') as guids_file:
                uuidWriter = csv.writer(guids_file, delimiter=',', quotechar='"')
                outRow = [rowList[0], bagUUID]
                uuidWriter.writerow(outRow)

elif spreadSheet.endswith('.xlsx'):
    
    from openpyxl import load_workbook
    
    # Open the workbook at read only
    wb = load_workbook(spreadSheet, read_only=True)
    sheetName = wb.sheetnames[0]
    ws = wb[sheetName]
    
    # Store the column headers in the first line of the spreadsheet
    firstrow = ws[1]
    replaceFields = []
    for cell in firstrow:
        replaceFields.append(cell.value)
    
    # Iterate over the remaining rows, storing the cell values in a list
    for row in ws[2: ws.max_row]:
        rowList = []
        for cell in row:
            rowList.append(cell.value)
        # Store the contents of each row, and create a dictionary of column
        # headers and values that will replace placeholder text in the bag-info
        # template file
        replacementDict = {}
        for i in range(len(replaceFields)):
            replacementDict[replaceFields[i]] = rowList[i]

        # Iterate over the template file, assembling a bag-info dictionary for each entry
        baginfoDict = {}
        # Generate a UUID to be used as an External-Identifier later
        bagUUID = str(uuid.uuid4())
        with open(bagInfo, 'r', encoding='utf-8') as templateFile:
            for line in templateFile:
                line = line.lstrip()
                
                # Skip blank lines, to avoid duplicating the final field in the 
                # bag-info template
                if line == '':
                    next
                    
                # If the line contains a colon, check, to see if the line begins
                # with one of the labels found in the fieldsList
                if line.split(':')[0] in fieldsList:
                        # If the label is already a key in the bag-info dictionary, 
                        # append the new value to the existing value, to combine 
                        # values of repeated fields
                        if line.split(':')[0] in baginfoDict:
                            label = line.split(':')[0]
                            value = ':'.join(line.split(':')[1:])
                            value = value.lstrip()
                            baginfoDict[label] = baginfoDict[label] + ' | ' + value
                            
                        # If the label is not already in the dictionary, create a new
                        # dictionary entry with the label and value
                        else:
                            label = line.split(':')[0]
                            value = ':'.join(line.split(':')[1:])
                            value = value.lstrip()
                            baginfoDict[label] = value
                
                # If the text before the colon is not found in the fieldsList,
                # it is a continuation of the previous line's text, and just
                # happens to contain a colon. In this case, add the line to 
                # the previous dictionary entry                
                elif not line.split(':')[0] in fieldsList:
                    baginfoDict[label] = baginfoDict[label] + line
        
        # Add the UUID External-Identifier to the bag-info dictionary. If other
        # External-Identifier values are present, insert the UUID first.
        if 'External-Identifier' in baginfoDict:
            baginfoDict['External-Identifier'] = bagUUID + ' | ' + baginfoDict['External-Identifier']
        else:
            baginfoDict['External-Identifier'] = bagUUID
            
        # Iterate over the bag-info dictionary, replacing keywords with corresponding
        # values from the spreadsheet. Also remove trailing white space in each value,
        # replace line break characters with spaces, and replace double spaces with
        # single spaces.
        for key, value in baginfoDict.items():
            value = value.rstrip()
            value = value.replace('\n', ' ')
            value = value.replace('  ', ' ')
            
            # Look for text in the bag-info dictionary wrapped in double brackets,
            # matching the column headers in the spreadsheet. Replace that text with
            # the corresponding text in the spreadsheet. Update the bag-info dictionary
            # with the updated text.
            for label, replacementValue in replacementDict.items():
                value = re.sub('\[\[' + label + '\]\]', replacementValue, value, flags=re.IGNORECASE)
                baginfoDict[key] = value
        
        # Set the UUIDs.txt file path
        outPath = os.path.join(bagsDir, 'UUIDs.csv')
        
        # Name the bag being created
        print('Bagging: ' + rowList[0])
        
        # If the user has added a -v switch, print the bag-info dictionary
        if verbose:
            for key in baginfoDict:
                print(key + ': ' + baginfoDict[key])
                
        # Create the bag at the bag directory, using the first column of the spreadsheet                
        bagPath = os.path.join(bagsDir, rowList[0])
        bag = bagit.make_bag(bagPath, baginfoDict, checksum=['sha256']) # Specify a sha256 hash algorithm
        
        # After the bag has been written, write the UUIDs.txt file
        with open(outPath, 'a+', encoding='utf-8', newline='') as guids_file:
            uuidWriter = csv.writer(guids_file, delimiter=',', quotechar='"')
            outRow = [rowList[0], bagUUID]
            uuidWriter.writerow(outRow)
