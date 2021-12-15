# -*- coding: utf-8 -*-
import bagit
import sys
import uuid
import csv
import os
import re


# Require the users to provide a bags folder, a bag-info text template file, and a CSV spreadsheet
# At a minimum, the spreadsheet must list the bags in the bag folder. 
try:
    bagsDir = sys.argv[1]
    bagInfo = sys.argv[2]
    spreadSheet = sys.argv[3]
except:
    sys.exit("Usage: python batch_bagger.py <path to bags folder> <path to bag-info txt template> <path to bag-info CSV spreadsheet>")

# Allow users to add a -v switch, which will print each bag-info file as it is written
try:
    if sys.argv[4] == '-v':
        verbose = True
except:
    verbose = False
    
# Define the text strings that will be recognized as bag-info field labels
fieldsList = ['Source-Organization', 'Organization-Address', 'Contact-Name', 'Contact-Phone', 'Contact-Email', 'External-Description', 
              'External-Identifier', 'Internal-Sender-Description', 'Internal-Sender-Identifier', 'Rights-Statement', 'Bag-Group-Identifier']

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
        txtPath = os.path.join(bagsDir, 'UUIDs.txt')
        
        # Name the bag being created
        print('Bagging: ' + entry[0])
        
        # If the user has added a -v switch, print the bag-info dictionary
        if verbose:
            for key in baginfoDict:
                print(key + ': ' + baginfoDict[key])
                
        # Create the bag at the bag directory, using the first column of the spreadsheet                
        bagPath = os.path.join(bagsDir, entry[0])
        bag = bagit.make_bag(bagPath, baginfoDict, checksum=['sha256']) # Specify a sha256 hash algorithm
        
        # After the bag has been written, write the UUIDs.txt file
        with open(txtPath, 'a+') as guids_file:
            guids_file.write(bagUUID + ' - ' + entry[0] + '\n')
