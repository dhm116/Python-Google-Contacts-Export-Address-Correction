#!/usr/bin/env python
'''
Copyright (c) 2010 Doug McCall [dhm116@dougmccall.com]

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
of the Software, and to permit persons to whom the Software is furnished to do
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''

from Tkinter import *

import csv, tkFileDialog

master = Tk()
master.withdraw()

# Get the CSV file
filename = tkFileDialog.askopenfilename(filetypes=[('CSV Files', '.csv')])

people = open(filename, "rb")

rows = people.readlines()

# We've got the rows, so close the actual file
people.close()

# Remove the header row so it's not processed
header = rows.pop(0);

# Start a list of corrected addresses
fixed = []

# Iterate through the rows of data
for row in rows:
    # Store this row in something else in case we can't directly manipulate the row data
    data = row
    
    # This should probably be corrected, but for now just arbitrarily pick some length
    # for the contact line to be longer than
    if len(row) > 50:
        #print "Adding new address line"
        
        # If we already have some corrected addresses, append a new line
        # * This probably isn't necessary any more now that I'm feeding the
        #   fixed data as a list to the CSV reader
        if len(fixed) > 0:
            index = len(fixed) - 1
            fixed[index] = ''.join((fixed[index], '\r\n'))
            
        # Add the data to the list of corrected data with any carriage returns removed
        fixed.append(data.rstrip())
        
        #print "\t" + fixed[len(fixed) - 1]
        
    else:
        # We've found a fragmented address line
        
        #print "\tAdding split address line"
        
        # Strip off the carriage return at the end
        data = data.rstrip()
        
        #print "\t\t" + data
        
        # Join this fragmented address line to the previous line of data
        index = len(fixed) - 1
        fixed[index] = ', '.join((fixed[index], data))
        
        #print "\tFixed version"
        #print "\t\t" + fixed[index]
        
        #if data.count(',,,,,,,,,,,,,,,,,,,,,,,') == 0:
        #    print 'Looking for the next full address'
        #    row_start = True

# Put the header row back in the corrected data
fixed.insert(0,header)

# Load the corrected data in to the CSV reader to more easily manipulate the columns
spamReader = csv.reader(fixed, delimiter=',')
 
# Ask the user where to save the resulting data    
filename = tkFileDialog.asksaveasfilename(filetypes=[('CSV Files', '.csv')])

f2 = open(filename, 'w')

writer = csv.writer(f2, delimiter=',')

rownum = 0
for row in spamReader:
    if rownum > 0:
        # Print out the fully corrected address for debugging
        print str(rownum) + ": " + row[36]
        
        # Isolate the address parts
        addr = row[36].split(',')
        
        # Obtain the State and Zip code from the last part of the address
        state_zip = addr.pop().split()
                
        # If we're not using state abbreviations and there is a space in the state
        # name, we need to merge those back together
        if len(state_zip) > 2:
            state_zip[0] = state_zip[0] + " " + state_zip[1]
            
            # Get rid of the unneeded parts from the merge
            state_zip.pop(1)
        
        # Combine the state and zip back in with the address
        addr.extend(state_zip)
        #print addr
        
        # Set the address column data
        row[37] = addr[0]
        
        # Set the city
        row[38] = addr[-3].lstrip()
        
        # If there is a secondary address (apt. num, etc.)
        if len(addr) > 4:
            row[39] = addr[1]
            
        # Set the state
        row[40] = addr[-2]
        
        # Set the zip code
        row[41] = addr[-1]
        
        #print row
        
    writer.writerow(row)
    rownum += 1
    
f2.close() 