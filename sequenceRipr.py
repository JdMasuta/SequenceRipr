#
#
#   Name: SequenceRipr  Rev: 2.0
#
#
#   Description:
#   This code will automatically find the 'Z' numbers and associated sequence numbers in the sequence sheet, and creates a 'Project' class instance that holds each item (Z#),
#       which in turn holds each sequence number for that item.
#
#   This is a python script that is to be used in conjunction with the Neptyne spreadsheet application (NOT the extension for Google Sheets).
#       To use this script, import a sequence sheet into app.neptyne.com, paste this code into the Code Editor (right side of screen),
#       then type "createTable()" in the command line at the bottom and press enter.
#
#
#   Future Features (to add...):
#       Using excel macros to automatically generate BOMs and add part number to each sequence


import re
import neptyne as nt

BOM = nt.sheets["BOMs"]
SEQUENCE = nt.sheets["SEQUENCE"]

# Main function to create a table structure for a project
# It initializes a Project object, finds and adds items (Z) to it, and then adds sequences to each item
def createTable():
    global project
    project = Project(G3, B3)  # Initialize project with id and name
    project.addItems()  # Find and add items to the project
    for item in project.items:  # Add sequences to each item
        item.addSequences()
        #for seq in item.sequences:
            #seq.addParts()
    print("Completed. Running test...")
    #test()  # Run test function to print sequences
    print("Done!")
    #if(!(BOM)):
    #    nt.sheets.new_sheet("BOMs")
    BOMs!A2 = "= project.printTable()"
    BOM[1,5] = project.project_id
    return 0

def res():
    print("Resetting table...")
    BOM.cols[0:99].set_width(100)
    BOM.rows[0:40].set_height(20)
    for y in range(0,40):
        for x in range(1,99):
            BOM[x,y].clear()
    print("Done!")
    return 0

# Project class to represent a project with items
class Project:
    def __init__(self, project_id, project_name):
        self.project_id = project_id
        self.project_name = project_name
        self.items = []  # List to hold associated items

    def addItems(self):
        colData = str(E:E)  # Convert column data to string
        items = re.findall(r'Z-(?:[0-9][A-Z]){3}[0-9]', colData)  # Regex to find Z numbers
        n = 0
        for z_number in items:
            item = Item(n, z_number, self.project_id, "tempZName", "0")
            self.items.append(item)  # Append the item to the project
            n = n + 1
            
    def printTable(self):
        seqListList = []
        item_list = []
        checks = 5
        # Write out the data
        for item in project.items:
            seq_list = []
            for seq in item.sequences:
                seq_list.append(seq.sequence_id)
            seqListList.append(seq_list)
            item_list.append(item.item_id)
        # Format the table
        n = 0
        for item in item_list:
            col = n*checks
            BOM[2,col] = item
            BOM[3,col] = seqListList[n]
            BOM.cols[(col+1):(col+checks)].set_width(20)
            for seqList in seqListList[n]:
                for seq in seqList:
                    #seq.pos = [3 + seqList.index(seq), col]
                    BOM[3 + seqList.index(seq), col + 1] = '\u2022'
            n += 1
        return self.project_name

# Item class to represent an item within a project
class Item:
    def __init__(self, index, item_id, project_id, item_name, row):
        self.index = index
        self.item_id = item_id
        self.project_id = project_id
        self.item_name = item_name
        self.row = list(E:E).index(self.item_id) + 1  # Find row number of the item
        self.sequences = []  # List to hold associated sequences
    
    # Method to add sequences to the item
    def addSequences(self):
        #print("Adding sequences for:", self.item_id)
        # Determine start and end rows for searching sequences
        if(self.index == len(project.items) - 1):
            rowStart = self.row
            rowEnd = rowStart + 99
        else:
            next_item = project.items[self.index + 1]
            rowStart = self.row
            rowEnd = next_item.row

        # Check for N/A (no sequences) in the range
        string1 = str(E1:E[rowStart:rowEnd])
        if re.search(r'N/A', string1):
            return 0
        else:
            # Find and add sequences to this item
            string2 = str(B1:B[rowStart:rowEnd])
            seqs = re.findall(r'[0-9]{3}|REDB', string2)
            for id in seqs:
                sequence = Sequence(id, self.item_id)
                self.sequences.append(sequence)   

# Sequence class to represent a sequence within an item
class Sequence:
    def __init__(self, sequence_id, item_id):
        self.sequence_id = sequence_id
        self.item_id = item_id
        self.state = "404"
        self.pos = [] # Position of cell in table
        self.parts = []  # List to hold associated parts

    def addParts():
        return 0 # REPLACE ME

# Part class to represent a part within a sequence
class Part:
    def __init__(self, part_id, sequence_id, part_number):
        self.part_id = part_id
        self.sequence_id = sequence_id
        self.part_number = part_number

# Function to print sequences for each item in the project
def printSequences():
    for item in project.items:
        for s in item.sequences:
            print(s.sequence_number)

# Test function to print the number of sequences for each item
def test():
    for item in project.items:
        print(item.item_id, len(item.sequences))
