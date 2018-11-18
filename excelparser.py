import argparse
import openpyxl

#Argument parsing
parser = argparse.ArgumentParser()
parser.add_argument("-if", "--inputfile", dest="inputfile")
parser.add_argument("-of", "--outputfile", dest="outputfile")
args = parser.parse_args()

#Setting up worksheet (Excel)
wb = openpyxl.load_workbook(args.inputfile)
sheet = wb.active
#Setting up some variables
exlist = []
column = 'a'

#Reading 7 headers, and asks for input
for i in sheet['1']:
    print('Fill in the %s ' % (sheet[column+'1'].value))
    column = chr(ord(column.lower()) + 1).upper()
    exlist.append(input())

#Appending our input to the worksheet
sheet.append(exlist)
#Saving the worksheet
wb.save(args.outputfile)

