########################################################################################################################
#                          Python Script --> Embedded LLT python script                                                #
#                                       DF - CSW 2019                                                                  #
# Input parameters template: 1st - script.py 2nd - directory/file.xlsx 3rd - directory/file.docx or just directory     #
########################################################################################################################

import sys
import os.path

print("Before run this script please look for the initial dependencies written on the code file!")
# Check input parameters length
if len(sys.argv) != 3:
    print('Invalid number of arguments to input!')
    print('Use this template: python function.py <directory.file.xlsx> <directory.filedocx or just directory>')

    sys.exit(2)
else:
    print("Processing your path!")

# Check for excel file path
if os.path.exists(sys.argv[1]):
    path = os.path.abspath(sys.argv[1])  # if path is relative transform it on absolute
    print("found excel file path")
    print(path)
else:
    print("Cant find " + sys.argv[1])