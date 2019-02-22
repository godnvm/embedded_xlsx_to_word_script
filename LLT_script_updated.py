########################################################################################################################
#                          Python Script --> Embedded LLT python script                                                #
#                                       DF - CSW 2019                                                                  #
# Input parameters template: 1st - script.py 2nd - directory/file.xlsx 3rd - directory/file.docx or just directory     #
########################################################################################################################

import sys
import os.path
from docx import Document
import win32com.client as win32
import docx

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


if os.path.exists(sys.argv[2]):
    path1 = os.path.abspath(sys.argv[2])  # if path is relative transform it on absolute
    print("found docx/directory docx file path")
    print(path1)
else:
    print("Cannot find " + sys.argv[2])
    exit()


# This function allows us to trim, select and use text beetween 2 input words
def between(value, a, b):
    # Find and validate before-part.
    pos_a = value.find(a)
    if pos_a == -1:
        return ""
    # Find and validate after part.
    pos_b = value.rfind(b)
    if pos_b == -1:
        return ""
    # Return middle part.
    adjusted_pos_a = pos_a + len(a)
    if adjusted_pos_a >= pos_b:
        return ""
    value2 = value[adjusted_pos_a:pos_b]
    return value2


doc = docx.Document(path1)
word = win32.gencache.EnsureDispatch('Word.Application')
for paragraph in doc.paragraphs:
        doc = word.Documents.Open(path1)
        word.Visible = False

        str1 = path
        a, b = str1.split(".", 1)
        str2 = a
        c, d = str2.rsplit("\\", 1)
        e = d + ('.' + b)

        str3 = e
        print(str3)
        for rge in doc.Words:
            rge.Find.Text = '<EMBEDDED_TSP_SPREADSHEET>'
            found = rge.Find.Execute()
            if (found == True):
                print (found)
                doc.InlineShapes.AddOLEObject(ClassType="Excel.Sheet", FileName=path, DisplayAsIcon=True, IconLabel=str3, IconFileName=path)

word.ActiveDocument.SaveAs(path1)
doc.Close()








