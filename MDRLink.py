import os, sys
#for creating shortcut
import pythoncom
from win32com.shell import shell, shellcon

#for cleaining up the filename
import re
#for popup and graphics display
import PySimpleGUI as sg

#for dataframes
import pandas as pd

# Retrieve current working directory (`cwd`)
cwd = os.getcwd()


# Import `load_workbook` module from `openpyxl`
#from openpyxl import load_workbook

# Load in the workbook
#wb = load_workbook('MDR.xlsx')


#taking user input for path - not used anymore, using popup GUI
#user_path = input("Enter the path of documents folder: ")

layout = [[sg.Text('Select Excel Files')],
          [sg.Text('Select Latest MDR File with Tags', size=(40,1)), sg.InputText(), sg.FileBrowse()],
          [sg.Text('Select Macro Xlxm File', size=(40,1)), sg.InputText(), sg.FileBrowse()],
          [sg.Text('Select Documents Folder Location', size=(40,1)), sg.InputText(), sg.FolderBrowse()],
          [sg.Submit()]]
window = sg.Window('Select Files', layout)

event, values = window.read()
window.close()
QuickReportMDR = values[0]
ExcelMacroMDR  = values[1]
user_path = values[2]
print (QuickReportMDR, ExcelMacroMDR)



#opening Big MDR file with all the tags / folders 

df = pd.ExcelFile(QuickReportMDR)
#WorkSheet = pd.read_excel(df, 0, header=None)

#line below to be deleted?
#sheetlist = df.sheet_names

#docnumbers = WorkSheet[5].tolist()
#Folder1 = WorkSheet[1].tolist()
#Folder2 = WorkSheet[2].tolist()

#different way of doing above steps:
docnumbers = pd.read_excel(df, 0, usecols=['Doc Number'], squeeze = True).values.tolist()
Folder1 = pd.read_excel(df, 0, usecols=['Plant Class'], squeeze = True).values.tolist()
Folder2 = pd.read_excel(df, 0, usecols=['Tag'], squeeze = True).values.tolist()



print ("List of items in Tagged MDR list " + str(len(docnumbers)))
#writing to the file
with open('Log.txt', mode='w', encoding='utf-8') as a_file:
    a_file.write("List of items in Tagged MDR list " + str(len(docnumbers)) + '\n')
a_file.close()

#writing to the missing.txt file once so that old data is over written since the data below is only appended. 
with open('Missing.txt', mode='w', encoding='utf-8') as b_file:
            b_file.write('start' + '\n')
b_file.close()


#opening MDR file with Macros
df2 = pd.ExcelFile(ExcelMacroMDR)
WorkSheet2 = pd.read_excel(df2, 2, header=None)

#sheetlist2 = df2.sheet_names
docnumbers2 = WorkSheet2[3].tolist()
doclinks2 = WorkSheet2[23].tolist()
doctitle2 = WorkSheet2[8].tolist()



print ("List of items in Macro MDR list " + str(len(docnumbers2)))
print ("List of items in Macro MDR title " + str(len(doctitle2)))
#writing to the file
with open('Log.txt', mode='a', encoding='utf-8') as a_file:
    a_file.write("List of items in Macro MDR list " + str(len(docnumbers2)) + '\n')
    a_file.write("List of items in Macro MDR title " + str(len(doctitle2)) + '\n')
a_file.close()


shortcut = pythoncom.CoCreateInstance (
  shell.CLSID_ShellLink,
  None,
  pythoncom.CLSCTX_INPROC_SERVER,
  shell.IID_IShellLink
)



#for i in range (10, len(docnumbers)):
for i in range (10, len(docnumbers)):

                                                        
    
    
    rootpath = os.getcwd()
    #cleaning up the folder text so that there is no illegal characters for directory structure. 
    Folder1[i] = Folder1[i].replace("/", "_")
    Folder2[i] = Folder2[i].replace("/", "_")
    re.sub('[^\w\-_\. ]', '_', Folder1[i])
    re.sub('[^\w\-_\. ]', '_', Folder2[i])
    
    #Creating first folder
    path = os.path.join(rootpath, Folder1[i])
    
    if os.path.isdir(path) == False:
    
        try:
            os.mkdir(path)
        except OSError as error:
            print(error)
    #creating second folder        
    path = os.path.join(rootpath, Folder1[i], Folder2[i])
        
    if os.path.isdir(path) == False:
    
        try:
            os.mkdir(path)
        except OSError as error:
            print(error)
    
    #try to find the document number in the MDR file where are entries are with file link. 
    try:
        docIndextemp = docnumbers2.index(str(docnumbers[i]))
        docIndex = docIndextemp[:7]
    except:
        #if not found, write it to the missing dot txt file. 
        with open('Missing.txt', mode='a', encoding='utf-8') as b_file:
            b_file.write(str(docnumbers[i]) + " - Document number doesn't exist in MDR Lookup"  + '\n')
        b_file.close()

        
    else:
        doclink = str(doclinks2[docIndex])
        
        #to get document title from the link file
        #find the position of last slash and the last dot. the file name is between these two characters
        docslashpositon = int(doclink.rfind("\\"))
        docdotposition = int(doclink.rfind('.'))
        cleandoclink = doclink[docslashpositon+1:docdotposition]
        
        #prepare file path as per user_path and file name
        #for some reason os path join is not working correctly
        #need to replace dwg extension with pdf extension
       
        if doclink[docdotposition:] == '.DWG':
            shortcut_lnk_file_name = str(doclink[docslashpositon+1:docdotposition]) + '.PDF'
            print (shortcut_lnk_file_name)
        
        else:
            shortcut_lnk_file_name = doclink[docslashpositon:]
        shortcut_lnk_full = str(user_path) +  shortcut_lnk_file_name
       
        
        #old doctitle was claculated this way
        #doctitle = doctitle2[docIndex].replace("/", "_")
        #re.sub('[^\w\-_\. ]', '_', doctitle)
        #doctitle = doctitle.replace("\n", "_")
        #doctitle = doctitle.replace("\t", "_")
        
        #after update we just assign cleandoclink to doctitle
        doctitle = cleandoclink

        #print ("Link to document number " + str(docnumbers[i])+  " is " + str (doclink) + " title " + str (doctitle))
        #writing to the file
        #with open('Log.txt', mode='a', encoding='utf-8') as a_file:
        #    a_file.write("Link to document number " + str(docnumbers[i])+  " is " + str (doclink) + " title " + str (doctitle) + '\n')
        #a_file.close()
        
        sg.one_line_progress_meter('Creating Shortcuts', i, len(docnumbers), 'key', 'Now Creating '+ Folder1[i] + ' \\ ' + Folder2[i] + ' \\ ' + cleandoclink, orientation='h', size=(1000,400) )
        
        #if doclink is not empty
        if len(str(doclink)) > 2:
            program_location = str(shortcut_lnk_full)
            shortcut.SetPath (program_location)
            shortcut.SetDescription (cleandoclink)
            shortcut.SetIconLocation (program_location, 0)

            #desktop_path = shell.SHGetFolderPath (0, shellcon.CSIDL_DESKTOP, 0, 0)
            desktop_path = path
            file_name = str(cleandoclink) + "_shortcut.lnk"
            persist_file = shortcut.QueryInterface (pythoncom.IID_IPersistFile)
            #enable below line for showing all the paths in debug window
            #print (str(os.path.join (desktop_path, file_name)))
            try:
                persist_file.Save (os.path.join (desktop_path, file_name), 0)
            except Exception as e:
                print ("error" + str (doclink) + str (doctitle) + str (docnumbers[i]))
                #writing to the file
                with open('Log.txt', mode='a', encoding='utf-8') as a_file:
                    a_file.write("error code: " + str(e) + "Unable to create shortcut for - " + str (cleandoclink) + " Doc Number - " + str (docnumbers[i]) + " Doc Link - " + str(os.path.join (desktop_path, file_name)) + '\n')
                a_file.close()
        
        
