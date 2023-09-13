import xlwings as xw
import l5x
import shutil
import os
import time

print('#############################################################################################')
print('######    This script will update Tag & Description in L5X file based on Excel file    ######')
print('###### Created by KTolloczko but highly inspired by https://github.com/jvalenzuela/l5x ######')
print('######        Version 1.0 early beta - 2021-09-15. Please use on your own risk.        ######')
print('#############################################################################################')
print('!!! ATTENTION !!!')
print('-> Manually modify lines 19-20 and 31 in the script in order to select Excel file, Excel sheet and L5X file.')
print('-> Make sure that all files do not contain any spaces in their names, sheets, and that they are in the same folder!')
print('-> The script will save L5X file as "_output.L5X" and create backup file "_output_backup.L5X" in the same folder.')
print('-> In Logix Designer click import Routine and select all tags to be overwritten.')
print('-> Check for errors and enjoy your time saved!:)')

# Opening EXCEL File once
print('Opening Excel file...', end=' ')
exepath = r"C:\Users\ktolloczko\Desktop\PLC_31X_IO_Allocation_by_Rack.xlsx"
exesheet = 'Rack_10'
try:
    # Change manually file name / please exclude any spaces!!!
    exelf = xw.Book(exepath).sheets[exesheet]
except:
    print('Cannot open Excel file!')
else:
    print('Done.')

# Opening L5X File once
print('Opening L5X file...', end=' ')
l5xpath = r"C:\Users\ktolloczko\Desktop\_011_Input_Output_Mapping_Rack10_modified.L5X"
try:
    # Change manually file name / please exclude any spaces!!!
    xmlf = l5x.Project(l5xpath)
except:
    print('Cannot open L5X File!')
else:
    print('Done.')

# Get the current working directory
cwd = os.getcwd()

# Saving L5X File output file with backup
print('Saving L5X File output files in "',cwd,'"...', end=' ')
try:
    shutil.copyfile(l5xpath, r'_output.L5X')
    shutil.copyfile(r'_output.L5X', r'_output_backup.L5X')
except:
    print('Cannot save L5X File!')
else:
    print('Done.')

# Input Excel reading ranges
print('Please enter Start range of Tags in Excel file (e.g. 2, 10):')
starow = int(input())
print('Please enter End range of Tags in Excel file (e.g. 2, 10):')
endrow = int(input())

# Create Tag list of names in L5X
L5Xtaglist = xmlf.controller.tags.names
print('Raw Tag list from L5X:')
print(L5Xtaglist)

# Remove prefix from Tags list in L5X #TODO: add loop for removing prefixes
nopreL5Xtaglist20 = [str(s).removeprefix('PSC') for s in L5Xtaglist]
nopreL5Xtaglist19 = [str(s).removeprefix('LSO') for s in nopreL5Xtaglist20]
nopreL5Xtaglist18 = [str(s).removeprefix('PRO') for s in nopreL5Xtaglist19]
nopreL5Xtaglist17 = [str(s).removeprefix('LVO') for s in nopreL5Xtaglist18]
nopreL5Xtaglist16 = [str(s).removeprefix('RLY') for s in nopreL5Xtaglist17]
nopreL5Xtaglist15 = [str(s).removeprefix('TE') for s in nopreL5Xtaglist16]
nopreL5Xtaglist14 = [str(s).removeprefix('LDC') for s in nopreL5Xtaglist15]
nopreL5Xtaglist13 = [str(s).removeprefix('LDT') for s in nopreL5Xtaglist14]
nopreL5Xtaglist12 = [str(s).removeprefix('LVC') for s in nopreL5Xtaglist13]
nopreL5Xtaglist11 = [str(s).removeprefix('SSO') for s in nopreL5Xtaglist12]
nopreL5Xtaglist10 = [str(s).removeprefix('SOL') for s in nopreL5Xtaglist11]
nopreL5Xtaglist9 = [str(s).removeprefix('IND') for s in nopreL5Xtaglist10]
nopreL5Xtaglist8 = [str(s).removeprefix('LSC') for s in nopreL5Xtaglist9]
nopreL5Xtaglist7 = [str(s).removeprefix('ZSO') for s in nopreL5Xtaglist8]
nopreL5Xtaglist6 = [str(s).removeprefix('DSO') for s in nopreL5Xtaglist7]
nopreL5Xtaglist5 = [str(s).removeprefix('PBO') for s in nopreL5Xtaglist6]
nopreL5Xtaglist4 = [str(s).removeprefix('CRO') for s in nopreL5Xtaglist5]
nopreL5Xtaglist3 = [str(s).removeprefix('BKR') for s in nopreL5Xtaglist4]
nopreL5Xtaglist2 = [str(s).removeprefix('PPC') for s in nopreL5Xtaglist3]
nopreL5Xtaglist = [str(s).removeprefix('TSC') for s in nopreL5Xtaglist2]
print('Tag list without prefix from L5X:')
print(nopreL5Xtaglist)

# Selecting data from range in Excel file /+1 because of range function
for i in range(starow,endrow+1):

    print('Searching Tags in Excel File...', end=' ')

    # Get Tag and Description from Excel
    try:
        Exetag = exelf.range(f'B{i}').value
        Exedes = exelf.range(f'C{i}').value
        print('Found.')

        print('Excel Tag found:', Exetag, end=' ')
        print('- Description:', Exedes)
    except:
        print('Cannot find any Tags in Excel File!')
        break
    
    print('Searching Tags in L5X File...', end=' ')

    # Remove prefix from Excel Tag #TODO: add loop for removing prefixes
    nopreExetag20 = str(Exetag).removeprefix('PSC')
    nopreExetag19 = str(nopreExetag20).removeprefix('LSO')
    nopreExetag18 = str(nopreExetag19).removeprefix('PRO')
    nopreExetag17 = str(nopreExetag18).removeprefix('LVO')
    nopreExetag16 = str(nopreExetag17).removeprefix('RLY')
    nopreExetag15 = str(nopreExetag16).removeprefix('TE')
    nopreExetag14 = str(nopreExetag15).removeprefix('LDC')
    nopreExetag13 = str(nopreExetag14).removeprefix('LDT')
    nopreExetag12 = str(nopreExetag13).removeprefix('LVC')
    nopreExetag11 = str(nopreExetag12).removeprefix('SSO')
    nopreExetag10 = str(nopreExetag11).removeprefix('SOL')
    nopreExetag9 = str(nopreExetag10).removeprefix('IND')
    nopreExetag8 = str(nopreExetag9).removeprefix('LSC')
    nopreExetag7 = str(nopreExetag8).removeprefix('ZSO')
    nopreExetag6 = str(nopreExetag7).removeprefix('DSO')
    nopreExetag5 = str(nopreExetag6).removeprefix('PBO')
    nopreExetag4 = str(nopreExetag5).removeprefix('CRO')
    nopreExetag3 = str(nopreExetag4).removeprefix('BKR')
    nopreExetag2 = str(nopreExetag3).removeprefix('PPC')
    nopreExetag = str(nopreExetag2).removeprefix('TSC')
    
    # Searching in modified Tag list with no prefix
    if nopreExetag in nopreL5Xtaglist:

        # Check if tag exists in L5X
        try:
            # Get tag name from L5X list
            L5Xtag = L5Xtaglist[nopreL5Xtaglist.index(nopreExetag)]

            # Get tag and description from L5X
            L5Xdes = xmlf.controller.tags[L5Xtag].description
            print('Found.')
            
            print('L5X Tag found:',  L5Xtag, end=' ')
            print('- Description:', L5Xdes)
            
            print('Checking if description is matching...', end=' ')

            if L5Xdes != Exedes:
                print('not matching!')   

                print('Updating Tag Description to:', Exedes, end=' ')

                # Update description in L5X
                xmlf.controller.tags[L5Xtag].description = exelf.range(f'C{i}').value 
                print('Done.')

                # Save L5X
                print('Saving L5X file... in "',cwd,'"...', end=' ')
                xmlf.write(r'_output.L5X') 
                print('Done.')
            else:
                print('yes.')
        except:
            print('Error while searching for Tag in L5X!')
            break
    print('Completed.')

print('Veryfing if Tag Names are matching...')

# Selecting data from range in Excel file /+1 because of range function
for i in range(starow,endrow+1):

    print('Searching Tags in Excel File...', end=' ')

    # Get Tag and Description from Excel
    try:
        Exetag = exelf.range(f'B{i}').value
        Exedes = exelf.range(f'C{i}').value
        print('Found.')

        print('Excel Tag found:', Exetag, end=' ')
        print('- Description:', Exedes)
    except:
        print('Cannot find any Tags in Excel File!')
        break
    
    print('Searching Tags in L5X File...', end=' ')
   
    # Remove prefix from Excel Tag #TODO: add loop for removing prefixes
    nopreExetag20 = str(Exetag).removeprefix('PSC')
    nopreExetag19 = str(nopreExetag20).removeprefix('LSO')
    nopreExetag18 = str(nopreExetag19).removeprefix('PRO')
    nopreExetag17 = str(nopreExetag18).removeprefix('LVO')
    nopreExetag16 = str(nopreExetag17).removeprefix('RLY')
    nopreExetag15 = str(nopreExetag16).removeprefix('TE')
    nopreExetag14 = str(nopreExetag15).removeprefix('LDC')
    nopreExetag13 = str(nopreExetag14).removeprefix('LDT')
    nopreExetag12 = str(nopreExetag13).removeprefix('LVC')
    nopreExetag11 = str(nopreExetag12).removeprefix('SSO')
    nopreExetag10 = str(nopreExetag11).removeprefix('SOL')
    nopreExetag9 = str(nopreExetag10).removeprefix('IND')
    nopreExetag8 = str(nopreExetag9).removeprefix('LSC')
    nopreExetag7 = str(nopreExetag8).removeprefix('ZSO')
    nopreExetag6 = str(nopreExetag7).removeprefix('DSO')
    nopreExetag5 = str(nopreExetag6).removeprefix('PBO')
    nopreExetag4 = str(nopreExetag5).removeprefix('CRO')
    nopreExetag3 = str(nopreExetag4).removeprefix('BKR')
    nopreExetag2 = str(nopreExetag3).removeprefix('PPC')
    nopreExetag = str(nopreExetag2).removeprefix('TSC')

    # Searching in modified Tag list with no prefix
    if nopreExetag in nopreL5Xtaglist:

        # Check if tag exists in L5X
        try:
            # Get tag name from L5X list
            L5Xtag = L5Xtaglist[nopreL5Xtaglist.index(nopreExetag)]

            # Get tag and description from L5X
            L5Xdes = xmlf.controller.tags[L5Xtag].description
            print('Found.')
            
            print('L5X Tag found:',  L5Xtag, end=' ')
            print('- Description:', L5Xdes)
            
            print('Checking if tag name is matching...', end=' ')

            if L5Xtag != Exetag:
                print('not matching!')

                # ====================================================================================================
                # Open L5X as txt file *because of bug in l5x library for not being able to write to tag name*
                with open(r'_output.L5X', 'r') as txtfr: 
                    txtfdata = txtfr.read()

                print('Updating Tag Name to:', Exetag, end=' ')
                # Replace the target string *bug in l5x library for not being able to write to tag name*/see above
                txtfdata = txtfdata.replace(L5Xtag, Exetag)
                print('Done.')
                
                # Save L5X
                print('Saving L5X file... in "',cwd,'"...', end=' ')
                # Write the file out again *bug in l5x library for not being able to write to tag name*/see above
                with open(r'_output.L5X', 'w') as txtfw:  
                    txtfw.write(txtfdata)
                print('Done.')
                # ====================================================================================================
            else:
                print('yes.')
        except:
            print('Error while searching for Tag in L5X!')
            break
    print('Completed.')