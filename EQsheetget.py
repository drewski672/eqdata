import openpyxl, pprint, os
print('Opening Workbook...')
os.chdir('C:\\Users\\Assembly1\\Desktop')
wb = openpyxl.load_workbook('EQsheet2.xlsx')
sheet = wb['EQsheet2']
eqData = {}

print('Reading Rows...')
for row in range(2, sheet.max_row + 1):
    flowRate = sheet['A' + str(row)].value
    units = sheet['B' + str(row)].value
    pressureInt = sheet['C' + str(row)].value
    bodyType = sheet['D' + str(row)].value
    diaphragmMaterial = sheet['E' + str(row)].value
    diaphragmThickness = sheet['F' + str(row)].value

    eqData.setdefault(bodyType, {})
    eqData[bodyType].setdefault(bodyType, {'Diaphragm Material': diaphragmMaterial, 'Diaphragm Thickness': diaphragmThickness, ' Pressure': pressureInt, ' Flow Rate': flowRate, ' Flow Units': units})
    

print('Writing Results...')
resultFile = open('EQdata.py', 'w')
resultFile.write('allData = ' + pprint.pformat(eqData))
resultFile.close()
print('Done.')


    
    

