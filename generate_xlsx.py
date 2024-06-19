from openpyxl import workbook, load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import os



def generate_xlsx_file(newFolderName, newFolderPath, projectName, projectNum, date):
    # Load the workbook
    wb = load_workbook(filename='template.xlsx')
    img = Image('tmtblack.png')
    img.height = 60
    img.width = 240
    # Select the active sheet
    sheet = wb.active

    # Write the values
    dateObj = datetime.strptime(date, "%Y-%m-%d")
    sheet['A2'] = projectName.upper()
    sheet['G2'] = f'{projectNum}-XXX-XX'
    sheet['G3'] = f'Datum revizije: {dateObj.strftime("%d.%m.%Y")}'
    sheet['K2'] = dateObj.strftime("%d/%m/%Y")
    sheet['D3'] = 'rev-1'
    sheet.add_image(img, 'L1')
    
    # Save the workbook with the new name
    file_path = os.path.join(newFolderPath, newFolderName + '.xlsx')
    if os.path.exists(newFolderPath):
        wb.save(file_path)