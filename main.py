import PySimpleGUI as sg
from datetime import datetime
from generate_xlsx import generate_xlsx_file
import os

#theme
sg.theme_add_new(
    'Synthwave', 
    {
        'BACKGROUND': '#2e4342', 
        'TEXT': 'white', 
        'INPUT': '#B5E5CF', 
        'TEXT_INPUT': '#2e4342', 
        'SCROLL': '#B99095', 
        'BUTTON': ('#2e4342', '#FCB5AC'), 
        'PROGRESS': ('black', '#FCB5AC'), 
        'BORDER': 1, 
        'SLIDER_DEPTH': 0, 
        'PROGRESS_DEPTH': 1, 
        'COLOR_LIST': ['#160f30', '#241663', '#a72693', '#eae7af'], 
    }
)   
sg.theme('Synthwave')
#folder name start
folderName = 'tmt'
# Define the layout of the GUI
layout = [
    [sg.Text("Project name:"), sg.InputText(key="projectName")],
    [sg.Text("Project number:"), sg.InputText(key="num")],
    [sg.Text("Date:"), sg.InputText(key="date"), sg.CalendarButton("Select Date", target="date", format='%Y-%m-%d')],
    [sg.Text("Technology:"), sg.Combo(["MILLING", "TURNING", "LASER", "3D PRINT - PA", "3D PRINT - DMLS", "TURNING-MILLING", "ABRASIVE WATER JET", "W-EDM", "DRILLING-CHAMFERING", "TUBE LASER"], key="technology")],
    [sg.Text("Contractor:"), sg.Combo(["TIMTEC", "ALMI", "Laser HRIBAR", "WALTECH", "PETKOVŠEK", "WEERG", "JEŽ", "MISLEJ", "ŠINKOVEC", "BOGADI", "LA&CO"], key="contractor")],
    [sg.Text("Additional:"), sg.InputText(key="additional")],
    [sg.Text("Folder location:"), sg.InputText(key="location"), sg.FolderBrowse()],
    [sg.Text("Folder name:"), sg.Text(folderName, key="folder_name_display", background_color="white", text_color="black")],  # Added element to display folderName
    [sg.Button("Update folder name", key='updateFolderName'), sg.Button("Generate the folder and XSLX", key = "Create")],
    ]


# Create the window
window = sg.Window("TIMTEC folder automation - Martinov digitalni asistent", layout,  icon='timtecLogo.ico', resizable=False, font=("Helvetica", 12, "bold"))

# Open the folder function
def openFolder(folderPath):
    os.startfile(folderPath)

#Handle the "folder creation successful" popup
def folderCreationSuccess(folderToOpen):
    layout = [
        [sg.Text("Folder Created!")],
        [sg.Button("Open Folder", key="open_button")]
    ]
    window = sg.Window("Folder creation successful" , layout, icon='timtecLogo.ico', resizable=True, size=(350,70))

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == "open_button":
            openFolder(folderToOpen)
            window.close()
            break
    


#Handle the "folder already exists" popup and creates the folder with the name "part X" extension
def folderExtension(folderName, folderPath):
    num = 1
    while True:
        choice = sg.popup_yes_no(f"Folder '{folderName}' already exists! Do you want to add extension to the name?", title='Folder Exists', icon='timtecLogo.ico')
        if choice == 'No':
            break
        else:
            newFolderName = f"{folderName} - part {num}"
            newFolderPath = os.path.join(folderPath, newFolderName)
            try:
                os.makedirs(newFolderPath)
                generate_xlsx_file(newFolderName, newFolderPath, values["projectName"], values["num"], values["date"])
                folderCreationSuccess(newFolderPath)
                break
            except FileExistsError:
                num += 1

                
#Create the folder
def createFolder():
    if values['projectName'] != '' and values['num'] != '' and values['technology'] != '' and values['location'] != '':
        if values['date'] == '':
            values['date'] = datetime.today().strftime('%Y-%m-%d')
        if values['additional'] == '':
            folderName = f"tmt {values['projectName'].upper()} {values['num']} - {values['date']} -rev-1 - {values['technology']} - {values['contractor']}"
        else:
            folderName = f"tmt {values['projectName'].upper()} {values['num']} - {values['date']} -rev-1 - {values['technology']} - {values['contractor']} - {values['additional']}"           
        location = values["location"]
        folder_path = os.path.join(location, folderName)
        try:
            os.makedirs(folder_path)
            generate_xlsx_file(folderName, folder_path, values["projectName"], values["num"], values["date"])  
            folderCreationSuccess(folder_path)
        except FileExistsError:
            folderExtension(folderName, location)
    else:
        sg.popup("Folder creation failed! Please check that all the required fields are filled.", title='User is always wrong, the program is perfect', icon='timtecLogo.ico')
    
# Event loop
while True:
    event, values = window.read()

    # Exit if the window is closed
    if event == sg.WINDOW_CLOSED:
        break

   # Update folderName and display it
    if event == "updateFolderName":
        if values['date'] == '':
            values['date'] = datetime.today().strftime('%Y-%m-%d')
        folderName = f"tmt {values['projectName'].upper()} {values['num']} - {values['date']} -rev-1 - {values['technology']} - {values['contractor']} {values['additional']}"
        window["folder_name_display"].update(folderName)
        window.refresh()

    # Create the folder if the "Create" button is clicked
    if event == "Create":
        createFolder()

# Close the window
window.close()

#pip install pyinstaller
#pyinstaller main.py --onefile --name "Timtec folder automation" --icon timtecLogo.ico --noconsole

#pyinstaller --add-data "C:\Users\User\Desktop\TimtecAutoFolder/template.xlsx:." --workpath "./dist" main.py --onefile --name "Timtec folder automation" --icon timtecLogo.ico --noconsole --windowed