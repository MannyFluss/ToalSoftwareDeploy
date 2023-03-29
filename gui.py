import asset

import PySimpleGUI as sg

def main():
    layout = [
        [sg.Text("Manny's Toal Time Trimmer")],
        [sg.Text("Input CSV File:")],
        [sg.Input(key="csvFile"), sg.FileBrowse()],
        [sg.Text("Input Excel File:")],
        [sg.Input(key="excelFile"), sg.FileBrowse()],
        [sg.Text("Input starting line")],
        [sg.Input(key="input_num", size=(10,1))],
        [sg.Text("Input ending Line")],
        [sg.Input(key="input_num2", size=(10,1))],
        [sg.CalendarButton("Select Date", target="-IN-", key="_CALENDAR_", format="%Y/%m/%d")],
        [sg.Input(key="-IN-")],
        [sg.Text("Input Target Sheet (usually MONCTRL)")],
        [sg.Input(key="input_sheet", size=(10,1))],
        [sg.Button('Proccess File')],

        [sg.Text('Debug Console:')],
        [sg.Text(key="-DEBUGOUTPUT-")],
        ]

    window = sg.Window('My window', layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == "_CALENDAR_":
            window["_IN_"].update(values["_CALENDAR_"])
        elif event == 'Proccess File':
            returnStatus = proccessFile(values["csvFile"],values["excelFile"],values["input_num"],values["-IN-"],values["input_num2"],values["input_sheet"])
            window["-DEBUGOUTPUT-"].update(returnStatus)

    window.close()

def proccessFile(csvFilePath,excelFilePath,startingLine,selectedDate,EndingLine,SheetName)->str:
    if (csvFilePath == None or excelFilePath == None or startingLine == None or selectedDate == None or EndingLine==None or SheetName==None):
        return "Error: input required in all fields"
    if (csvFilePath[-4:] != ".csv"):
        return "Error: Invalid input (.csv is not detected)"
    if (excelFilePath[-5:]!= ".xlsx"):
        return "Error: Invalid input (.xlsx is not detected)"
    try:
        asset.proccessExcelSheet(csvFilePath,excelFilePath,selectedDate,startingLine,EndingLine,SheetName)
    except:
        return "error occurred during excel sheet reading, check that file paths are correct"

    
    return "Job Done (No Errors)"

if __name__ == '__main__':
    main()
