import openpyxl


class donor:
    def __init__(self, first, last, address, city, state, zip):
        self.first = first
        self.last = last
        self.address = address
        self.city = city
        self.state = state
        self.zip = zip
        
def excelOps(loc):
    #open excel doc and specific sheet
    wb = openpyxl.load_workbook(filename = loc)
    ws = wb.get_sheet_by_name("AmplifyOriginal")

    #create dictionary of column headers names and the index of the column
    headers = getColumnKey(ws)

    firstColumn = headers['*First Name']
    lastColumn = headers['*Last Name']
    nameColumn = headers['PROfileName']
    row = 2
    name = ws.cell(row=row,column=nameColumn).value
    delim = ' '

    #Split the name column into separate first and last name columns
    while name:
        first = delim.join(name.split(' ')[:-1])
        ws.cell(row=row,column=firstColumn).value = first
        ws.cell(row=row,column=lastColumn).value = name.split(' ')[-1]
        row += 1
        name = ws.cell(row=row,column=nameColumn).value

    wb.save('AmplifyUpdated.xlsx')

def getColumnKey(ws):
    headers = {}
    row = 1
    column = 1
    header = ws.cell(row=row,column=column).value

    while header:
        headers[header] = column
        column += 1
        header = ws.cell(row=row,column=column).value
    
    return headers

def mergeInfo(loc1, loc2):
    #open excel file and sheets - one that data is being read from and one that data will be added to
    wb = openpyxl.load_workbook(filename = loc1)
    ws = wb.get_sheet_by_name("Sheet1")
    wb2 = openpyxl.load_workbook(filename = loc2)
    ws2 = wb2.get_sheet_by_name("AmplifyOriginal")

    #create dictionary of column headers names and the index of the column for both sheets
    key = getColumnKey(ws)
    key2 = getColumnKey(ws2)

    #read the information from sheet 2 into a dictionary of class donor
    donors = []
    firstColumn = key2['*First Name']
    row = 2
    name = ws2.cell(row=row,column=firstColumn).value

    while name:
        donors.append(donor(name, ws2.cell(row=row,column=key2['*Last Name']).value, ws2.cell(row=row,column=key2['Address1']).value, ws2.cell(row=row,column=key2['*City']).value, ws2.cell(row=row,column=key2['*State/Province']).value, ws2.cell(row=row,column=key2['ZIP/Postal Code']).value))
        row += 1
        name = ws2.cell(row=row,column=firstColumn).value

    firstColumn = key['First Name']
    lastColumn = key['Last Name']
    addressColumn = key['Street Address']
    cityColumn = key['City']
    stateColumn = key['State ']
    zipColumn = key['Zip']
    #phoneColumn = key['Phone']
    #emailColumn = key['Email address']

    row = 2
    name = ws.cell(row=row,column=firstColumn).value

    #add the information from each member of the donor dictionary into the excel sheet being written to
    while name:
        for d in donors:
            if d.first == name and d.last == ws.cell(row=row,column=lastColumn).value:
                if d.address and len(d.address.split()) > 2 and d.state == 'Minnesota':
                    ws.cell(row=row,column=firstColumn).value = d.first
                    ws.cell(row=row,column=lastColumn).value = d.last
                    ws.cell(row=row,column=addressColumn).value = d.address
                    ws.cell(row=row,column=cityColumn).value = d.city
                    ws.cell(row=row,column=stateColumn).value = d.state
                    ws.cell(row=row,column=zipColumn).value = d.zip
        row += 1
        name = ws.cell(row=row,column=firstColumn).value

    wb.save('PostcardUpdated.xlsx')


if __name__ == "__main__":
    loc = "../../Documents/AmplifyOriginal.xlsx"
    excelOps(loc)
    loc1 = "../../Documents/PostcardOriginal.xlsx"
    loc2 = "AmplifyUpdated.xlsx"
    mergeInfo(loc1, loc2)