# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

from openpyxl import load_workbook

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.

def getCoordinateSameRow(coordinate, column):
    coordinate = list(coordinate)
    coordinate[0] = column
    return ''.join(coordinate)

def getCoordinateSameColumn(coordinate):
    coordinate = list(coordinate)
    result = coordinate[1:]# ['2','3']
    result = ''.join(result)#"23"
    rowNumber = int(result)# 23
    rowNumber -= 1# 22
    coordinate = ''.join(coordinate[:1]) + str(rowNumber)
    return ''.join(coordinate)

def checkFile():
    filename = 'import.xlsx'
    wb = load_workbook(filename)
    return wb;

def parseFile(data):
    mainDictionnary = {
        "Catégorie article: ": "categorie_article",
        "Type article RR": "type_article_rr",
        "Type article Leblanc": "type_article_leblanc",
        "Ligne de produit": "ligne_de_produit",
        "Article BIO?": "article_bio",
        "Remplace le code article:": "remplace_le_code_article",
        "Code article :": "code_article",
        "Designation article francais:": "designation_article_francais",
        "Designation article Anglais:": "designation_article_anglais",
        "Export caisse RR:": "export_caisse_rr",
        "Designation caisse:": "designation_caisse",
        "Classe article:": "classe_article",
        "Familles statistiques:": "familles_statistiques"
    }
    # TODO remplir le mainDictionnary
    result = {}
    for v in data['B':'C']:
        for cellObj in v:
            #print(cellObj.coordinate, cellObj.value)
            try:
                if mainDictionnary[cellObj.value] == "familles_statistiques":
                    result[mainDictionnary[cellObj.value]] = []
                    newCoordinate = getCoordinateSameColumn(cellObj.coordinate)
                    for column in ['D', 'E', 'F', 'G', 'H', 'I', 'J']:
                        previousRow = getCoordinateSameRow(newCoordinate, column)
                        c = getCoordinateSameRow(cellObj.coordinate, column)
                        #print(data[previousRow].value.replace(" ","").lower(), data[c].value)
                        if data[c].value is not None:
                            result[mainDictionnary[cellObj.value]].append({
                                data[previousRow].value.replace(" ", "").lower(): data[c].value
                            })
                else:
                    newCoordinate = getCoordinateSameRow(cellObj.coordinate, 'C')
                    #print(newCoordinate, data[newCoordinate].value)
                    if data[newCoordinate].value is not None:
                        result[mainDictionnary[cellObj.value]] = data[newCoordinate].value
            except:
                pass
    print(result)

# {'familles_statistiques': [
#     { 'stat1': 'CALISSON'},
#     { 'stat2': 'CALISSON'},
#     { 'stat3': 'CALISSON'},
#     { 'stat4': 'CALISSON'},
# ]}


if __name__ == '__main__':
    # Give the location of the file
    data = checkFile();
    ws = data.active
    parseFile(ws)
    #formatData()


    # for row in ws.columns:
    #     for cell in row:
    #         print(cell.value)






    #for row in ws.values:
        #for value in row:
            #print(value)
    #wb.close()
    #for sheet in wb.worksheets:
    #    print(sheet.)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
