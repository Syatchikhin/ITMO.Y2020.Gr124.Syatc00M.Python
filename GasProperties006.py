#-------------------------------------------------------------------------------
# Name:        GasProperties001
# Purpose:     Gas properties calculation
#
# Author:      Maksim Syatchikhin
#
# Created:     16.07.2020
#-------------------------------------------------------------------------------

def main():
    path1 = "C:\\Temp\\source.xlsx"

    class Gas:
        Name = ""
        Size = 0
        Path = ""
        Dencity = 0 
        MixtureR = 0 
        componentName = list("")
        componentFormula = list("")
        componentData = list()
        componentWeight = list()
    
    myGas1 = Gas()
    myGas1 = ReadExcelFile(path1, myGas1)
    myGas1NormalizedComposition = NormalizedComposition( myGas1)
    myGas1Calculated = GasCalculation(myGas1NormalizedComposition)
    Output(myGas1Calculated)

def ReadExcelFile(path, Gas):
    import win32com.client
    Excel = win32com.client.Dispatch("Excel.Application")

    wb = Excel.Workbooks.Open(path)
    sheet = wb.ActiveSheet
    Gas.Name = sheet.Cells(2,1).value
    #print(Gas.Name)
    #read all the cells of active sheet as instance
    readData = wb.Worksheets('page1')
    allData = readData.UsedRange
    # Get number of rows used on active sheet
    max_row = allData.Rows.Count
    start_row = 5
    Gas.Size = max_row - start_row + 1
    Gas.Path = path
    
    for i in range (0, Gas.Size):
        Gas.componentName.append(sheet.Cells(i+start_row,2).value)
        Gas.componentFormula.append(sheet.Cells(i+start_row,3).value)
        Gas.componentData.append(sheet.Cells(i+start_row,4).value)
        Gas.componentWeight.append(sheet.Cells(i+start_row,5).value)
   # print (Gas)

    return Gas

def NormalizedComposition(Gasx):
    actualMass = 0
    temp = 0
    for i in range (0, Gasx.Size):
        #Фактическая масса
        actualMass += Gasx.componentWeight[i]
    for i in range (0, Gasx.Size):
        temp =  Gasx.componentWeight[i]
        Gasx.componentWeight[i] = (temp * 100) / actualMass

    return Gasx

def GasCalculation(Gasz):
    GasConstant = 8314.462 #8314.462618
    GasMoleVolume = 22.414 #22.41396954
    totalMiRi = 0.0
    TotalWeight = 0.0
    componentVolume = 0.0
    componentMi = 0.0
    componentRi = 0.0
    componentMiRi = 0.0
    for i in range (0, Gasz.Size):
        componentVolume = Gasz.componentWeight[i]*10 #Volume in litres
        componentMoleWeight = componentVolume / GasMoleVolume # Moles amount
        Gasz.componentWeight[i] = componentMoleWeight * Gasz.componentData[i]
        TotalWeight  += Gasz.componentWeight[i]  #Full mass
        Gasz.Dencity = TotalWeight / 1000
    for i in range (0, Gasz.Size):
        componentMi = Gasz.componentWeight[i] / TotalWeight * 100 # Доля компонента от всей массы
        componentRi = GasConstant / Gasz.componentData[i] #Ri
        componentMiRi = componentMi * componentRi #RiMi
        totalMiRi  += componentMiRi
        Gasz.MixtureR = totalMiRi / 100 #R mixture

    return Gasz

def Output(Gasv):
    print("Плотность смеси, RO={0} кг/м3".format(round(Gasv.Dencity,3)))
    print("Газовая постоянная смеси, R={0} Дж/(кг*К)".format(round(Gasv.MixtureR,3)))
    return

# Вызов функции main
main()
