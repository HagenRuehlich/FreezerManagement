Attribute VB_Name = "FreezerMacros"
Const csBasicFormula As String = "='§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SourcSheetName'!$§§§SourcCellString1&ZEICHEN(10)&'§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SourcSheetName'!$§§§SourcCellString2&ZEICHEN(10)&'§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SourcSheetName'!$§§§SourcCellString3&ZEICHEN(10)&'§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SourcSheetName'!$§§§SourcCellString4"

Sub fillFreezerDataAutomatedSelection()
    'user the selected a single source cell, software creates automaticall a 4 cell row
    fillFreezerData (True)
End Sub
Sub fillFreezerDataManualSelection()
    'user has to select max 4 cells by hand
    fillFreezerData (False)
End Sub




Sub fillFreezerData(bAutomatedCellSelection As Boolean)
    '----------------------------
    'VARIABLES
    Dim oWb As Workbook
    Dim oCurrentWorkbook As Workbook
    Dim bCurrentSelectionOk As Boolean
    Dim vStrFileToOpen As Variant
    Dim vCurrentSelection As Variant
    Dim sTargetCell As String
    Dim sSourceCellStart As String
    Dim sFormula As String
    '----------------------------
    ' CONSTS
    'Const csBasicFormula As String = "='C:\Users\Kathrin\Desktop\AG Sabass\[Borrelia burgdorferi strain persica.xlsx]Tabelle1'!$A2&ZEICHEN(10)&'C:\Users\Kathrin\Desktop\AG Sabass\[Borrelia burgdorferi strain persica.xlsx]Tabelle1'!$B2&ZEICHEN(10)&'C:\Users\Kathrin\Desktop\AG Sabass\[Borrelia burgdorferi strain persica.xlsx]Tabelle1'!$C2&ZEICHEN(10)&'C:\Users\Kathrin\Desktop\AG Sabass\[Borrelia burgdorferi strain persica.xlsx]Tabelle1'!$D2"
    'Const csBasicFormula As String = "='§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SheetName'!$§§§CellString1&ZEICHEN(10)&§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SheetName'!$§§§CellString1&ZEICHEN(10)&'C:&ZEICHEN(10)&'C:\Users\Kathrin\Desktop\AG Sabass\[Borrelia burgdorferi strain persica.xlsx]Tabelle1'!$D2"
    
    Const csPrompForDataFile As Variant = "Please select the file to transfer the data from"
    '----------------------------------------------------------------------------------------#
    Set oCurrentWorkbook = Application.ActiveWorkbook
    sTargetCell = ActiveCell.Address
    
    'bCurrentSelectionOk = bCeckCurrentSelection()
    'check if there is a cell selects which is the target cell for data...
    
    'let the user select the file...  TO DO: Replace the prompt by a const string, results in error at the moment
    vStrFileToOpen = Application.GetOpenFilename(, , "Please select the file to transfer the data from")
    'check if really a file has been selcted
    If vStrFileToOpen = False Then Exit Sub
    'open the selected file..
    Set oWb = Workbooks.Open(vStrFileToOpen, UpdateLinks:=0, ReadOnly:=0)
    oWb.Activate
    sFormula = getFormula(vStrFileToOpen)
    oWb.Close False
    'assign formula
    oCurrentWorkbook.Activate
    Range(sTargetCell).FormulaLocal = "=Summe"
    Debug.Print (sFormula)
    Range(sTargetCell).FormulaLocal = sFormula
End Sub


Function getFormula(sCurrentFileWithPath As Variant) As String
    Dim sFileNameWithExtention As String
    Dim sSourceFilePath As String
    Dim sFileName As String
    Dim sExtention As String
    Dim sSourcFileNameWithExtention As String
    Dim sSourcSheetName As String
    Dim sSourcCellString1 As String
    Dim sSourcCellString2 As String
    Dim sSourcCellString3 As String
    Dim sSourcCellString4 As String
    Dim sReturnFormula As String
    Dim objFso As Object
    Dim oVarUserInput As Range
    '----------------------
    getFormula = ""
    Set objFso = CreateObject("Scripting.FileSystemObject")
'// ------------------------------------------------------------------------
    sSourceFilePath = objFso.GetParentFolderName(sCurrentFileWithPath) + "\"
    sFileName = objFso.GetBaseName(sCurrentFileWithPath)
    sExtention = objFso.GetExtensionName(sCurrentFileWithPath)
    sSourcFileNameWithExtention = sFileName + "." + sExtention
    sSourcSheetName = ActiveWorkbook.ActiveSheet.Name
        
    Set oVarUserInput = Application.InputBox("Please select the start source cell", Type:=8)
    sSourcCellString1 = "$" + Chr(oVarUserInput.Column + 64)
    sSourcCellString1 = sSourcCellString1 + CStr(oVarUserInput.Row)
    sSourcCellString2 = getIncreasColumCellIndex(sSourcCellString1)
    sSourcCellString3 = getIncreasColumCellIndex(sSourcCellString2)
    sSourcCellString4 = getIncreasColumCellIndex(sSourcCellString3)
    sReturnFormula = Replace(csBasicFormula, "§§§SourceFilePath", sSourceFilePath)
    sReturnFormula = Replace(sReturnFormula, "§§§SourcFileNameWithExtention", sSourcFileNameWithExtention)
    sReturnFormula = Replace(sReturnFormula, "§§§SourcSheetName", sSourcSheetName)
    sReturnFormula = Replace(sReturnFormula, "$§§§SourcCellString1", sSourcCellString1)
    sReturnFormula = Replace(sReturnFormula, "$§§§SourcCellString2", sSourcCellString2)
    sReturnFormula = Replace(sReturnFormula, "$§§§SourcCellString3", sSourcCellString3)
    sReturnFormula = Replace(sReturnFormula, "$§§§SourcCellString4", sSourcCellString4)
    
    getFormula = sReturnFormula
End Function
Function getIncreasColumCellIndex(sCurrentCell As String) As String
    Dim sCurrentColumn As String
    Dim sIncreasedColum As String
    Dim sResultString As String
    '-----------------------------
    getIncreasColumCellIndex = ""
    sCurrentColumn = Mid(sCurrentCell, 2, 1)
    sIncreasedColum = Chr(Asc(sCurrentColumn) + 1)
    sResultString = Replace(sCurrentCell, sCurrentColumn, sIncreasedColum)
    getIncreasColumCellIndex = sResultString
End Function


Function bCeckCurrentSelection() As Boolean
    Dim vCurrentSelection As Variant
    '--------------------------------
    bCeckCurrentSelection = False
    Set vCurrentSelection = ActiveWindow.Selection

    'oTargetCells = ActiveCell.Range.Cells
    
    
    
End Function
