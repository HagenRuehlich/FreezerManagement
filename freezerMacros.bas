Attribute VB_Name = "FreezerMacros"
'-------------------------------------------------
'These parameters controll the layout of the cells to be filled
Const ciTextSize As Integer = 8
Const cdCellRowHightFactor As Double = 6#


'-----------------------------------------------------------------------------------
Const csBasicFormula As String = "='§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SourcSheetName'!$§§§SourcCellString1&ZEICHEN(10)&'§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SourcSheetName'!$§§§SourcCellString2&ZEICHEN(10)&'§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SourcSheetName'!$§§§SourcCellString3&ZEICHEN(10)&'§§§SourceFilePath[§§§SourcFileNameWithExtention]§§§SourcSheetName'!$§§§SourcCellString4"
Const csPromptSelectStartSourceCell As String = "Please select the start source cell"
Const csPrompForDataFile As String = "Please select the file to transfer the data from"
'-----------------------------------------------------------------------------------

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
    
    
    '----------------------------------------------------------------------------------------#
    Set oCurrentWorkbook = Application.ActiveWorkbook
    sTargetCell = ActiveCell.Address
    'let the user select the file...  TO DO: Replace the prompt by a const string, results in error at the moment
    vStrFileToOpen = openFile(csPrompForDataFile)
    'check if really a file has been selcted
    If vStrFileToOpen = False Then
        Exit Sub
    End If
    'open the selected file..
    Set oWb = Workbooks.Open(vStrFileToOpen, UpdateLinks:=0, ReadOnly:=0)
    oWb.Activate
    sFormula = getCellValueByUserSelection(vStrFileToOpen, bAutomatedCellSelection)
    oWb.Close False
    'assign formula
    If sFormula <> "" Then
        oCurrentWorkbook.Activate
        With Range(sTargetCell)
            .FormulaLocal = sFormula
            .Value = .Value
        End With
        FormatCell (sTargetCell)
    End If
End Sub
Function getCellValueByUserSelection(sCurrentFileWithPath As Variant, bAutomatedCellSelection As Boolean) As String
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
    Dim bCellSelected As Boolean
    '----------------------
    getCellValueByUserSelection = ""
    Set objFso = CreateObject("Scripting.FileSystemObject")
'// ------------------------------------------------------------------------
    sSourceFilePath = objFso.GetParentFolderName(sCurrentFileWithPath) + "\"
    sFileName = objFso.GetBaseName(sCurrentFileWithPath)
    sExtention = objFso.GetExtensionName(sCurrentFileWithPath)
    sSourcFileNameWithExtention = sFileName + "." + sExtention
    sSourcSheetName = ActiveWorkbook.ActiveSheet.Name
    If bAutomatedCellSelection Then
        bCellSelected = selectSingleCell(csPromptSelectStartSourceCell, oVarUserInput)
        If bCellSelected = False Then
            Exit Function
        End If
        sSourcCellString1 = "$" + Chr(oVarUserInput.Column + 64)
        sSourcCellString1 = sSourcCellString1 + CStr(oVarUserInput.Row)
        sSourcCellString2 = getIncreasColumCellIndex(sSourcCellString1)
        sSourcCellString3 = getIncreasColumCellIndex(sSourcCellString2)
        sSourcCellString4 = getIncreasColumCellIndex(sSourcCellString3)
    Else
        bCellSelected = selectFoursCells(csPromptSelectStartSourceCell, oVarUserInput)
        If bCellSelected = False Then
            Exit Function
        End If
    End If
    
    sReturnFormula = Replace(csBasicFormula, "§§§SourceFilePath", sSourceFilePath)
    sReturnFormula = Replace(sReturnFormula, "§§§SourcFileNameWithExtention", sSourcFileNameWithExtention)
    sReturnFormula = Replace(sReturnFormula, "§§§SourcSheetName", sSourcSheetName)
    sReturnFormula = Replace(sReturnFormula, "$§§§SourcCellString1", sSourcCellString1)
    sReturnFormula = Replace(sReturnFormula, "$§§§SourcCellString2", sSourcCellString2)
    sReturnFormula = Replace(sReturnFormula, "$§§§SourcCellString3", sSourcCellString3)
    sReturnFormula = Replace(sReturnFormula, "$§§§SourcCellString4", sSourcCellString4)
    getCellValueByUserSelection = sReturnFormula
End Function
Sub FormatCell(psCell As String)
    Range(psCell).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .RowHeight = getRequiredCellRowHeight(psCell)
        .Rows.EntireRow.AutoFit
    End With
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = ciTextSize
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
End Sub

Function getRequiredCellRowHeight(psCell As String) As Integer
    Dim iNumberMergedRows As Integer
'check if psCell is in an area of verdically merged cells...
    iNumberMergedRows = Range(psCell).MergeArea.Rows.Count
    getRequiredCellRowHeight = (ciTextSize * cdCellRowHightFactor) / iNumberMergedRows
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
