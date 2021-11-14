Attribute VB_Name = "excelBasics"
'------------------------------------------------------------------------
' Excel related stuff.......
'------------------------------------------------------------------------
Public Function openFile(psPrompt As String) As Variant
    '----------------------
    Dim vPrompt As Variant
    '-----------------------------
    vPrompt = psPrompt
    openFile = Application.GetOpenFilename(, , vPrompt)
End Function
Public Function selectCells(psPrompt As String, poReturnRange As Range) As Boolean
    selectCells = True
    On Error GoTo ErrorHandler
        Set poReturnRange = Application.InputBox(psPrompt, Type:=8)
    Exit Function
   
ErrorHandler:
    selectCells = False
End Function
Public Function getIntegerExcelColumn(pStringValue As String) As Integer
    Dim iReturnValue As Integer
    iReturnValue = 0
    'works until "pStringValue" isn't larger then "Z"...
     If Len(pStringValue) = 1 Then
        iReturnValue = Asc(pStringValue) - 64
    End If
    getIntegerExcelColumn = iReturnValue
End Function

Public Function getCellValue(pWSName As String, piLine As Integer, psColumn As String) As Variant
    Dim iColumn As Integer
    iColumn = getIntegerExcelColumn(psColumn)
    getCellValue = getCellValueInt(pWSName, piLine, iColumn)
End Function

Public Function getCellValueInt(pWSName As String, piLine As Integer, piColumn As Integer) As Variant
    getCellValueInt = Worksheets(pWSName).Cells(piLine, piColumn).Value
End Function

Public Sub setCellValue(pWSName As String, piLine As Integer, psColumn As String, pvValue As Variant, psFormat As String)
    Sheets(pWSName).Range(psColumn & piLine).NumberFormat = psFormat
    Sheets(pWSName).Range(psColumn & piLine).HorizontalAlignment = xlHAlignCenter
    Sheets(pWSName).Range(psColumn & piLine).Value = pvValue
End Sub
