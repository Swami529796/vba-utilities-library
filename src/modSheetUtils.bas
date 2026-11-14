Attribute VB_Name = "modSheetUtils"
'==============================================================================
' modSheetUtils - worksheets, last rows and header lookups
' Part of: https://github.com/Swami529796/vba-utilities-library
'
' This module has no dependencies and can be imported on its own.
'==============================================================================
Option Explicit

Public Function SheetExists(ByVal sheetName As String, Optional ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Return the sheet with this name, adding it at the end of the workbook if needed
Public Function GetOrCreateSheet(ByVal sheetName As String, Optional ByVal wb As Workbook, _
                                 Optional ByVal clearContents As Boolean = False) As Worksheet
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook

    If SheetExists(sheetName, wb) Then
        Set ws = wb.Worksheets(sheetName)
        If clearContents Then ws.Cells.Clear
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

' Last row that contains anything. Give a column number to check that column only.
' Unlike UsedRange this ignores cells that are merely formatted. Returns 0 for an empty sheet.
Public Function LastRow(ByVal ws As Worksheet, Optional ByVal columnIndex As Long = 0) As Long
    Dim found As Range
    If columnIndex > 0 Then
        Set found = ws.Columns(columnIndex).Find(What:="*", LookIn:=xlFormulas, _
                                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Else
        Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    End If
    If Not found Is Nothing Then LastRow = found.Row
End Function

' Last column that contains anything. Give a row number to check that row only.
Public Function LastColumn(ByVal ws As Worksheet, Optional ByVal rowIndex As Long = 0) As Long
    Dim found As Range
    If rowIndex > 0 Then
        Set found = ws.Rows(rowIndex).Find(What:="*", LookIn:=xlFormulas, _
                                           SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    Else
        Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    End If
    If Not found Is Nothing Then LastColumn = found.Column
End Function

' Column number of a header (ignores case and extra spaces). Returns 0 if it is not there.
' Looking columns up by name means the code keeps working when someone inserts a column.
Public Function HeaderColumn(ByVal ws As Worksheet, ByVal headerName As String, _
                             Optional ByVal headerRow As Long = 1) As Long
    Dim c As Long, v As Variant
    For c = 1 To LastColumn(ws, headerRow)
        v = ws.Cells(headerRow, c).Value
        If Not IsError(v) Then
            If StrComp(Trim$(CStr(v)), Trim$(headerName), vbTextCompare) = 0 Then
                HeaderColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

' The block from A1 to the last used row and column, or Nothing for an empty sheet
Public Function DataRange(ByVal ws As Worksheet) As Range
    Dim r As Long, c As Long
    r = LastRow(ws)
    c = LastColumn(ws)
    If r > 0 And c > 0 Then Set DataRange = ws.Range(ws.Cells(1, 1), ws.Cells(r, c))
End Function
