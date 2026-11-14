Attribute VB_Name = "modLogger"
'==============================================================================
' modLogger - simple logging to a worksheet and, optionally, a text file
' Part of: https://github.com/Swami529796/vba-utilities-library
'
'   LogInfo "ImportData", "Started"
'   LogWarning "ImportData", "3 rows had no date"
'   LogError "ImportData", "File not found: " & filePath
'
'   LogSetLevel LevelWarning          ' ignore Debug and Info from now on
'   LogToFile "C:\Logs\import.log"    ' also append every entry to a file
'   LogClear                          ' empty the Log sheet
'
' This module has no dependencies and can be imported on its own.
'==============================================================================
Option Explicit

Public Enum LogLevel
    LevelDebug = 0
    LevelInfo = 1
    LevelWarning = 2
    LevelError = 3
End Enum

Public Const LOG_SHEET As String = "Log"

Private mMinimumLevel As LogLevel   ' entries below this level are ignored
Private mLogFile As String          ' empty = no file logging

'---- Public API --------------------------------------------------------------
Public Sub LogDebug(ByVal source As String, ByVal message As String)
    WriteEntry LevelDebug, source, message
End Sub

Public Sub LogInfo(ByVal source As String, ByVal message As String)
    WriteEntry LevelInfo, source, message
End Sub

Public Sub LogWarning(ByVal source As String, ByVal message As String)
    WriteEntry LevelWarning, source, message
End Sub

Public Sub LogError(ByVal source As String, ByVal message As String)
    WriteEntry LevelError, source, message
End Sub

' Ignore entries below this level (default: everything is logged)
Public Sub LogSetLevel(ByVal minimumLevel As LogLevel)
    mMinimumLevel = minimumLevel
End Sub

' Also append each entry to a text file. Pass "" to stop file logging.
Public Sub LogToFile(ByVal filePath As String)
    mLogFile = filePath
End Sub

' Remove all entries from the Log sheet (the header row is kept)
Public Sub LogClear()
    Dim ws As Worksheet
    Set ws = LogSheet()
    If ws.UsedRange.Rows.Count > 1 Then
        ws.Rows("2:" & ws.Rows.Count).Clear
    End If
End Sub

' Number of entries currently on the Log sheet
Public Function LogCount() As Long
    Dim ws As Worksheet
    Set ws = LogSheet()
    LogCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
End Function

'---- Internals ---------------------------------------------------------------
Private Sub WriteEntry(ByVal level As LogLevel, ByVal source As String, ByVal message As String)
    If level < mMinimumLevel Then Exit Sub

    Dim ws As Worksheet, r As Long
    Set ws = LogSheet()
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 2).Value = LevelName(level)
    ws.Cells(r, 3).Value = source
    ws.Cells(r, 4).Value = message

    Select Case level
        Case LevelWarning
            ws.Cells(r, 2).Font.Color = RGB(204, 102, 0)
        Case LevelError
            ws.Cells(r, 2).Font.Color = RGB(192, 0, 0)
            ws.Cells(r, 2).Font.Bold = True
    End Select

    If Len(mLogFile) > 0 Then AppendToFile level, source, message
End Sub

Private Sub AppendToFile(ByVal level As LogLevel, ByVal source As String, ByVal message As String)
    Dim f As Integer
    On Error GoTo Fail      ' logging must never stop the calling code
    f = FreeFile
    Open mLogFile For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & LevelName(level) & vbTab & source & vbTab & message
    Close #f
    Exit Sub
Fail:
    On Error Resume Next
    Close #f
End Sub

Private Function LevelName(ByVal level As LogLevel) As String
    Select Case level
        Case LevelDebug:   LevelName = "DEBUG"
        Case LevelInfo:    LevelName = "INFO"
        Case LevelWarning: LevelName = "WARNING"
        Case Else:         LevelName = "ERROR"
    End Select
End Function

' The Log sheet, created with a header row the first time it is needed
Private Function LogSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET
    End If

    If Len(CStr(ws.Range("A1").Value)) = 0 Then
        ws.Range("A1:D1").Value = Array("Time", "Level", "Source", "Message")
        ws.Range("A1:D1").Font.Bold = True
        ws.Columns(1).ColumnWidth = 20
        ws.Columns(2).ColumnWidth = 10
        ws.Columns(3).ColumnWidth = 24
        ws.Columns(4).ColumnWidth = 80
    End If
    Set LogSheet = ws
End Function
