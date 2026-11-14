Attribute VB_Name = "modErrorHandler"
'==============================================================================
' modErrorHandler - one consistent way to handle errors and Excel's state
' Part of: https://github.com/Swami529796/vba-utilities-library
'
' The pattern (see ExampleProcedure at the bottom):
'
'   Public Sub MyMacro()
'       On Error GoTo Fail
'       FastMode True
'       ' ... the work ...
'   Done:
'       FastMode False
'       Exit Sub
'   Fail:
'       HandleError "MyMacro"
'       Resume Done
'   End Sub
'
' Requires: modLogger
'==============================================================================
Option Explicit

Public Const APP_ERROR As Long = vbObjectError + 513   ' number used by RaiseAppError

Private mFastModeOn As Boolean
Private mSavedCalculation As XlCalculation
Private mLastError As String

' Turn off screen updating, events and automatic calculation while a macro runs,
' and put everything back exactly as it was afterwards.
Public Sub FastMode(ByVal turnOn As Boolean)
    With Application
        If turnOn Then
            If Not mFastModeOn Then mSavedCalculation = .Calculation
            mFastModeOn = True
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        Else
            .ScreenUpdating = True
            .EnableEvents = True
            If mFastModeOn Then .Calculation = mSavedCalculation
            .StatusBar = False
            mFastModeOn = False
        End If
    End With
End Sub

' Call from an error handler: restores Excel, logs the error and tells the user.
Public Sub HandleError(ByVal procedureName As String, _
                       Optional ByVal showMessage As Boolean = True, _
                       Optional ByVal friendlyMessage As String = "")
    ' read the error first: almost anything else clears the Err object
    Dim errorNumber As Long, errorText As String, errorLine As Long
    errorNumber = Err.Number
    errorText = Err.Description
    errorLine = Erl

    FastMode False

    mLastError = "Error " & errorNumber & ": " & errorText
    If errorNumber = APP_ERROR Then mLastError = errorText   ' our own errors need no number
    If errorLine > 0 Then mLastError = mLastError & " (line " & errorLine & ")"
    LogError procedureName, mLastError

    If showMessage Then
        If Len(friendlyMessage) = 0 Then friendlyMessage = "Something went wrong and the task was stopped."
        MsgBox friendlyMessage & vbLf & vbLf & mLastError & vbLf & vbLf & _
               "Details have been written to the '" & LOG_SHEET & "' sheet.", vbExclamation, procedureName
    End If
End Sub

' Raise a business-rule error with a clear message, for example:
'   If lastRow < 2 Then RaiseAppError "The Data sheet is empty."
Public Sub RaiseAppError(ByVal message As String, Optional ByVal source As String = "")
    Err.Raise APP_ERROR, source, message
End Sub

' Text of the most recent error passed to HandleError
Public Function LastErrorText() As String
    LastErrorText = mLastError
End Function

'------------------------------------------------------------------------------
' A template showing the full pattern. Copy it as the starting point for a macro.
'------------------------------------------------------------------------------
Public Sub ExampleProcedure()
    Const PROC As String = "ExampleProcedure"
    On Error GoTo Fail
    FastMode True
    LogInfo PROC, "Started"

    ' --- the real work goes here ---
    Dim total As Double
    total = Application.WorksheetFunction.Sum(ThisWorkbook.Worksheets(1).UsedRange)
    If total < 0 Then RaiseAppError "The total cannot be negative.", PROC

    LogInfo PROC, "Finished"
Done:
    FastMode False
    Exit Sub
Fail:
    HandleError PROC
    Resume Done
End Sub
