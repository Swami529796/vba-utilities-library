Attribute VB_Name = "modTests"
'==============================================================================
' modTests - self-tests for the utilities library
' Part of: https://github.com/Swami529796/vba-utilities-library
'
' Run: RunAllTests
' Every check is written to the "Tests" sheet as Pass or Fail, so a change that
' breaks a helper is noticed straight away. The tests also serve as examples of
' how each function is meant to be called.
'
' Requires: all other modules in this library
'==============================================================================
Option Explicit

Private Const TEST_SHEET As String = "Tests"

Private mSheet As Worksheet
Private mRow As Long
Private mPassed As Long
Private mFailed As Long
Private mGroup As String

Public Sub RunAllTests()
    RunTests True
End Sub

' Same tests without the message box at the end (for automated runs)
Public Sub RunAllTestsSilently()
    RunTests False
End Sub

Private Sub RunTests(ByVal showMessage As Boolean)
    Dim startTime As Double
    startTime = Timer

    Set mSheet = GetOrCreateSheet(TEST_SHEET, , True)
    mSheet.Range("A1:D1").Value = Array("Group", "Test", "Result", "Detail")
    mSheet.Range("A1:D1").Font.Bold = True
    mRow = 2
    mPassed = 0
    mFailed = 0

    FastMode True
    RunGroup "Strings"
    RunGroup "Arrays"
    RunGroup "Sheets"
    RunGroup "Files"
    RunGroup "Logging and errors"
    FastMode False

    mSheet.Columns("A:D").AutoFit
    mSheet.Cells(mRow + 1, 1).Value = "Summary"
    mSheet.Cells(mRow + 1, 2).Value = mPassed & " passed, " & mFailed & " failed"
    mSheet.Cells(mRow + 1, 3).Value = IIf(mFailed = 0, "Pass", "Fail")
    mSheet.Range(mSheet.Cells(mRow + 1, 1), mSheet.Cells(mRow + 1, 3)).Font.Bold = True
    mSheet.Activate

    If showMessage Then
        MsgBox mPassed & " passed, " & mFailed & " failed" & vbLf & _
               "Time taken: " & Format$(Timer - startTime, "0.0") & " sec", _
               IIf(mFailed = 0, vbInformation, vbExclamation), "Utilities library tests"
    End If
End Sub

' Run one group; an unexpected error in a group is recorded as a failure, not a crash
Private Sub RunGroup(ByVal groupName As String)
    mGroup = groupName
    On Error GoTo Fail
    Select Case groupName
        Case "Strings":            TestStrings
        Case "Arrays":             TestArrays
        Case "Sheets":             TestSheets
        Case "Files":              TestFiles
        Case "Logging and errors": TestLoggingAndErrors
    End Select
    Exit Sub
Fail:
    Check "Group ran without an unexpected error", False, "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub Check(ByVal testName As String, ByVal condition As Boolean, Optional ByVal detail As String = "")
    mSheet.Cells(mRow, 1).Value = mGroup
    mSheet.Cells(mRow, 2).Value = testName
    If condition Then
        mSheet.Cells(mRow, 3).Value = "Pass"
        mSheet.Cells(mRow, 3).Font.Color = RGB(0, 128, 0)
        mPassed = mPassed + 1
    Else
        mSheet.Cells(mRow, 3).Value = "Fail"
        mSheet.Cells(mRow, 3).Font.Color = RGB(192, 0, 0)
        mSheet.Cells(mRow, 3).Font.Bold = True
        mSheet.Cells(mRow, 4).Value = detail
        mFailed = mFailed + 1
    End If
    mRow = mRow + 1
End Sub

Private Sub CheckEqual(ByVal testName As String, ByVal actual As Variant, ByVal expected As Variant)
    Check testName, (actual = expected), "Expected [" & CStr(expected) & "] but got [" & CStr(actual) & "]"
End Sub

'------------------------------------------------------------------------------
Private Sub TestStrings()
    CheckEqual "CleanText trims and collapses spaces", CleanText("  Acme    Traders  "), "Acme Traders"
    CheckEqual "CleanText removes line breaks and tabs", CleanText("Line1" & vbLf & "Line2" & vbTab & "End"), "Line1 Line2 End"
    CheckEqual "CleanText handles non-breaking spaces", CleanText(Chr$(160) & "Total" & Chr$(160)), "Total"
    CheckEqual "ProperCase", ProperCase("jOHN   smith"), "John Smith"
    CheckEqual "DigitsOnly", DigitsOnly("+91 (98) 7654-3210"), "919876543210"
    CheckEqual "PadLeft pads short text", PadLeft("42", 6), "000042"
    CheckEqual "PadLeft leaves long text alone", PadLeft("1234567", 6), "1234567"
    CheckEqual "TextBetween finds the middle", TextBetween("Order [A-1042] shipped", "[", "]"), "A-1042"
    CheckEqual "TextBetween returns blank when a marker is missing", TextBetween("No markers", "[", "]"), ""
    CheckEqual "CountOccurrences ignores case", CountOccurrences("a,B,b,c", "b"), 2

    Dim parts As Variant
    parts = SplitTrim(" North, South ,, East ")
    CheckEqual "SplitTrim drops blanks", UBound(parts) + 1, 3
    CheckEqual "SplitTrim trims each part", parts(1), "South"
    Check "SplitTrim of blank text is empty", ArrayIsEmpty(SplitTrim("   "))

    Check "IsValidEmail accepts a normal address", IsValidEmail("name@example.com")
    Check "IsValidEmail rejects a missing domain dot", Not IsValidEmail("name@example")
    Check "IsValidEmail rejects two @ signs", Not IsValidEmail("a@b@example.com")
    Check "IsValidEmail rejects spaces", Not IsValidEmail("first last@example.com")
End Sub

'------------------------------------------------------------------------------
Private Sub TestArrays()
    Dim empty1 As Variant, list As Variant
    Check "ArrayIsEmpty: non-array", ArrayIsEmpty(empty1)
    Check "ArrayIsEmpty: Array()", ArrayIsEmpty(Array())
    list = Array(5, 3, 9, 1)
    Check "ArrayIsEmpty: filled array is not empty", Not ArrayIsEmpty(list)
    CheckEqual "ArrayLength", ArrayLength(list), 4
    CheckEqual "ArrayDimensions of a list", ArrayDimensions(list), 1

    SortArray list
    CheckEqual "SortArray ascending: first", list(0), 1
    CheckEqual "SortArray ascending: last", list(3), 9
    SortArray list, True
    CheckEqual "SortArray descending: first", list(0), 9

    Dim names As Variant
    names = Array("pear", "Apple", "banana")
    SortArray names
    CheckEqual "SortArray sorts text", names(2), "pear"

    Check "ArrayContains finds text without case", ArrayContains(Array("North", "South"), "south")
    Check "ArrayContains respects case when asked", Not ArrayContains(Array("North", "South"), "south", False)
    Check "ArrayContains finds a number", ArrayContains(Array(1, 2, 3), 2)

    Dim unique As Variant
    unique = UniqueValues(Array("East", "west", "EAST", "", "North", "West"))
    CheckEqual "UniqueValues removes repeats and blanks", UBound(unique) + 1, 3
    CheckEqual "UniqueValues keeps first-seen order", unique(1), "west"

    ' a small table: header + 4 rows
    Dim table(1 To 5, 1 To 3) As Variant, data As Variant
    table(1, 1) = "Region": table(1, 2) = "Product": table(1, 3) = "Units"
    table(2, 1) = "North":  table(2, 2) = "Laptop":  table(2, 3) = 10
    table(3, 1) = "South":  table(3, 2) = "Printer": table(3, 3) = 4
    table(4, 1) = "north":  table(4, 2) = "Monitor": table(4, 3) = 7
    table(5, 1) = "East":   table(5, 2) = "Laptop":  table(5, 3) = 2
    data = table
    CheckEqual "ArrayDimensions of a table", ArrayDimensions(data), 2

    Dim northRows As Variant
    northRows = FilterRows(data, 1, "North")
    CheckEqual "FilterRows keeps the header and matching rows", UBound(northRows, 1), 3
    CheckEqual "FilterRows copies the right values", northRows(3, 2), "Monitor"
    Check "FilterRows with no match and no header is empty", ArrayIsEmpty(FilterRows(data, 1, "West", False))

    Dim units As Variant
    units = GetColumn(data, 3)
    CheckEqual "GetColumn skips the header", ArrayLength(units), 4
    CheckEqual "GetColumn returns the values", units(1) + units(2) + units(3) + units(4), 23

    ' round trip through a worksheet
    Dim ws As Worksheet, back As Variant
    Set ws = GetOrCreateSheet("TestScratch", , True)
    ArrayToRange data, ws.Range("B2")
    back = RangeToArray(ws.Range("B2").Resize(5, 3))
    CheckEqual "ArrayToRange / RangeToArray round trip", back(4, 2), "Monitor"
    CheckEqual "RangeToArray of one cell is still 2-D", ArrayDimensions(RangeToArray(ws.Range("B2"))), 2
    ArrayToRange Array("a", "b", "c"), ws.Range("G1")
    CheckEqual "ArrayToRange writes a list down a column", ws.Range("G3").Value, "c"
    DeleteSheet ws
End Sub

'------------------------------------------------------------------------------
Private Sub TestSheets()
    Dim ws As Worksheet
    Check "SheetExists: missing sheet", Not SheetExists("No Such Sheet")
    Set ws = GetOrCreateSheet("TestScratch", , True)
    Check "GetOrCreateSheet creates the sheet", SheetExists("TestScratch")

    CheckEqual "LastRow of an empty sheet", LastRow(ws), 0
    CheckEqual "DataRange of an empty sheet is Nothing", (DataRange(ws) Is Nothing), True

    ws.Range("A1:C1").Value = Array("Date", " Region ", "Amount")
    ws.Range("A2:A6").Value = 1
    ws.Range("C2:C9").Value = 2
    ws.Range("E20").Interior.Color = vbYellow     ' formatting only: must be ignored

    CheckEqual "LastRow of the sheet", LastRow(ws), 9
    CheckEqual "LastRow of one column", LastRow(ws, 1), 6
    CheckEqual "LastColumn of the sheet", LastColumn(ws), 3
    CheckEqual "LastColumn of one row", LastColumn(ws, 1), 3
    CheckEqual "HeaderColumn ignores case and spaces", HeaderColumn(ws, "region"), 2
    CheckEqual "HeaderColumn returns 0 when missing", HeaderColumn(ws, "Customer"), 0
    CheckEqual "DataRange covers the used block", DataRange(ws).Address(False, False), "A1:C9"

    Set ws = GetOrCreateSheet("TestScratch", , True)
    CheckEqual "GetOrCreateSheet can clear an existing sheet", LastRow(ws), 0
    DeleteSheet ws
End Sub

'------------------------------------------------------------------------------
Private Sub TestFiles()
    CheckEqual "FileNameOf", FileNameOf("C:\Data\Sales.xlsx"), "Sales.xlsx"
    CheckEqual "FileBaseName", FileBaseName("C:\Data\Sales.2026.xlsx"), "Sales.2026"
    CheckEqual "FileExtensionOf is lower case", FileExtensionOf("C:\Data\Sales.XLSX"), "xlsx"
    CheckEqual "ParentFolder", ParentFolder("C:\Data\Sales.xlsx"), "C:\Data"
    CheckEqual "JoinPath with a trailing separator", JoinPath("C:\Data\", "a.csv"), "C:\Data\a.csv"
    CheckEqual "JoinPath without a trailing separator", JoinPath("C:\Data", "a.csv"), "C:\Data\a.csv"
    Check "TimestampedName has the expected shape", TimestampedName("Report", ".xlsx") Like "Report_########_######.xlsx"

    ' real files in a temporary folder
    Dim root As String, nested As String, filePath As String
    root = JoinPath(Environ$("TEMP"), "vba_utils_" & Format$(Now, "yyyymmddhhnnss"))
    nested = JoinPath(JoinPath(root, "level1"), "level2")

    Check "FolderExists: missing folder", Not FolderExists(root)
    EnsureFolder nested
    Check "EnsureFolder creates nested folders", FolderExists(nested)
    EnsureFolder nested
    Check "EnsureFolder is safe to call twice", FolderExists(nested)

    filePath = JoinPath(root, "notes.txt")
    Check "FileExists: missing file", Not FileExists(filePath)
    WriteTextFile filePath, "first"
    WriteTextFile filePath, " second", True
    Check "WriteTextFile creates the file", FileExists(filePath)
    CheckEqual "ReadTextFile returns what was written", ReadTextFile(filePath), "first second"
    Check "FileExists is False for a folder", Not FileExists(root)

    WriteTextFile JoinPath(nested, "deep.txt"), "x"
    WriteTextFile JoinPath(root, "data.csv"), "a,b"
    CheckEqual "ListFiles: top folder only", ListFiles(root).Count, 2
    CheckEqual "ListFiles: with a pattern", ListFiles(root, "*.csv").Count, 1
    CheckEqual "ListFiles: including subfolders", ListFiles(root, "*.txt", True).Count, 2

    ' tidy up
    Kill JoinPath(nested, "deep.txt")
    Kill filePath
    Kill JoinPath(root, "data.csv")
    RmDir nested
    RmDir JoinPath(root, "level1")
    RmDir root
    Check "Temporary folder removed", Not FolderExists(root)
End Sub

'------------------------------------------------------------------------------
Private Sub TestLoggingAndErrors()
    Dim before As Long
    LogSetLevel LevelDebug
    LogClear
    CheckEqual "LogClear empties the log", LogCount(), 0

    LogInfo "Tests", "An information message"
    LogWarning "Tests", "A warning message"
    CheckEqual "LogInfo and LogWarning add entries", LogCount(), 2

    LogSetLevel LevelWarning
    LogInfo "Tests", "This one is below the minimum level"
    CheckEqual "LogSetLevel filters lower levels", LogCount(), 2
    LogSetLevel LevelDebug

    ' file logging
    Dim logFile As String
    logFile = JoinPath(Environ$("TEMP"), "vba_utils_" & Format$(Now, "yyyymmddhhnnss") & ".log")
    LogToFile logFile
    LogInfo "Tests", "Written to the file as well"
    LogToFile ""
    Check "LogToFile writes to a text file", FileExists(logFile)
    If FileExists(logFile) Then
        Check "Log file contains the message", InStr(ReadTextFile(logFile), "Written to the file as well") > 0
        Kill logFile
    End If

    ' HandleError: an application error is logged and Excel's state is restored
    before = LogCount()
    FastMode True
    On Error Resume Next
    RaiseAppError "The Data sheet is empty.", "Tests"
    HandleError "TestLoggingAndErrors", False
    On Error GoTo 0
    CheckEqual "HandleError logs the error", LogCount(), before + 1
    CheckEqual "HandleError keeps the message", LastErrorText(), "The Data sheet is empty."
    Check "HandleError restores screen updating", Application.ScreenUpdating
    Check "HandleError restores events", Application.EnableEvents
    FastMode True       ' the test runner turns it off again at the end

    ' a runtime error carries its number
    Dim zero As Long
    On Error Resume Next
    before = 1 \ zero
    HandleError "TestLoggingAndErrors", False
    On Error GoTo 0
    Check "HandleError reports runtime errors with their number", LastErrorText() Like "Error 11:*"
    FastMode True
End Sub

Private Sub DeleteSheet(ByVal ws As Worksheet)
    Dim alerts As Boolean
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = alerts
End Sub
