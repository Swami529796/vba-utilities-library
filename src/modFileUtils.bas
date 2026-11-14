Attribute VB_Name = "modFileUtils"
'==============================================================================
' modFileUtils - files, folders and paths
' Part of: https://github.com/Swami529796/vba-utilities-library
'
' This module has no dependencies and can be imported on its own.
'==============================================================================
Option Explicit

'---- Existence ---------------------------------------------------------------
Public Function FileExists(ByVal filePath As String) As Boolean
    If Len(filePath) = 0 Then Exit Function
    On Error Resume Next            ' GetAttr fails when the path does not exist
    FileExists = ((GetAttr(filePath) And vbDirectory) = 0)
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    If Len(folderPath) = 0 Then Exit Function
    On Error Resume Next
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
End Function

' Create a folder, including any missing parent folders. Safe to call if it exists.
Public Sub EnsureFolder(ByVal folderPath As String)
    folderPath = TrimSeparator(folderPath)
    If Len(folderPath) = 0 Then Exit Sub
    If FolderExists(folderPath) Then Exit Sub

    Dim parts() As String, current As String, i As Long, firstPart As Long
    parts = Split(folderPath, "\")

    If Left$(folderPath, 2) = "\\" Then
        ' network path: \\server\share is the root and cannot be created
        If UBound(parts) < 3 Then Exit Sub
        current = "\\" & parts(2) & "\" & parts(3)
        firstPart = 4
    Else
        current = parts(0)          ' drive, for example C:
        firstPart = 1
    End If

    For i = firstPart To UBound(parts)
        current = current & "\" & parts(i)
        If Not FolderExists(current) Then MkDir current
    Next i
End Sub

'---- Path parts --------------------------------------------------------------
' "C:\Data\Sales.xlsx" -> "Sales.xlsx"
Public Function FileNameOf(ByVal filePath As String) As String
    FileNameOf = Mid$(filePath, InStrRev(filePath, "\") + 1)
End Function

' "C:\Data\Sales.xlsx" -> "Sales"
Public Function FileBaseName(ByVal filePath As String) As String
    Dim fileName As String, dotAt As Long
    fileName = FileNameOf(filePath)
    dotAt = InStrRev(fileName, ".")
    If dotAt > 1 Then
        FileBaseName = Left$(fileName, dotAt - 1)
    Else
        FileBaseName = fileName
    End If
End Function

' "C:\Data\Sales.XLSX" -> "xlsx" (lower case, no dot)
Public Function FileExtensionOf(ByVal filePath As String) As String
    Dim fileName As String, dotAt As Long
    fileName = FileNameOf(filePath)
    dotAt = InStrRev(fileName, ".")
    If dotAt > 0 Then FileExtensionOf = LCase$(Mid$(fileName, dotAt + 1))
End Function

' "C:\Data\Sales.xlsx" -> "C:\Data"
Public Function ParentFolder(ByVal filePath As String) As String
    Dim slashAt As Long
    slashAt = InStrRev(TrimSeparator(filePath), "\")
    If slashAt > 0 Then ParentFolder = Left$(filePath, slashAt - 1)
End Function

' JoinPath("C:\Data\", "Sales.xlsx") and JoinPath("C:\Data", "Sales.xlsx") both give "C:\Data\Sales.xlsx"
Public Function JoinPath(ByVal folderPath As String, ByVal fileName As String) As String
    JoinPath = TrimSeparator(folderPath) & "\" & fileName
End Function

' "Report" + "xlsx" -> "Report_20260630_174501.xlsx" (never overwrites an earlier file)
Public Function TimestampedName(ByVal baseName As String, ByVal extension As String) As String
    TimestampedName = baseName & "_" & Format$(Now, "yyyymmdd_hhnnss") & "." & Replace(extension, ".", "")
End Function

'---- Listing -----------------------------------------------------------------
' Full paths of the files in a folder that match a pattern such as "*.xlsx".
' Excel lock files (~$...) are left out.
Public Function ListFiles(ByVal folderPath As String, Optional ByVal pattern As String = "*.*", _
                          Optional ByVal includeSubfolders As Boolean = False) As Collection
    Dim result As New Collection
    CollectFiles TrimSeparator(folderPath), pattern, includeSubfolders, result
    Set ListFiles = result
End Function

Private Sub CollectFiles(ByVal folderPath As String, ByVal pattern As String, _
                         ByVal includeSubfolders As Boolean, ByVal result As Collection)
    Dim itemName As String
    itemName = Dir(folderPath & "\" & pattern)
    Do While Len(itemName) > 0
        If Left$(itemName, 2) <> "~$" Then result.Add folderPath & "\" & itemName
        itemName = Dir()
    Loop

    If Not includeSubfolders Then Exit Sub

    ' Dir cannot be nested, so list the subfolders first and visit them afterwards
    Dim subfolders As New Collection, subfolder As Variant
    itemName = Dir(folderPath & "\*", vbDirectory)
    Do While Len(itemName) > 0
        If itemName <> "." And itemName <> ".." Then
            If FolderExists(folderPath & "\" & itemName) Then subfolders.Add folderPath & "\" & itemName
        End If
        itemName = Dir()
    Loop
    For Each subfolder In subfolders
        CollectFiles CStr(subfolder), pattern, True, result
    Next subfolder
End Sub

'---- Text files --------------------------------------------------------------
Public Function ReadTextFile(ByVal filePath As String) As String
    Dim f As Integer, content As String
    f = FreeFile
    Open filePath For Binary Access Read As #f
    content = Space$(LOF(f))
    Get #f, , content
    Close #f
    ReadTextFile = content
End Function

Public Sub WriteTextFile(ByVal filePath As String, ByVal content As String, _
                         Optional ByVal appendToFile As Boolean = False)
    Dim f As Integer
    f = FreeFile
    If appendToFile Then
        Open filePath For Append As #f
    Else
        Open filePath For Output As #f
    End If
    Print #f, content;
    Close #f
End Sub

'---- Dialogs -----------------------------------------------------------------
' Ask the user for a file. Returns "" if they cancel.
Public Function PickFile(Optional ByVal title As String = "Select a file", _
                         Optional ByVal filterName As String = "Excel files", _
                         Optional ByVal filterPattern As String = "*.xls*;*.csv") As String
    With Application.FileDialog(3)      ' msoFileDialogFilePicker
        .title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add filterName, filterPattern
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

' Ask the user for a folder. Returns "" if they cancel.
Public Function PickFolder(Optional ByVal title As String = "Select a folder") As String
    With Application.FileDialog(4)      ' msoFileDialogFolderPicker
        .title = title
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

'---- Internals ---------------------------------------------------------------
Private Function TrimSeparator(ByVal path As String) As String
    path = Trim$(path)
    Do While Len(path) > 3 And Right$(path, 1) = "\"
        path = Left$(path, Len(path) - 1)
    Loop
    TrimSeparator = path
End Function
