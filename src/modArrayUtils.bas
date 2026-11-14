Attribute VB_Name = "modArrayUtils"
'==============================================================================
' modArrayUtils - helpers for working with arrays instead of cells
' Part of: https://github.com/Swami529796/vba-utilities-library
'
' Reading a range into an array, working in memory and writing it back once is
' usually 50-100 times faster than reading and writing cell by cell.
'
' This module has no dependencies and can be imported on its own.
'==============================================================================
Option Explicit

' True for an array with no elements, and for anything that is not an array
Public Function ArrayIsEmpty(ByRef arr As Variant) As Boolean
    ArrayIsEmpty = True
    If Not IsArray(arr) Then Exit Function
    On Error Resume Next            ' UBound fails on an unallocated array
    ArrayIsEmpty = (UBound(arr) < LBound(arr))
End Function

' Number of dimensions: 0 = not an array, 1 = list, 2 = table
Public Function ArrayDimensions(ByRef arr As Variant) As Long
    Dim n As Long, upper As Long
    If Not IsArray(arr) Then Exit Function
    On Error GoTo Done
    Do
        upper = UBound(arr, n + 1)
        n = n + 1
    Loop
Done:
    ArrayDimensions = n
End Function

' Number of elements in the first dimension (rows of a table, items of a list)
Public Function ArrayLength(ByRef arr As Variant) As Long
    If Not ArrayIsEmpty(arr) Then ArrayLength = UBound(arr, 1) - LBound(arr, 1) + 1
End Function

' Range -> 2-D array (1 To rows, 1 To columns). Unlike Range.Value this also
' returns a 2-D array for a single cell, so the calling code needs no special case.
Public Function RangeToArray(ByVal source As Range) As Variant
    Dim result As Variant
    If source.Cells.CountLarge = 1 Then
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = source.Value
    Else
        result = source.Value
    End If
    RangeToArray = result
End Function

' Write an array to a sheet in one operation, starting at topLeft.
' A 2-D array is written as a table, a 1-D array as a single column.
Public Sub ArrayToRange(ByRef arr As Variant, ByVal topLeft As Range)
    If ArrayIsEmpty(arr) Then Exit Sub

    If ArrayDimensions(arr) = 2 Then
        topLeft.Resize(UBound(arr, 1) - LBound(arr, 1) + 1, _
                       UBound(arr, 2) - LBound(arr, 2) + 1).Value = arr
    Else
        Dim columnArr() As Variant, i As Long
        ReDim columnArr(1 To ArrayLength(arr), 1 To 1)
        For i = LBound(arr) To UBound(arr)
            columnArr(i - LBound(arr) + 1, 1) = arr(i)
        Next i
        topLeft.Resize(UBound(columnArr, 1), 1).Value = columnArr
    End If
End Sub

' Distinct non-blank values of a 1-D or 2-D array, in the order first seen.
' Returns a zero-based 1-D array (empty if there are no values).
Public Function UniqueValues(ByRef arr As Variant, Optional ByVal ignoreCase As Boolean = True) As Variant
    Dim seen As Object, v As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = IIf(ignoreCase, vbTextCompare, vbBinaryCompare)

    If Not ArrayIsEmpty(arr) Then
        For Each v In arr
            If Not IsError(v) Then
                If Len(CStr(v)) > 0 Then
                    If Not seen.Exists(v) Then seen.Add v, True
                End If
            End If
        Next v
    End If
    UniqueValues = seen.Keys
End Function

' True if the value is anywhere in the array
Public Function ArrayContains(ByRef arr As Variant, ByVal value As Variant, _
                              Optional ByVal ignoreCase As Boolean = True) As Boolean
    Dim v As Variant
    If ArrayIsEmpty(arr) Then Exit Function

    For Each v In arr
        If Not IsError(v) Then
            If VarType(v) = vbString And VarType(value) = vbString Then
                If StrComp(v, value, IIf(ignoreCase, vbTextCompare, vbBinaryCompare)) = 0 Then
                    ArrayContains = True
                    Exit Function
                End If
            ElseIf v = value Then
                ArrayContains = True
                Exit Function
            End If
        End If
    Next v
End Function

' Sort a 1-D array in place (quicksort). Works for numbers, dates and text.
Public Sub SortArray(ByRef arr As Variant, Optional ByVal descending As Boolean = False)
    If ArrayLength(arr) < 2 Then Exit Sub
    QuickSort arr, LBound(arr), UBound(arr)

    If descending Then
        Dim i As Long, j As Long, temp As Variant
        i = LBound(arr)
        j = UBound(arr)
        Do While i < j
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        Loop
    End If
End Sub

Private Sub QuickSort(ByRef arr As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Variant, temp As Variant
    i = low
    j = high
    pivot = arr((low + high) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then QuickSort arr, low, j
    If i < high Then QuickSort arr, i, high
End Sub

' One column of a 2-D array as a 1-D array (1-based)
Public Function GetColumn(ByRef data As Variant, ByVal columnIndex As Long, _
                          Optional ByVal skipHeader As Boolean = True) As Variant
    Dim firstRow As Long, r As Long, result() As Variant
    firstRow = LBound(data, 1) + IIf(skipHeader, 1, 0)
    If firstRow > UBound(data, 1) Then
        GetColumn = Array()
        Exit Function
    End If

    ReDim result(1 To UBound(data, 1) - firstRow + 1)
    For r = firstRow To UBound(data, 1)
        result(r - firstRow + 1) = data(r, columnIndex)
    Next r
    GetColumn = result
End Function

' Rows of a 2-D array where one column equals a value (text is compared without case).
' The header row is kept when hasHeader is True. Returns a 2-D array (1-based).
Public Function FilterRows(ByRef data As Variant, ByVal columnIndex As Long, ByVal matchValue As Variant, _
                           Optional ByVal hasHeader As Boolean = True) As Variant
    Dim firstRow As Long, r As Long, c As Long, n As Long
    Dim keep() As Boolean
    firstRow = LBound(data, 1) + IIf(hasHeader, 1, 0)
    ReDim keep(LBound(data, 1) To UBound(data, 1))

    ' pass 1: decide which rows to keep
    If hasHeader Then
        keep(LBound(data, 1)) = True
        n = 1
    End If
    For r = firstRow To UBound(data, 1)
        If SameValue(data(r, columnIndex), matchValue) Then
            keep(r) = True
            n = n + 1
        End If
    Next r

    If n = 0 Then
        FilterRows = Array()
        Exit Function
    End If

    ' pass 2: copy them
    Dim result() As Variant, outRow As Long
    ReDim result(1 To n, 1 To UBound(data, 2) - LBound(data, 2) + 1)
    For r = LBound(data, 1) To UBound(data, 1)
        If keep(r) Then
            outRow = outRow + 1
            For c = LBound(data, 2) To UBound(data, 2)
                result(outRow, c - LBound(data, 2) + 1) = data(r, c)
            Next c
        End If
    Next r
    FilterRows = result
End Function

Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then Exit Function
    If VarType(a) = vbString Or VarType(b) = vbString Then
        SameValue = (StrComp(CStr(a), CStr(b), vbTextCompare) = 0)
    Else
        SameValue = (a = b)
    End If
End Function
