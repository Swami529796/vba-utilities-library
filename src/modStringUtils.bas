Attribute VB_Name = "modStringUtils"
'==============================================================================
' modStringUtils - text cleaning and parsing
' Part of: https://github.com/Swami529796/vba-utilities-library
'
' This module has no dependencies and can be imported on its own.
'==============================================================================
Option Explicit

' Tidy text that came from another system: removes line breaks, tabs and other
' non-printing characters, turns non-breaking spaces into normal ones, trims the
' ends and reduces repeated spaces to one.
Public Function CleanText(ByVal text As String) As String
    Dim i As Long, code As Long, result As String
    result = text
    For i = 1 To Len(result)
        code = AscW(Mid$(result, i, 1))
        If (code >= 0 And code < 32) Or code = 127 Or code = 160 Then
            Mid$(result, i, 1) = " "
        End If
    Next i

    result = Trim$(result)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    CleanText = result
End Function

' "jOHN   smith" -> "John Smith"
Public Function ProperCase(ByVal text As String) As String
    ProperCase = StrConv(LCase$(CleanText(text)), vbProperCase)
End Function

' "+91 (98) 7654-3210" -> "919876543210"
Public Function DigitsOnly(ByVal text As String) As String
    Dim i As Long, ch As String, result As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then result = result & ch
    Next i
    DigitsOnly = result
End Function

' PadLeft("42", 6) -> "000042". Text already long enough is returned unchanged.
Public Function PadLeft(ByVal text As String, ByVal totalWidth As Long, _
                        Optional ByVal padCharacter As String = "0") As String
    If Len(text) >= totalWidth Then
        PadLeft = text
    Else
        PadLeft = String$(totalWidth - Len(text), Left$(padCharacter, 1)) & text
    End If
End Function

' TextBetween("Order [A-1042] shipped", "[", "]") -> "A-1042". Returns "" if a marker is missing.
Public Function TextBetween(ByVal text As String, ByVal startMarker As String, ByVal endMarker As String) As String
    Dim startAt As Long, endAt As Long
    startAt = InStr(1, text, startMarker, vbTextCompare)
    If startAt = 0 Then Exit Function
    startAt = startAt + Len(startMarker)
    endAt = InStr(startAt, text, endMarker, vbTextCompare)
    If endAt = 0 Then Exit Function
    TextBetween = Mid$(text, startAt, endAt - startAt)
End Function

' " North, South ,, East " -> Array("North", "South", "East"): trimmed, blanks removed, zero-based
Public Function SplitTrim(ByVal text As String, Optional ByVal delimiter As String = ",") As Variant
    Dim parts() As String, result() As String, i As Long, n As Long
    If Len(Trim$(text)) = 0 Then
        SplitTrim = Array()
        Exit Function
    End If

    parts = Split(text, delimiter)
    ReDim result(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            result(n) = Trim$(parts(i))
            n = n + 1
        End If
    Next i

    If n = 0 Then
        SplitTrim = Array()
    Else
        ReDim Preserve result(0 To n - 1)
        SplitTrim = result
    End If
End Function

' How many times one piece of text appears in another (ignores case)
Public Function CountOccurrences(ByVal text As String, ByVal findText As String) As Long
    If Len(findText) = 0 Then Exit Function
    CountOccurrences = (Len(text) - Len(Replace(text, findText, "", , , vbTextCompare))) \ Len(findText)
End Function

' A basic shape check: one @, something before it, and a dot in the part after it.
' It catches typing mistakes; it does not prove the address exists.
Public Function IsValidEmail(ByVal text As String) As Boolean
    Dim atSign As Long, domain As String
    text = Trim$(text)
    If InStr(text, " ") > 0 Then Exit Function
    atSign = InStr(text, "@")
    If atSign < 2 Or atSign <> InStrRev(text, "@") Then Exit Function

    domain = Mid$(text, atSign + 1)
    If InStr(domain, ".") < 2 Then Exit Function
    If Right$(domain, 1) = "." Or InStr(domain, "..") > 0 Then Exit Function
    IsValidEmail = True
End Function
