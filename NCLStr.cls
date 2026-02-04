VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NCLStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' https://github.com/seagullua/NameCaseLib/blob/master/Library/NCL/NCLStr.php

Function substr(istr As String, start As Integer, Optional length As Integer) As String
    If IsMissing(length) Then
        substr = Mid(istr, start)
    Else
        substr = Mid(istr, start, length)
    End If
End Function

Function strpos(haystack As String, needle As String, Optional offset As Integer) As Integer
    If IsMissing(offset) Then
        strpos = InStr(haystack, needle)
    Else
        strpos = InStr(Mid(haystack, offset), needle) + offset
    End If
End Function

Function strlen(istr As String) As Integer
    strlen = Len(istr)
End Function

Function strtolower(istr As String) As String
    strtolower = LCase(istr)
End Function

Function strtoupper(istr As String) As String
    strtoupper = UCase(istr)
End Function

Function strrpos(haystack As String, needle As String, Optional offset As Integer) As Integer
    If IsMissing(offset) Then
        strpos = InStrRev(haystack, needle)
    Else
        strpos = InStrRev(Mid(haystack, 1, offset), needle)
    End If
End Function

Function islowercase(istr As String) As Boolean
    islowercase = istr = strtolower(istr)
End Function

Function isUpperCase(istr As String) As Boolean
    isUpperCase = istr = strtoupper(istr)
End Function

Function splitLetters(istr As String) As Variant
    Dim buff(), i
    ReDim buff(Len(istr) - 1)
    For i = 1 To Len(istr)
        buff(i - 1) = Mid(istr, i, 1)
    Next
    splitLetters = buff
End Function

Function connectLetters(iArray As Variant) As String
    connectLetters = Join(iArray, "")
End Function

Function explode(pattern As String, istr As String) As Variant
    Debug.Print pattern, istr
    explode = ""
End Function
