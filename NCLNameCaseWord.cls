VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NCLNameCaseWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' https://github.com/seagullua/NameCaseLib/blob/master/Library/NCL/NCLNameCaseWord.php
Option Explicit

Const чоловік = 1
Const жінка = 2

Const Називний = 0
Const Родовий = 1
Const Давальний = 2
Const Знахідний = 3
Const Орудний = 4
Const Місцевий = 5
Const Кличний = 6

Private Word
Private word_orig
Private namepart
Private genderMan
Private genderWoman
Private genderSolved
Private letterMask
Private isUpperCase
Private NameCases
Private Rule

'Sub makenew(iword As String)  '__construct($word)
'    word = ""
'    word_orig = ""
'    namePart = ""
'    genderMan = 0
'    genderWoman = 0
'    genderSolved = 0
'    letterMask = Array()
'    isUpperCase = False
'    NameCases = Array()
'    rule = 0
'    word_orig = iword
'    generatemask iword
'    word = NCLStr.strtolower(iword)
'End Sub

Function splitLetters(istr) As Variant
    Dim buff(), i
    ReDim buff(1 To Len(istr))
    For i = 1 To Len(istr)
        buff(i) = Mid(istr, i, 1)
    Next
    splitLetters = buff
End Function

Function strtolower(istr) As String
    strtolower = LCase(istr)
End Function

Function strtoupper(istr) As String
    strtoupper = UCase(istr)
End Function


Function islowercase(istr) As Boolean
    islowercase = istr = strtolower(istr)
End Function


Sub generatemask(iword)
    Dim letters(), imask(), i
    letters = splitLetters(iword)
    ReDim imask(1 To Len(iword))
    isUpperCase = True
    For i = 1 To Len(iword)
        If islowercase(letters(i)) Then
            imask(i) = "x"
            isUpperCase = False
        Else
            imask(i) = "X"
        End If
    Next
    letterMask = imask
End Sub

Function max(x, y As Variant) As Variant
  max = IIf(x > y, x, y)
End Function

Function min(x, y As Variant) As Variant
   min = IIf(x < y, x, y)
End Function


Sub returnMask()
    Dim i, splittedMask, masklength, icase, caselength, imax, LetterIndex, letter
    If isUpperCase Then
        For i = LBound(NameCases, 1) To UBound(NameCases, 1)
            NameCases(i) = NCLStr.strtoupper(NameCases(i))
        Next
    Else
        splittedMask = letterMask
        masklength = UBound(splittedMask) - LBound(splittedMask) + 1
        For i = LBound(NameCases, 1) To UBound(NameCases, 1)
            icase = NameCases(i)
            caselength = NCLStr.strlen(icase)
            imax = min(caselength, masklength)
            NameCases(i) = ""
            For LetterIndex = 0 To imax - 1
                letter = NCLStr.substr(icase, LetterIndex, 1)
                If splettedMask(LetterIndex) = "X" Then
                    letter = NCLStr.strtoupper(letter)
                End If
                namecase(i) = namecase(i) & letter
            Next
            NameCases(i) = NameCases(i) & NCLStr.substr(icase, imax, caselength - masklength)
        Next
    End If
End Sub

Sub setNameCases(inameCases As Variant, Optional is_return_mask As Boolean = False)
    NameCases = inameCases
'    If is_return_mask Then returnMask
End Sub

Function getNameCases()
    getNameCases = NameCases
End Function

Function getNameCase(inumber) As String
    getNameCase = NameCases(inumber)
End Function

Function gender()
    If genderSolved = 0 Then
        If genderMan >= genderWoman Then
            genderSolved = чоловік
        Else
            genderSolved = жінка
        End If
    End If
    gender = genderSolved
End Function

Sub setGender(чоловік, жінка)
    genderMan = чоловік
    genderWoman = жінка
End Sub

Sub setTrueGender(gender)
    genderSolved = gender
End Sub

Function getGender()
    Dim arrTemp(2)
    arrTemp(чоловік) = genderMan
    arrTemp(жінка) = genderWoman
    getGender = arrTemp
End Function

Sub setNamePart(inamePart)
    namepart = inamePart
End Sub

Function getNamePart() As String
    getNamePart = namepart
End Function

Function getWord() As String
    getWord = Word
End Function

Function getWordOrig() As String
    getWordOrig = word_orig
End Function

Function isGenderSolved() As Boolean
    If genderSolved > 0 Then
        isGenderSolved = True
    Else
        isGenderSolved = False
    End If
End Function

Sub setRule(ruleID)
    Rule = ruleID
End Sub

Sub Init(iword)
    Word = ""
    word_orig = ""
    namepart = ""
    genderMan = 0
    genderWoman = 0
    genderSolved = 0
    letterMask = Array()
    isUpperCase = False
    NameCases = Array()
    Rule = 0

    word_orig = iword
'    generatemask iword
    Word = strtolower(iword)
End Sub

