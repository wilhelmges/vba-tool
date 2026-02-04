VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NCLNameCaseUa_U"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' https://github.com/seagullua/NameCaseLib/blob/master/Library/NCLNameCaseUa.php
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

Private ready As Boolean
Private finished As Boolean
'Private words() As NCLNameCaseWord
Private words()
Private workingWord As String
Private last_add As String
'Private workingLastCache As Variant
Private lastRule As Integer
Private lastResult As Variant
Private indexN() As Variant
Private indexS() As Variant
Private indexF() As Variant
Public gender_koef As Integer

Private Const CaseCount = 7
Private Const vowels = "аеиоуіїєюя"
Private Const consonant = "бвгджзйклмнпрстфхцчшщ"
Private Const mconsonant = "дтзсцлнзр"
Private Const tconsonant = "бпвмфжчшгкхґдтзцслнр"
Private Const shyplyachi = "жчшщ"
Private Const neshyplyachi = "бвгдзклмнпрстфхц"
Private Const myaki = "ьюяєї"
Private Const gubni = "мвпбф"


Public Function getFirstNameCase(Optional number = 0)
    AllWordCases
    getFirstNameCase = getCasesConnected(indexN, number)
End Function

Public Function getSecondNameCase(Optional number = 0)
    AllWordCases
    getSecondNameCase = getCasesConnected(indexS, number)
End Function

Public Function getFatherNameCase(Optional number = 0)
    AllWordCases
    getFatherNameCase = getCasesConnected(indexF, number)
End Function


Public Function getFormatted(Optional casenum = 0, Optional format = "S N F")
Dim length, result, i, symbol
    AllWordCases
    If casenum < 1 Then
        getFormatted = getFormattedArray(format)
        Exit Function
    Else
        length = Len(format)
        result = ""
        For i = 1 To length
            symbol = substr(format, i, 1)
            Select Case symbol
                Case "S"
                    result = result & getSecondNameCase(casenum)
                Case "N"
                    result = result & getFirstNameCase(casenum)
                Case "F"
                    result = result & getFatherNameCase(casenum)
                Case Else
                    result = result & symbol
            End Select
        Next
    End If
    If Len(last_add) > 0 Then
        result = result & " " & last_add
    End If
    getFormatted = result
End Function


Public Function getFormattedArray(Optional format = "S N F")
Dim length, Line, result(), symbol, curCase, casesS(), casesF(), casesN(), i
    length = Len(format)
    ReDim result(1 To CaseCount)
    casesS = Array()
    casesN = Array()
    casesF = Array()
    If UBound(indexS) > -1 Then casesS = getCasesConnected(indexS)
    If UBound(indexN) > -1 Then casesN = getCasesConnected(indexN)
    If UBound(indexF) > -1 Then casesF = getCasesConnected(indexF)
    For curCase = 1 To CaseCount
        Line = ""
        For i = 1 To length
            symbol = substr(format, i, 1)
            Select Case symbol
                Case "S"
                    If UBound(indexS) > -1 Then Line = Line & casesS(curCase - 1)
                Case "N"
                    If UBound(indexN) > -1 Then Line = Line & casesN(curCase - 1)
                Case "F"
                    If UBound(indexF) > -1 Then Line = Line & casesF(curCase - 1)
                Case Else
                    Line = Line & symbol
            End Select
        Next
        If Len(last_add) > 0 Then
            Line = Line & " " & last_add
        End If
        result(curCase) = Line
    Next
'    If InStr(result(3), "ві ") <> InStrRev(result(3), "ві ") Then
'        result(3) = Left(result(3), InStrRev(result(3), "ві ") - 2) & "у" & Right(result(3), Len(result(3)) - InStrRev(result(3), "ві ") - 1)
'    End If
'    If InStr(result(6), "ві ") <> InStrRev(result(6), "ві ") Then
'        result(6) = Left(result(6), InStrRev(result(6), "ві ") - 2) & Right(result(6), Len(result(6)) - InStrRev(result(6), "ві "))
'    End If
    getFormattedArray = result
End Function

Public Function zv(fullname, Optional casenum = 0)

End Function

Public Function q(fullname, Optional casenum = 0, Optional gender = -1)
Dim format, z
    format = "S N F"
    fullReset
    splitFullName fullname
    If gender > -1 Then setGender gender
    q = getFormatted(casenum, format)

' Проблеми:
'   Поломарчук Леонід Віталійович
'   Пазенко Ігор Вікторович
'   Шкіль Руслан Олександрович
'   Сокоть Олег Анатолійович
'   Воронов Федір Петрович




'    Debug.Print z
'    q = z
End Function

Sub reset()
    lastRule = 0
    lastResult = Array()
End Sub

Sub fullReset()
    Erase words
    indexN = Array()
    indexS = Array()
    indexF = Array()
    reset
    notReady
End Sub

Sub notReady()
    ready = False
    finished = False
End Sub

Sub Rule(index As Integer)
    lastRule = index
End Sub

Sub setWorkingWord(Word)
    reset
    workingWord = Word
'    workingLastCache = Array()
End Sub

Sub array_fill(istart As Integer, count As Integer, ivalue)
   Dim thisarray As Variant
   ReDim thisarray(istart To istart + count) As Integer
   For counter = istart To istart + count
      thisarray(counter) = ivalue
   Next
End Sub

Sub makeResultTheSame()
    lastResult = array_fill(0, CaseCount, workingWord)
End Sub

Function substr(istr, start, Optional length)
    If IsMissing(length) Then
        substr = Mid(istr, start)
    Else
        substr = Mid(istr, start, length)
    End If
End Function

Function Last(Optional ByRef length = 1, Optional ByRef stopAfter = 0)
    Dim ccut, start
    If stopAfter = 0 Then ccut = length Else ccut = stopAfter
'    If Not workindLastCache(length, stopAfter) Is Nothing Then
'        workindLastCache(length, stopAfter) = NCLStr.substr(workingWord, Len(workingWord) - length, Cut)
'    End If
'    Last = workindLastCache(length, stopAfter)
    start = Len(workingWord) - length + 1
    If start < 1 Then start = 1
    Last = Mid(workingWord, start, ccut)
End Function

'Function RulesChain(gender As String, rulesArray As Variant) As Boolean
'    Dim ruleID, ruleMethod
'    For Each ruleID In rulesArray
'        ruleMethod = gender & "Rule" & ruleID
'        If Application.Run(ruleMethod) Then
'            RulesChain = True
'            Exit Function
'        End If
'    Next
'    RulesChain = False
'End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Function inThis(iletter As String, istring As Variant) As Boolean
    If IsArray(istring) Then
        inThis = IsInArray(iletterm, istring)
    Else
        If iletter = "" Or NCLStr.strpos(istring, iletter) = 0 Then
            inThis = False
        Else
            inThis = True
        End If
    End If
End Function

Function inNames(nameNeedle As String, names As Variant) As Boolean
    Dim name
    If Not IsArray(names) Then
        names = Array(names)
    End If
    For Each name In names
        If StrComp(name, nameNeedle, 1) = 0 Then
            inNames = True
        End If
    Next
    inNames = False
End Function

Sub wordForms(Word, endings, Optional replaceLast = 0)
    Dim result() As Variant
    Dim padegIndex As Integer
    ReDim result(0 To CaseCount - 1)
    result(0) = workingWord
    Word = Left(Word, Len(Word) - replaceLast)
    For padegIndex = 1 To CaseCount - 1
        result(padegIndex) = Word & endings(padegIndex - 1)
    Next
    lastResult = result
End Sub

Sub setFirstName(Optional firstname As String)
    Dim i As Integer
    Dim obj As NCLNameCaseWord
    If firstname Is Nothing Then
        i = count(words)
        ReDim Preserve words(i + 1)
        Set obj = New NCLNameCaseWord
        obj.Init (firstname)
        obj.setNamePart ("N")
        Set words(i) = obj
        notReady
    End If
End Sub

Sub setSecondName(Optional secondname As String)
    Dim i As Integer
    Dim obj As NCLNameCaseWord
    If secondname Is Nothing Then
        i = count(words)
        ReDim Preserve words(i + 1)
        Set obj = New NCLNameCaseWord
        obj.Init (secondname)
        obj.setNamePart ("S")
        Set words(i) = obj
        notReady
    End If
End Sub

Sub setFatherName(Optional fathername As String)
    Dim i As Integer
    Dim obj As NCLNameCaseWord
    If fathername Is Nothing Then
        i = count(words)
        ReDim Preserve words(i + 1)
        Set obj = New NCLNameCaseWord
        obj.Init (fathername)
        obj.setNamePart ("F")
        Set words(i) = obj
        notReady
    End If
End Sub

Sub setGender(gender)
    Dim iword
'    If gender Is Nothing Then
'        gender = 0
'    End If
    For Each iword In words
        iword.setTrueGender gender
    Next
End Sub

Sub setFullName(Optional secondname = "", Optional firstname = "", Optional fathername = "")
    setFirstName firstname
    setSecondName secondname
    setFatherName fathername
End Sub

Sub setName(Optional firstname = "")
    setFirstName firstname
End Sub

Sub setLastName(Optional secondname = "")
    setSecondName secondname
End Sub
' ????
Sub setSirName(Optional fathername = "")
    setFatherName fathername
End Sub

Sub prepareNamePart(iword)
    If iword.getNamePart() = "" Then
        detectNamePart iword
    End If
End Sub

Sub prepareAllNameParts()
Dim n
    For n = 1 To UBound(words)
        Select Case n
        Case 1
            words(1).setNamePart "S"
        Case 2
            words(2).setNamePart "N"
        Case 3
            words(3).setNamePart "F"
        End Select
    Next
'    words(1).setNamePart "S"
'    words(2).setNamePart "N"
'    words(3).setNamePart "F"
''    Dim iword
''    For Each iword In words
''        prepareNamePart iword
''    Next
''    Debug.Print "we here"
End Sub

Sub prepareGender(iword)
    Dim namepart As String
    If Not iword.isGenderSolved() Then
        namepart = iword.getNamePart()
                
        Select Case namepart
            Case "N"
                GenderByFirstName iword
            Case "F"
                GenderByFatherName iword
            Case "S"
                GenderBySecondName iword
        End Select
    End If
End Sub

Function solveGender() As Boolean
    Dim iword, gender
    Dim man, woman
    For Each iword In words
        If iword.isGenderSolved() Then
            setGender iword.gender()
            solveGender = True
            Exit Function
        End If
    Next
    man = 0
    woman = 0
    ' тут
    For Each iword In words
        prepareGender iword
        gender = iword.getGender()
        man = man + gender(чоловік)
        woman = woman + gender(жінка)
    Next
    If (man > woman) Then
        setGender чоловік
    Else
        setGender жінка
    End If
    solveGender = True
End Function

Sub generateIndex()
'Debug.Print "we here"
Dim i, namepart
    indexS = Array()
    indexF = Array()
    indexN = Array()
    For i = LBound(words) To UBound(words)
        namepart = words(i).getNamePart
        Select Case namepart
            Case "N"
                ReDim Preserve indexN(0 To UBound(indexN) + 1)
                indexN(UBound(indexN)) = i
            Case "F"
                ReDim Preserve indexF(0 To UBound(indexF) + 1)
                indexF(UBound(indexF)) = i
            Case "S"
                ReDim Preserve indexS(0 To UBound(indexS) + 1)
                indexS(UBound(indexS)) = i
        End Select
    Next
End Sub

Sub prepareEverything()
    If Not ready Then
        prepareAllNameParts
        solveGender
        generateIndex
        ready = True
    End If
End Sub

Public Function genderDetect(fullname)
    fullReset
    splitFullName fullname
    genderDetect = genderAutoDetect()
End Function

Function genderAutoDetect() As Integer
    Dim m As Integer, n As Integer, max_koef As Integer, Word As NCLNameCaseWord, genders(), imin, imax, koef
    prepareEverything
    If count(words) > 0 Then
        n = -1
        max_koef = -1
        For m = 0 To UBound(words)
            Set Word = words(m)
            genders = Word.getGender()
            imin = min(genders)
            imax = max(genders)
            koef = max - min
            If koef > max_koef Then
                max_koef = koef
                n = m
            End If
        Next
        If n >= 0 Then
            If Not Nothing Is words(n) Then
                genders = words(n).getGender()
                imin = min(genders)
                imax = max(genders)
                gender_koef = imax - imin
                genderAutoDetect = words(n).gender()
                Exit Function
            End If
        End If
    End If
    genderAutoDetect = False
End Function

Sub splitFullName(fullname)
    Dim list() As String, Word, iword As NCLNameCaseWord, formatArray(), n
    fullname = Application.trim(fullname)
    list = Split(fullname, " ")
    last_add = ""
    n = 0
    For Each Word In list
        n = n + 1
        If n > 3 Then Exit For
        ReDim Preserve words(1 To n)
    Next
    n = 0
    For Each Word In list
        n = n + 1
        If n = 4 Then
            last_add = Word
        ElseIf n > 4 Then
            last_add = last_add & " " & Word
        Else
            Set iword = New NCLNameCaseWord
            iword.Init Word
            Set words(n) = iword
        End If
    Next
    prepareEverything
' ????
'    For Each iword In words
'        ReDim formatArray(count(formatArray) + 1)
'        formatArray(count(formatArray) - 1) = iword.getNamePart()
'    Next
'    splitFullName = formatArray
End Sub

Function getFullNameFormat(fullname As String) As String
    Dim format As String, iword As NCLNameCaseWord
    fullReset
    splitFullName (fullname)
    format = ""
    For Each iword In words
        format = format & iword.getNamePart() & " "
    Next
    getFullNameFormat = format
End Function

Sub WordCase(iword)
    Dim gender, namepart, name_part_letter, method, tmp, cur_words
    Dim o_cur_words, result(CaseCount - 1), last_rule, cnt, k, i, cur_word, o_ncw, result_tmp(), inameCases, namecase, res
    gender = IIf(iword.gender() = чоловік, "man", "woman")
    namepart = ""
    name_part_letter = iword.getNamePart()
    Select Case name_part_letter
        Case "F"
            namepart = "Father"
        Case "N"
            namepart = "First"
        Case "S"
            namepart = "Second"
    End Select
    method = gender & namepart & "Name"
    tmp = iword.getWordOrig()
    cur_words = Split(tmp, "-")
    o_cur_words = Array()
'    result = Array(CaseCount)
    last_rule = -1
    cnt = UBound(cur_words) - LBound(cur_words) + 1
    For k = 0 To cnt - 1
        cur_word = cur_words(k)
        Set o_ncw = New NCLNameCaseWord
        o_ncw.Init (cur_word)
        setWorkingWord (cur_word)
        Debug.Print method
        Select Case method
            Case "manFirstName"
                res = manFirstName
            Case "womanFirstName"
                res = womanFirstName
            Case "manSecondName"
                res = manSecondName
            Case "womanSecondName"
                res = womanSecondName
            Case "manFatherName"
                res = manFatherName
            Case "womanFatherName"
                res = womanFatherName
            Case Else
                Debug.Print "Method " & method & " not found!"
        End Select
        If res Then
            result_tmp = lastResult
            last_rule = lastRule
        Else
            ReDim result_tmp(0 To CaseCount - 1)
            For i = 0 To CaseCount - 1
                result_tmp(i) = cur_word
            Next
            last_rule = -1
        End If
        o_ncw.setNameCases (result_tmp)
        ReDim Preserve o_cur_words(0 To UBound(o_cur_words) - LBound(o_cur_words) + 1)
        Set o_cur_words(UBound(o_cur_words) - LBound(o_cur_words)) = o_ncw
    Next
    For Each o_ncw In o_cur_words
        inameCases = o_ncw.getNameCases()
        For k = 0 To UBound(inameCases) - LBound(inameCases)
            namecase = inameCases(k)
            If Not result(k) = Empty Then
                result(k) = result(k) & "-" & namecase
            Else
                result(k) = namecase
            End If
        Next
    Next
    iword.setNameCases result, False
    iword.setRule last_rule
End Sub

Sub AllWordCases()
Dim n
    If Not finished Then
        prepareEverything
        For n = 1 To UBound(words)
            WordCase words(n)
        Next
        finished = True
    End If
End Sub

Function getWordCase(iword, Optional number = 0) As Variant
    Dim icases
    icases = iword.getNameCases
    If number < 1 Or number > CaseCount Then
        getWordCase = icases
    Else
        getWordCase = icases(number - 1)
    End If
End Function

Function getCasesConnected(indexArray, Optional number = 0)
Dim resultArray(CaseCount - 1), readyArray() As Variant, idx, all, m, i, tmp
    readyArray = Array()
    If UBound(indexArray) > -1 Then
        For Each idx In indexArray
            ReDim Preserve readyArray(0 To UBound(readyArray) - LBound(readyArray) + 1)
            readyArray(UBound(readyArray) - LBound(readyArray)) = getWordCase(words(idx), number)
        Next
        all = UBound(readyArray) - LBound(readyArray) + 1
    Else
        all = 0
    End If
    If all > 0 Then
        If IsArray(readyArray(0)) Then
            For m = 0 To CaseCount - 1
                tmp = ""
                For i = 0 To all - 1
                    tmp = tmp & readyArray(i)(m) & " "
                Next
                resultArray(m) = Left(tmp, Len(tmp) - 1)
            Next
            getCasesConnected = resultArray
            Exit Function
        Else
            getCasesConnected = Join(readyArray, " ")
            Exit Function
        End If
    End If
    getCasesConnected = ""
End Function

' ***********************************************************************************************************************************************************************************************
'     * Чергування українських приголосних
'     * Чергування г к х —» з ц с
'     * @param string $letter літера, яку необхідно перевірити на чергування
'     * @return string літера, де вже відбулося чергування
Private Function inverseGKH(letter)
    inverseGKH = letter
    Select Case letter
        Case "г"
            inverseGKH = "з"
        Case "к"
            inverseGKH = "ц"
        Case "х"
            inverseGKH = "с"
    End Select
End Function

'     * Перевіряє чи символ є апострофом чи не є
'     * @param string(1) $char симпол для перевірки
'     * @return bool true якщо символ є апострофом
Private Function isApostrof(char)
    If InStr(1, " " & consonant & vowels, char, vbBinaryCompare) > 0 Then
        isApostrof = False
    Else
        isApostrof = True
    End If
End Function

'     * Чергування українських приголосних
'     * Чергування г к —» ж ч
'     * @param string $letter літера, яку необхідно перевірити на чергування
'     * @return string літера, де вже відбулося чергування
Private Function inverse2(letter)
    inverse2 = letter
    Select Case letter
        Case "к"
            inverse2 = "ч"
        Case "г"
            inverse2 = "ж"
    End Select
End Function

'     * <b>Визначення групи для іменників 2-ї відміни</b>
'     * 1 - тверда
'     * 2 - мішана
'     * 3 - м’яка
'     *
'     * <b>Правило:</b>
'     * - Іменники з основою на твердий нешиплячий належать до твердої групи:
'     *   береза, дорога, Дніпро, шлях, віз, село, яблуко.
'     * - Іменники з основою на твердий шиплячий належать до мішаної групи:
'     *   пожеж-а, пущ-а, тиш-а, алич-а, вуж, кущ, плющ, ключ, плече, прізвище.
'     * - Іменники з основою на будь-який м'який чи пом'якше­ний належать до м'якої групи:
'     *   земля [земл'а], зоря [зор'а], армія [арм'ійа], сім'я [с'імйа], серпень, фахівець,
'     *   трамвай, су­зір'я [суз'ірйа], насіння [насін'н'а], узвишшя Іузвиш'ш'а
'     * @param string $word іменник, групу якого необхідно визначити
'     * @return int номер групи іменника
Private Function detect2Group(Word)
Dim osnova, stack, stacksize, lLast, osnovaEnd
    osnova = Word
    stack = Array()
    stacksize = 0
    Do While InStr(vowels & "ь", Right(osnova, 1)) > 0
        ReDim Preserve stack(0 To stacksize)
        stack(stacksize) = Right(osnova, 1)
        stacksize = stacksize + 1
        osnova = Left(osnova, Len(osnova) - 1)
    Loop
    lLast = "Z"
    If stacksize > 0 Then
        lLast = stack(stacksize - 1)
    End If
    osnovaEnd = Right(osnova, 1)
    If InStr(neshyplyachi, osnovaEnd) > 0 And Not InStr(myaki, lLast) > 0 Then
        detect2Group = 1
    Else
        If InStr(shyplyachi, osnovaEnd) > 0 And Not InStr(myaki, lLast) > 0 Then
            detect2Group = 2
        Else
            detect2Group = 3
        End If
    End If
End Function

'     * Шукаємо в слові <var>$word</var> перше входження літери з переліку <var>$vowels</var> з кінця
'     * @param string $word слово, якому необхідно знайти голосні
'     * @param string $vowels перелік літер, які треба знайти
'     * @return string(1) перша з кінця літера з переліку <var>$vowels</var>
Private Function FirstLastVowel(Word, vowels)
Dim length, i, char
    length = Len(Word)
    FirstLastVowel = ""
    For i = length - 1 To 1 Step -1
        char = Mid(Word, i, 1)
        If InStr(vowels, char) > 0 Then
            FirstLastVowel = char
            Exit For
        End If
    Next
End Function

'     * Пошук основи іменника <var>$word</var>
'     * <b>Основа слова</b> - це частина слова (як правило незмінна), яка вказує на його лексичне значення.
'     * @param string $word слово, в якому необхідно знати основу
'     * @return string основа іменника <var>$word</var>
Private Function getOsnova(Word)
Dim osnova
    osnova = Word
    Do While InStr(vowels & "ь", Right(osnova, 1)) > 0
        osnova = Left(osnova, Len(osnova) - 1)
    Loop
    getOsnova = osnova
End Function
     
'     * Українські чоловічі та жіночі імена, що в називному відмінку однини закінчуються на -а (-я),
'     * відмінються як відповідні іменники І відміни.
'     * <ul>
'     * <li>Примітка 1. Кінцеві приголосні основи г, к, х у жіночих іменах
'     *   у давальному та місцевому відмінках однини перед закінченням -і
'     *   змінюються на з, ц, с: Ольга - Ользі, Палажка - Палажці, Солоха - Солосі.</li>
'     * <li>Примітка 2. У жіночих іменах типу Одарка, Параска в родовому відмінку множини
'     *   в кінці основи між приголосними з'являється звук о: Одарок, Парасок. </li>
'     * </ul>
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function manRule1()
Dim BeforeLast, Last1
    Last1 = Last(1)
    BeforeLast = Last(2, 1)
    ' Debug.Print "BeforeLast=" & BeforeLast & ", Last(1)" & Last1
    If Last1 = "а" Then
        If InStr(shyplyachi, BeforeLast) > 0 Then
            If BeforeLast = "ж" Then
                wordForms workingWord, Array(BeforeLast & "и", inverseGKH(BeforeLast) & "і", BeforeLast & "у", BeforeLast & "ею", inverseGKH(BeforeLast) & "і", BeforeLast & "е"), 2
            Else
                wordForms workingWord, Array(BeforeLast & "і", inverseGKH(BeforeLast) & "і", BeforeLast & "у", BeforeLast & "ею", inverseGKH(BeforeLast) & "і", BeforeLast & "е"), 2
            End If
        Else
            wordForms workingWord, Array(BeforeLast & "и", inverseGKH(BeforeLast) & "і", BeforeLast & "у", BeforeLast & "ою", inverseGKH(BeforeLast) & "і", BeforeLast & "о"), 2
        End If
        Rule 101
        manRule1 = True
        Exit Function
    ElseIf Last1 = "я" Then
        If BeforeLast = "і" Then
            wordForms workingWord, Array("ї", "ї", "ю", "єю", "ї", "є"), 1
            Rule 102
            manRule1 = True
            Exit Function
        ElseIf BeforeLast = "л" Then
            wordForms workingWord, Array(BeforeLast & "і", inverseGKH(BeforeLast) & "і", BeforeLast & "ю", BeforeLast & "ею", inverseGKH(BeforeLast) & "і", BeforeLast & "є"), 2
            Rule 103
            manRule1 = True
            Exit Function
        Else
            wordForms workingWord, Array(BeforeLast & "і", inverseGKH(BeforeLast) & "і", BeforeLast & "ю", BeforeLast & "ею", inverseGKH(BeforeLast) & "і", BeforeLast & "е"), 2
            Rule 104
            manRule1 = True
            Exit Function
        End If
    End If
    manRule1 = False
End Function

'     * Імена, що в називному відмінку закінчуються на -р, у родовому мають закінчення -а:
'     * Віктор - Віктора, Макар - Макара, але: Ігор - Ігоря, Лазар - Лазаря.
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function manRule2()
Dim osnova, Last1, Last2, Last3, Last4, Last21, Last41, delim
    Last1 = Last(1)
    Last2 = Last(2)
    Last3 = Last(3)
    Last4 = Last(4)
    Last21 = Last(2, 1)
    Last41 = Last(4, 1)
    If Last1 = "р" Then
        If (Last2 = "ар" Or Last2 = "ир") And Last3 <> "вар" And Last3 <> "зар" And Not (InStr("гайдар гончар володимир назар", StrConv(workingWord, vbLowerCase)) > 0) Then
            wordForms workingWord, Array("я", "ю", "я", "ем", "ю", "ю")
            Rule 201
            manRule2 = True
            Exit Function
        ElseIf Last2 = "яр" Then 'Or Last3 = "нір" Then
            If Last3 = "ляр" Then
                wordForms workingWord, Array("а", "у", "а", "ем", "у", "е")
            Else
                wordForms workingWord, Array("а", "у", "а", "ем", "у", "е")
            End If
            Rule 202
            manRule2 = True
            Exit Function
        ElseIf Last3 = "нір" Or InStr("Авенір Олефір Лігер", StrConv(workingWord, vbProperCase)) > 0 Then
            wordForms workingWord, Array("а", "у", "а", "ом", "у", "е")
            Rule 202
            manRule2 = True
            Exit Function
        ElseIf Last3 = "вір" Then
            wordForms workingWord, Array("а", "у", "а", "ом", "і", "е")
            Rule 202
            manRule2 = True
            Exit Function
        ElseIf InStr("ігор лазар", StrConv(workingWord, vbLowerCase)) > 0 Then
            wordForms workingWord, Array("я", "ю", "я", "ем", "ю", "ю")
            Rule 203
            manRule2 = True
            Exit Function
        Else
            osnova = workingWord
            If Left(Right(osnova, 2), 1) = "і" Then
                osnova = Left(osnova, Len(osnova) - 2) & "о" & Right(osnova, 1)
            ElseIf Last21 = "е" And InStr(vowels & "с", Last41) > 0 And Not (InStr("ланер", StrConv(workingWord, vbLowerCase)) > 0) Then
                ' Випадання букви е при відмінюванні слів типу Вітер
                delim = InStrRev(osnova, "е")
                osnova = Left(osnova, delim - 1) & Mid(osnova, delim + 1, Len(osnova) - delim)
            End If
            wordForms osnova, Array("а", "у", "а", "ом", "у", "е")
            Rule 204
            manRule2 = True
            Exit Function
        End If
    End If
    manRule2 = False
End Function
Public Function manRule22()
Dim osnova, Last1, Last2, Last3, Last4, Last21, Last41, delim
    Last1 = Last(1)
    Last2 = Last(2)
    Last3 = Last(3)
    Last4 = Last(4)
    Last21 = Last(2, 1)
    Last41 = Last(4, 1)
    If Last1 = "р" Then
'        If (Last2 = "ар" Or Last2 = "ир") And Last3 <> "вар" And Not (InStr("гайдар гончар володимир назар", StrConv(workingWord, vbLowerCase)) > 0) Then
'            wordForms workingWord, Array("я", "ю", "я", "ем", "ю", "ю")
'            Rule 201
'            manRule22 = True
'            Exit Function
'        ElseIf Last2 = "яр" Then 'Or Last3 = "нір" Then
'            If Last3 = "ляр" Then
'                wordForms workingWord, Array("а", "у", "а", "ем", "у", "е")
'            Else
'                wordForms workingWord, Array("а", "у", "а", "ем", "у", "е")
'            End If
'            Rule 202
'            manRule22 = True
'            Exit Function
'        Else
        If Last3 = "нір" Or InStr("Авенір Олефір Лігер", StrConv(workingWord, vbProperCase)) > 0 Then
            wordForms workingWord, Array("а", "у", "а", "ом", "у", "е")
            Rule 202
            manRule22 = True
            Exit Function
        ElseIf InStr("ігор лазар", StrConv(workingWord, vbLowerCase)) > 0 Then
            wordForms workingWord, Array("я", "ю", "я", "ем", "ю", "ю")
            Rule 203
            manRule22 = True
            Exit Function
        Else
            osnova = workingWord
            If Left(Right(osnova, 2), 1) = "і" Then
                osnova = Left(osnova, Len(osnova) - 2) & "о" & Right(osnova, 1)
            ElseIf Last21 = "е" And InStr(vowels & "с", Last41) > 0 Then
                ' Випадання букви е при відмінюванні слів типу Вітер
                delim = InStrRev(osnova, "е")
                osnova = Left(osnova, delim - 1) & Mid(osnova, delim + 1, Len(osnova) - delim)
            End If
            wordForms osnova, Array("а", "у", "а", "ом", "у", "е")
            Rule 204
            manRule22 = True
            Exit Function
        End If
    End If
    manRule22 = False
End Function

'     * Українські чоловічі імена, що в називному відмінку однини закінчуються на приголосний та -о,
'     * відмінюються як відповідні іменники ІІ відміни.
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function manRule3()
Dim group, osnova, BeforeLast, osLast, delim, Last1, Last2, Last3, Last31, Last41, Last42, Last51, osnovaLast21
    BeforeLast = Last(2, 1)
    Last1 = Last(1)
    Last2 = Last(2)
    Last3 = Last(3)
    Last31 = Last(3, 1)
    Last41 = Last(4, 1)
    Last42 = Last(4, 2)
    Last51 = Last(5, 1)
    If InStr(consonant & "оь", Last1) > 0 Then
        group = detect2Group(workingWord)
        osnova = getOsnova(workingWord)
        osLast = Right(osnova, 1)
        If osLast <> "й" And Left(Right(osnova, 2), 1) = "і" And Not InStr("світ цвіт уліш кіль міль хліб онід йців уців відь дліх", Right(StrConv(workingWord, vbLowerCase), 4)) > 0 _
        And StrComp(workingWord, "Гліб", vbTextCompare) <> 0 And StrComp(workingWord, "Чіп", vbTextCompare) <> 0 And Not InStr("ік іч із", Last2) <> 0 Then
            If InStr("Корінь Кремінь Кисіль", StrConv(workingWord, vbProperCase)) > 0 Or Last3 = "ідь" And InStr(vowels, Last51) > 0 Then
                osnova = Left(osnova, Len(osnova) - 2) & "е" & Right(osnova, 1)
            Else
                osnova = Left(osnova, Len(osnova) - 2) & "о" & Right(osnova, 1)
            End If
        End If
        ' Випадання букви е при відмінюванні слів типу Орел                                   & "гк"
        osnovaLast21 = Left(Right(osnova, 2), 1)
'        If StrComp(Left(osnova, 1), "о", vbTextCompare) = 0 And FirstLastVowel(osnova, vowels & "гк") = "е" And Last2 <> "сь" And Last1 <> "г" Then
        If StrComp(Left(osnova, 1), "о", vbTextCompare) = 0 And osnovaLast21 = "е" And Last2 <> "сь" And Last1 <> "г" Or Right(workingWord, 4) = "вель" Then
            delim = InStrRev(osnova, "е")
            osnova = Left(osnova, delim - 1) & Mid(osnova, delim + 1, Len(osnova) - delim)
        End If
        If group = 1 Then
            If Last2 = "ок" And Last3 <> "лок" Then
'                wordForms workingWord, Array("ка", "ку", "ка", "ком", "кові", "че"), 2
                wordForms workingWord, Array("ка", "ку", "ка", "ком", "ку", "ку"), 2
                Rule 301
                manRule3 = True
                Exit Function
            ElseIf InStr("ов ев єв ів їв ин ін їн", Last2) > 0 And Not InStr("лев остромов", StrConv(workingWord, vbLowerCase)) > 0 Then
                osnova = workingWord
'                If Last2 = "ів" And Not InStr("райців луців", StrConv(workingWord, vbLowerCase)) > 0 Then
                If Last2 = "ів" And Last31 <> "ц" Then                                                                          ' And InStr(vowels & "ь", Last41) > 0
                    If Right(workingWord, 3) <> "лів" And (Not InStr(shyplyachi, Last31) > 0) And InStr(tconsonant, Last31) > 0 Or InStr("дж дз", Last42) > 0 And InStr(vowels, Last51) > 0 Then
                        osnova = Left(workingWord, Len(workingWord) - 2) & "ов"
                    ElseIf InStr(mconsonant, Last31) > 0 Or InStr(shyplyachi, Last31) > 0 Then
                        osnova = Left(workingWord, Len(workingWord) - 2) & "ев"
                    End If
                End If
                ' ashu
                If Last2 = "ин" Then
                    If InStr("Литвин Турчин Угрин Русин Сербин Волошин", StrConv(workingWord, vbProperCase)) > 0 Then
                        wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "ом", osLast & "і", inverse2(osLast) & "е"), 1
                    Else
                        wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "им", osLast & "і", inverse2(osLast) & "е"), 1
                    End If
                ElseIf Last2 = "ів" Or Last2 = "єв" Or Last2 = "їн" Then
                    wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "им", osLast & "і", inverse2(osLast) & "е"), 1
                Else
                    wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "им", osLast & "у", inverse2(osLast) & "е"), 1
                End If
                Rule 302
                manRule3 = True
                Exit Function
            ElseIf Last2 = "ін" Then
                wordForms workingWord, Array("а", "у", "а", "им", "у", "е")
                Rule 303
                manRule3 = True
                Exit Function
            ElseIf StrConv(workingWord, vbProperCase) = "Лев" Then
                wordForms workingWord, Array("ьва", "ьву", "ьва", "ьвом", "ьвові", "еве"), 2
                Rule 303
                manRule3 = True
                Exit Function
            Else
'                wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "ом", osLast & "ові", inverse2(osLast) & "е"), 1
                wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "ом", osLast & "у", osLast & "у"), 1
                Rule 304
                manRule3 = True
                Exit Function
            End If
        End If
        If group = 2 Then
            wordForms osnova, Array("а", "у", "а", "ем", "у", "е")
            Rule 305
            manRule3 = True
            Exit Function
        End If
        If group = 3 Then
            If Last2 = "ей" And InStr(gubni, Last31) Then
                osnova = Left(workingWord, Len(workingWord) - 2) & "'"
                wordForms osnova, Array("я", "ю", "я", "єм", "ю", "ю")
                Rule 306
                manRule3 = True
                Exit Function
            ElseIf Last1 = "й" Or BeforeLast = "і" Or StrConv(workingWord, vbLowerCase) = "швець" Or StrConv(workingWord, vbLowerCase) = "жнець" Then
                wordForms workingWord, Array("я", "ю", "я", "єм", "ю", "ю"), 1
                Rule 307
                manRule3 = True
                Exit Function
'            ElseIf StrConv(workingWord, vbLowerCase) = "швець" Then
'                wordForms workingWord, Array("евця", "евцеві", "евця", "евцем", "евцеві", "евцю"), 4
'                Rule 308
'                manRule3 = True
'                Exit Function
            ElseIf Last3 = "ець" And Not InStr("чернець мудрець перець", StrConv(workingWord, vbLowerCase)) > 0 And Len(workingWord) > 4 And InStr(vowels, Last51) > 0 Then
                If Last41 = "л" Then
                    wordForms workingWord, Array("ьця", "ьцю", "ьця", "ьцем", "ьцю", "ьцю"), 3
                Else
                    wordForms workingWord, Array("ця", "цю", "ця", "цем", "цю", "цю"), 3
                End If
                Rule 309
                manRule3 = True
                Exit Function
            ElseIf Last3 = "ень" And Not InStr("шершень сень", StrConv(workingWord, vbLowerCase)) > 0 Then
                If Last41 = "л" Then
                    wordForms workingWord, Array("ьня", "ьню", "ьня", "ьнем", "ьню", "ьню"), 3
                Else
                    wordForms workingWord, Array("ня", "ню", "ня", "нем", "ню", "ню"), 3
                End If
                Rule 310
                manRule3 = True
                Exit Function
            ElseIf InStr("єць яць", Last3) > 0 Then
                wordForms workingWord, Array("йця", "йцю", "йця", "йцем", "йцю", "йцю"), 3
                Rule 311
                manRule3 = True
                Exit Function
            ElseIf Last2 = "ьо" Then
                wordForms osnova, Array("я", "ю", "я", "ьом", "ї", "ю")
                Rule 312
                manRule3 = True
                Exit Function
            Else
                wordForms osnova, Array("я", "ю", "я", "ем", "ю", "ю")
                Rule 312
                manRule3 = True
                Exit Function
            End If
        End If
    End If
    manRule3 = False
End Function
Public Function manRule33()
Dim group, osnova, BeforeLast, osLast, delim, Last1, Last2, Last3, Last31, Last41, Last42, Last51, osnovaLast21
    BeforeLast = Last(2, 1)
    Last1 = Last(1)
    Last2 = Last(2)
    Last3 = Last(3)
    Last31 = Last(3, 1)
    Last41 = Last(4, 1)
    Last42 = Last(4, 2)
    Last51 = Last(5, 1)
    If InStr(consonant & "оь", Last1) > 0 Then
        group = detect2Group(workingWord)
        osnova = getOsnova(workingWord)
        osLast = Right(osnova, 1)
        If osLast <> "й" And Left(Right(osnova, 2), 1) = "і" And Not InStr("світ цвіт уліш кіль сіль міль хліб онід йців уців відь", Right(StrConv(workingWord, vbLowerCase), 4)) > 0 _
        And Not (InStr("Гліб Чіп Авенір Лаврін Олефір Кіріл Юсіф", StrConv(workingWord, vbProperCase)) > 0) And Not (InStr("ік іч", Last2) > 0) Then
            If StrComp(workingWord, "Корінь", vbTextCompare) = 0 Then
                osnova = Left(osnova, Len(osnova) - 2) & "е" & Right(osnova, 1)
            Else
                osnova = Left(osnova, Len(osnova) - 2) & "о" & Right(osnova, 1)
            End If
        End If
        ' Випадання букви е при відмінюванні слів типу Орел                                   & "гк"
        osnovaLast21 = Left(Right(osnova, 2), 1)
'        If StrComp(Left(osnova, 1), "о", vbTextCompare) = 0 And FirstLastVowel(osnova, vowels & "гк") = "е" And Last2 <> "сь" And Last1 <> "г" Then
        If StrComp(Left(osnova, 1), "о", vbTextCompare) = 0 And osnovaLast21 = "е" And Last2 <> "сь" And Last1 <> "г" Then
            delim = InStrRev(osnova, "е")
            osnova = Left(osnova, delim - 1) & Mid(osnova, delim + 1, Len(osnova) - delim)
        End If
        If group = 1 Then
            If Last2 = "ок" And Last3 <> "лок" Then
'                wordForms workingWord, Array("ка", "ку", "ка", "ком", "кові", "че"), 2
                wordForms workingWord, Array("ка", "ку", "ка", "ком", "ку", "ку"), 2
                Rule 301
                manRule33 = True
                Exit Function
'            ElseIf InStr("ов ев єв ів їв ин ін їн", Last2) > 0 And Not InStr("лев остромов", StrConv(workingWord, vbLowerCase)) > 0 Then
'                osnova = workingWord
''                If Last2 = "ів" And Not InStr("райців луців", StrConv(workingWord, vbLowerCase)) > 0 Then
'                If Last2 = "ів" And Last31 <> "ц" Then
'                    If Right(workingWord, 3) <> "лів" And (Not InStr(shyplyachi, Last31) > 0) And InStr(tconsonant, Last31) > 0 And InStr(vowels, Last41) > 0 Or InStr("дж дз", Last42) And InStr(vowels, Last51) > 0 Then
'                        osnova = Left(workingWord, Len(workingWord) - 2) & "ов"
'                    ElseIf InStr(mconsonant, Last31) > 0 Or InStr(shyplyachi, Last31) > 0 Then
'                        osnova = Left(workingWord, Len(workingWord) - 2) & "ев"
'                    End If
'                End If
'                ' ashu
'                If Last2 = "ин" Then
'                    If InStr("Литвин Турчин Угрин Русин Сербин Волошин", StrConv(workingWord, vbProperCase)) > 0 Then
'                        wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "ом", osLast & "і", inverse2(osLast) & "е"), 1
'                    Else
'                        wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "им", osLast & "і", inverse2(osLast) & "е"), 1
'                    End If
'                ElseIf Last2 = "ів" Or Last2 = "єв" Or Last2 = "їн" Then
'                    wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "им", osLast & "і", inverse2(osLast) & "е"), 1
'                Else
'                    wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "им", osLast & "у", inverse2(osLast) & "е"), 1
'                End If
'                Rule 302
'                manRule33 = True
'                Exit Function
''            ElseIf Last2 = "ін" Then
''                wordForms workingWord, Array("а", "у", "а", "им", "у", "е")
''                Rule 303
''                manRule33 = True
''                Exit Function
            ElseIf StrConv(workingWord, vbProperCase) = "Лев" Then
                wordForms workingWord, Array("ьва", "ьву", "ьва", "ьвом", "ьвові", "еве"), 2
                Rule 303
                manRule33 = True
                Exit Function
            Else
                If StrComp(osnova, Left(workingWord, Len(osnova))) = 0 Then
                    wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "ом", osLast & "у", inverse2(osLast) & "е"), 1
                Else
                    wordForms osnova, Array(osLast & "а", osLast & "у", osLast & "а", osLast & "ом", osLast & "у", osLast & "у"), 1
                End If
                Rule 304
                manRule33 = True
                Exit Function
            End If
        End If
        If group = 2 Then
            wordForms osnova, Array("а", "у", "а", "ем", "у", "е")
            Rule 305
            manRule33 = True
            Exit Function
        End If
        If group = 3 Then
            If Last2 = "ей" And InStr(gubni, Last31) Then
                osnova = Left(workingWord, Len(workingWord) - 2) & "'"
                wordForms osnova, Array("я", "ю", "я", "єм", "ю", "ю")
                Rule 306
                manRule33 = True
                Exit Function
            ElseIf Last1 = "й" Or BeforeLast = "і" Or StrConv(workingWord, vbLowerCase) = "швець" Or StrConv(workingWord, vbLowerCase) = "жнець" Then
                wordForms workingWord, Array("я", "ю", "я", "єм", "ю", "ю"), 1
                Rule 307
                manRule33 = True
                Exit Function
'            ElseIf StrConv(workingWord, vbLowerCase) = "швець" Then
'                wordForms workingWord, Array("евця", "евцеві", "евця", "евцем", "евцеві", "евцю"), 4
'                Rule 308
'                manRule3 = True
'                Exit Function
            ElseIf Last3 = "ець" And Not InStr("чернець мудрець перець", StrConv(workingWord, vbLowerCase)) > 0 Then
                If Last41 = "л" Then
                    wordForms workingWord, Array("ьця", "ьцю", "ьця", "ьцем", "ьцю", "ьцю"), 3
                Else
                    wordForms workingWord, Array("ця", "цю", "ця", "цем", "цю", "цю"), 3
                End If
                Rule 309
                manRule33 = True
                Exit Function
            ElseIf Last3 = "ень" And Not InStr("шершень", StrConv(workingWord, vbLowerCase)) > 0 Then
                If Last41 = "л" Then
                    wordForms workingWord, Array("ьня", "ьню", "ьня", "ьнем", "ьню", "ьню"), 3
                Else
                    wordForms workingWord, Array("ня", "ню", "ня", "нем", "ню", "ню"), 3
                End If
                Rule 310
                manRule33 = True
                Exit Function
            ElseIf InStr("єць яць", Last3) > 0 Then
                wordForms workingWord, Array("йця", "йцю", "йця", "йцем", "йцю", "йцю"), 3
                Rule 311
                manRule33 = True
                Exit Function
            Else
                wordForms osnova, Array("я", "ю", "я", "ем", "ю", "ю")
                Rule 312
                manRule33 = True
                Exit Function
            End If
        End If
    End If
    manRule33 = False
End Function

'     * Якщо слово закінчується на і, то відмінюємо як множину
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function manRule4()
Dim Last1
    Last1 = Last(1)
    manRule4 = False
    If Last1 = "і" Then
'        wordForms workingWord, Array("их", "им", "их", "ими", "их", "і"), 1
        wordForms workingWord, Array("і", "і", "і", "і", "і", "і"), 1
        Rule 4
        manRule4 = True
    End If
End Function

'     * Якщо слово закінчується на ий або ой
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function manRule5()
Dim Last2
    Last2 = Last(2)
    manRule5 = False
    If InStr("ий ой", Last2) > 0 Then
        wordForms workingWord, Array("ого", "ому", "ого", "им", "ому", "ий"), 2
        Rule 51
        manRule5 = True
    ElseIf Right(workingWord, 3) = "ній" Then
        wordForms workingWord, Array("ього", "ьому", "ього", "ім", "ьому", "ій"), 2
        Rule 52
        manRule5 = True
        Exit Function
    ElseIf Right(workingWord, 1) = "і" Or Right(workingWord, 3) = "дзе" Or Last2 = "их" Or InStr("акі ово аго", Right(workingWord, 3)) > 0 Then
        wordForms workingWord, Array("", "", "", "", "", "")
        Rule 53
        manRule5 = True
        Exit Function
    ElseIf Right(workingWord, 1) = "а" And InStr(vowels, Left(Right(workingWord, 2), 1)) > 0 Then
        wordForms workingWord, Array("а", "а", "а", "а", "а", "а"), 1
        Rule 56
        manRule5 = True
        Exit Function
    ElseIf InStr("віардо він’ї віньї готьє дідро картьє мане міро руссо шагане дюма золя", StrConv(workingWord, vbLowerCase)) > 0 Then
        wordForms workingWord, Array("", "", "", "", "", "")
        Rule 57
        manRule5 = True
        Exit Function
    End If
End Function

'     * Українські чоловічі та жіночі імена, що в називному відмінку однини закінчуються на -а (-я),
'     * відмінються як відповідні іменники І відміни.
'     * - Примітка 1. Кінцеві приголосні основи г, к, х у жіночих іменах
'     *   у давальному та місцевому відмінках однини перед закінченням -і
'     *   змінюються на з, ц, с: Ольга - Ользі, Палажка - Палажці, Солоха - Солосі.
'     * - Примітка 2. У жіночих іменах типу Одарка, Параска в родовому відмінку множини
'     *   в кінці основи між приголосними з'являється звук о: Одарок, Парасок
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function womanRule1()
Dim BeforeLast, osnova, Last4, Last1
    Last1 = Last(1)
    Last4 = Last(4)
    BeforeLast = Last(2, 1)
    If Last4 = "ніга" Then
        osnova = Left(workingWord, Len(workingWord) - 3) & "о"
        wordForms osnova, Array("ги", "зі", "гу", "гою", "зі", "го")
        Rule 101
        womanRule1 = True
        Exit Function
    ElseIf Last1 = "а" Then
        wordForms workingWord, Array(BeforeLast & "и", inverseGKH(BeforeLast) & "і", BeforeLast & "у", BeforeLast & "ою", inverseGKH(BeforeLast) & "і", BeforeLast & "о"), 2
        Rule 102
        womanRule1 = True
        Exit Function
    ElseIf Last1 = "я" Then
        If InStr(vowels, BeforeLast) > 0 Or isApostrof(BeforeLast) Then
            wordForms workingWord, Array("ї", "ї", "ю", "єю", "ї", "є"), 1
            Rule 103
            womanRule1 = True
            Exit Function
        Else
'            wordForms workingWord, Array(BeforeLast & "і", inverseGKH(BeforeLast) & "і", BeforeLast & "ю", BeforeLast & "ею", inverseGKH(BeforeLast) & "і", BeforeLast & "е"), 2
            wordForms workingWord, Array(BeforeLast & "і", inverseGKH(BeforeLast) & "і", BeforeLast & "ю", BeforeLast & "ею", inverseGKH(BeforeLast) & "і", BeforeLast & "я"), 2
            Rule 104
            womanRule1 = True
            Exit Function
        End If
    End If
    womanRule1 = False
End Function
Public Function womanRule11()
Dim BeforeLast, osnova, Last4, Last1
    Last1 = Last(1)
    Last4 = Last(4)
    BeforeLast = Last(2, 1)
    If Last4 = "ніга" Then
        osnova = Left(workingWord, Len(workingWord) - 3) & "о"
        wordForms osnova, Array("ги", "зі", "гу", "гою", "зі", "го")
        Rule 101
        womanRule11 = True
        Exit Function
    ElseIf Last1 = "а" Then
        wordForms workingWord, Array(BeforeLast & "и", inverseGKH(BeforeLast) & "і", BeforeLast & "у", BeforeLast & "ою", inverseGKH(BeforeLast) & "і", BeforeLast & "о"), 2
        Rule 102
        womanRule11 = True
        Exit Function
    ElseIf Last1 = "я" Then
        If InStr(vowels, BeforeLast) > 0 Or isApostrof(BeforeLast) Then
            wordForms workingWord, Array("ї", "ї", "ю", "єю", "ї", "є"), 1
            Rule 103
            womanRule11 = True
            Exit Function
        Else
            wordForms workingWord, Array(BeforeLast & "і", inverseGKH(BeforeLast) & "і", BeforeLast & "ю", BeforeLast & "ею", inverseGKH(BeforeLast) & "і", BeforeLast & "е"), 2
'            wordForms workingWord, Array(BeforeLast & "і", inverseGKH(BeforeLast) & "і", BeforeLast & "ю", BeforeLast & "ею", inverseGKH(BeforeLast) & "і", BeforeLast & "я"), 2
            Rule 104
            womanRule11 = True
            Exit Function
        End If
    End If
    womanRule11 = False
End Function

'     * Українські жіночі імена, що в називному відмінку однини закінчуються на приголосний,
'     * відмінюються як відповідні іменники ІІІ відміни
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function womanRule2()
Dim duplicate, osnova, apostrof, osLast, osBeforeLast, Last1
    Last1 = Last(1)
    If InStr(consonant & "ь", Last1) > 0 Then
        osnova = getOsnova(workingWord)
        apostrof = ""
        duplicate = ""
        osLast = Right(osnova, 1)
        osBeforeLast = Left(Right(osnova, 2), 1)
        ' чи треба ставити апостроф
        If InStr("мвпбф", osLast) > 0 And InStr(vowels, osBeforeLast) > 0 Then
            apostrof = "'"
        End If
        ' чи треба подвоювати
        If InStr("дтзсцлн", osLast) > 0 Then
            duplicate = osLast
        End If
        If Last1 = "ь" Then
            wordForms osnova, Array("і", "і", "ь", duplicate & apostrof & "ю", "і", "е")
            Rule 201
            womanRule2 = True
            Exit Function
        Else
            wordForms osnova, Array("і", "і", "", duplicate & apostrof & "ю", "і", "е")
            Rule 202
            womanRule2 = True
            Exit Function
        End If
    End If
    womanRule2 = False
End Function
     
'     * Якщо слово на ськ або це російське прізвище
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function womanRule3()
Dim BeforeLast, Last2, Last1, Last21, Last32
    Last21 = Last(2, 1)
    Last32 = Last(3, 2)
    Last1 = Last(1)
    Last2 = Last(2)
    BeforeLast = Last(2, 1)
'    If Right(workingWord, 4) = "ишин" Then
'        wordForms workingWord, Array("ої", "ій", "у", "ою", "ій", "")
'        Rule 300
'        womanRule3 = True
'        Exit Function
'    Else
    If Last2 = "ая" Then
        wordForms workingWord, Array("ої", "ій", "ую", "ою", "ій", "ая"), 2
        Rule 301
        womanRule3 = True
        Exit Function
    ElseIf Last1 = "а" And (InStr("чнвт", Last21) > 0 Or Last32 = "ьк" Or InStr("Мудра", StrConv(workingWord, vbProperCase)) > 0) And Not InStr("Щербина", StrConv(workingWord, vbProperCase)) > 0 Then
'        wordForms workingWord, Array(BeforeLast & "ої", BeforeLast & "ій", BeforeLast & "у", BeforeLast & "ою", BeforeLast & "ій", BeforeLast & "о"), 2
        wordForms workingWord, Array(BeforeLast & "ої", BeforeLast & "ій", BeforeLast & "у", BeforeLast & "ою", BeforeLast & "ій", BeforeLast & "а"), 2
        Rule 302
        womanRule3 = True
        Exit Function
    ElseIf Right(workingWord, 1) = "а" And InStr(vowels, Left(Right(workingWord, 2), 1)) > 0 Then
        wordForms workingWord, Array("а", "а", "а", "а", "а", "а"), 1
        Rule 303
        womanRule3 = True
        Exit Function
    ElseIf InStr("віардо він’ї віньї готьє дідро картьє мане міро руссо шагане дюма золя", StrConv(workingWord, vbLowerCase)) > 0 Then
        wordForms workingWord, Array("", "", "", "", "", "")
        Rule 304
        womanRule3 = True
        Exit Function
    End If
    womanRule3 = False
End Function

'     * Функція намагається застосувати ланцюг правил для чоловічих імен
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function manFirstName()
    manFirstName = False
    If manRule1 Then
        manFirstName = True
    Else
        If manRule22 Then
            manFirstName = True
        Else
            If manRule33 Then manFirstName = True
        End If
    End If
'    manFirstName = RulesChain("man", Array(1, 2, 3))
End Function
     
'     * Функція намагається застосувати ланцюг правил для жіночих імен
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function womanFirstName()
    womanFirstName = False
    If womanRule11 Then
        womanFirstName = True
    Else
        If womanRule2 Then womanFirstName = True
    End If
'    womanFirstName = RulesChain("woman", Array(1, 2))
End Function
     
'     * Функція намагається застосувати ланцюг правил для чоловічих прізвищ
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function manSecondName()
    manSecondName = False
    If manRule5 Then
        manSecondName = True
    Else
        If manRule1 Then
            manSecondName = True
        Else
            If manRule2 Then
                manSecondName = True
            Else
                If manRule3 Then
                    manSecondName = True
                Else
                    If manRule4 Then manSecondName = True
                End If
            End If
        End If
    End If
'    manSecondName = RulesChain("man", Array(5, 1, 2, 3, 4))
End Function
     
'     * Функція намагається застосувати ланцюг правил для жіночих прізвищ
'     * @return boolean true - якщо було задіяно правило з переліку, false - якщо правило не знайдено
Public Function womanSecondName()
    womanSecondName = False
    If womanRule3 Then
        womanSecondName = True
    Else
        If womanRule1 Then womanSecondName = True
    End If
'    womanSecondName = RulesChain("woman", Array(3, 1))
End Function
     
'     * Фунція відмінює чоловічі по-батькові
'     * @return boolean true - якщо слово успішно змінене, false - якщо невдалося провідміняти слово
Public Function manFatherName()
Dim Last2
    Last2 = Last(2)
    manFatherName = False
    If InStr("ич іч", Last2) > 0 Then
        wordForms workingWord, Array("а", "у", "а", "ем", "у", "у")
        manFatherName = True
    End If
End Function

'     * Фунція відмінює жіночі по-батькові
'     * @return boolean true - якщо слово успішно змінене, false - якщо невдалося провідміняти слово
Public Function womanFatherName()
Dim Last3
    Last3 = Last(3)
    womanFatherName = False
    If Last3 = "вна" Then
        wordForms workingWord, Array("и", "і", "у", "ою", "і", "о"), 1
        womanFatherName = True
    End If
End Function

'     * Визначення статі, за правилами імені
'     * @param NCLNameCaseWord $word об’єкт класу зі словом, для якого необхідно визначити стать
Public Sub GenderByFirstName(Word)
Dim man, woman, Last1, Last2, Last3
    setWorkingWord Word.getWord()
    Last3 = Last(3)
    Last2 = Last(2)
    Last1 = Last(1)
    man = 0
    woman = 0
    If Last1 = "й" Then man = man + 0.9
    If InStr("петро микола", workingWord) > 0 Then man = man + 30
    If InStr("он ов ав ам ол ан рд мп ко ло", Last2) > 0 Then man = man + 0.5
    If InStr("бов нка яра ила опа", Last3) > 0 Then woman = woman + 0.5
    If InStr(consonant, Last1) > 0 Then man = man + 0.01
    If Last1 = "ь" Then man = man + 0.02
    If Last2 = "дь" Then woman = woman + 0.1
    If InStr("ель бов", Last3) > 0 Then woman = woman + 0.4
    Word.setGender man, woman
End Sub

'     * Визначення статі, за правилами прізвища
'     * @param NCLNameCaseWord $word об’єкт класу зі словом, для якого необхідно визначити стать
Public Sub GenderBySecondName(Word)
Dim man, woman, Last2, Last3
    setWorkingWord Word.getWord()
    Last3 = Last(3)
    Last2 = Last(2)
    man = 0
    woman = 0
    If InStr("ов ин ев єв ін їн ий їв ів ой ей", Last2) > 0 Then man = man + 0.4
    If InStr("ова ина ева єва іна мін", Last3) > 0 Then woman = woman + 0.4
    If Last2 = "ая" Then woman = woman + 0.4
    Word.setGender man, woman
End Sub

'     * Визначення статі, за правилами по-батькові
'     * @param NCLNameCaseWord $word об’єкт класу зі словом, для якого необхідно визначити стать
Public Sub GenderByFatherName(Word)
Dim Last2
    setWorkingWord Word.getWord()
    Last2 = Last(2)
    If Last2 = "ич" Then Word.setGender 10, 0
    If Last2 = "на" Then Word.setGender 0, 12
End Sub
     
'     * Ідентифікує слово визначаючи чи це ім’я, чи це прізвище, чи це побатькові
'     * - <b>N</b> - ім’я
'     * - <b>S</b> - прізвище
'     * - <b>F</b> - по-батькові
'     * @param NCLNameCaseWord $word об’єкт класу зі словом, яке необхідно ідентифікувати
Public Sub detectNamePart(Word)
Dim namepart, first, second, father, max, Last3, Last4, Last2, Last1
    namepart = Word.getWord()
    setWorkingWord namepart
    Last1 = Last(1)
    Last2 = Last(2)
    Last4 = Last(4)
    Last3 = Last(3)
    first = 0
    second = 0
    father = 0
    If InStr("вна чна ліч", Last3) > 0 Or InStr("ьмич ович", Last4) > 0 Then father = father + 3
    If Last3 = "тин" Or InStr("ьмич юбов івна явка орив кіян", Last4) > 0 Then first = first + 0.5
    If InStr("лев гаїна афіна антоніна ангеліна альвіна альбіна аліна павло олесь микола мая англеліна елькін мерлін", namepart) > 0 Then first = first + 10
    If InStr("ов ін ев єв ий ин ой ко ук як ца их ик ун ок ша ая га єк аш ив юк ус це ак бр яр іл ів ич сь ей нс яс ер ай ян ах ць ющ іс ач уб ох юх ут ча ул вк зь уц їн де уз юр ік іч ро", Last2) > 0 Then second = second + 0.4
    If InStr("ова ева єва тих рик вач аха шен мей арь вка шир бан чий іна їна ька ань ива аль ура ран ало ола кур оба оль нта зій ґан іло шта юпа рна бла еїн има мар кар оха чур ниш ета тна зур нір йма орж рба іла лас дід роз аба чан ган", Last3) > 0 Then second = second + 0.4
    If InStr("ьник нчук тник кирь ский шена шина вина нина гана гана хній зюба орош орон сило руба лест мара обка рока сика одна нчар вата ндар грій", Last4) > 0 Then second = second + 0.4
    If Last1 = "і" Then second = second + 0.2
    max = IIf(first > second, first, second)
    max = IIf(max > father, max, father)
    If first = max Then
        Word.setNamePart "N"
    ElseIf second = max Then
        Word.setNamePart "S"
    Else
        Word.setNamePart "F"
    End If
End Sub

Private Sub Class_Initialize()
    words = Array()
End Sub


