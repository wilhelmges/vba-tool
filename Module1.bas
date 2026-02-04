Attribute VB_Name = "Module1"
Option Explicit
Const git = 24
Const чоловік = 1
Const жінка = 2
Const Називний = 0
Const Родовий = 1
Const Давальний = 2
Const Знахідний = 3
Const Орудний = 4
Const Місцевий = 5
Const Кличний = 6
Public Const wsFile = "Штат*"
Public Const wsShtat = "ШТАТ 01"
Public Const wsFileVtraty = "втрати *"
Public Const wsVtraty = "втрати"
Public Const rozpTempl = "У розп*"
Public Const wsFileVac = "Відпустки*"
Public Const wsVac = "Відпустки"
Public Const wsCom = "Відрядження"
Public Const wsFileBank = "BANK*"
Public Const wsBank = "Основний ПРИВАТ"
Public wbfound, vtratyfound, vacfound, bankfound, checkmode
Dim Cache As Cache


Function СумаПрописом(ByVal number As Double) As String
    Dim гривні As Long, копійки As Long
    Dim Результат As String

    ' Розділяємо число на гривні та копійки
    гривні = Int(number)
    копійки = Round((number - гривні) * 100, 0)

    ' Формуємо суму прописом
    If гривні = 0 Then
        Результат = "нуль гривень"
    Else
        Результат = ЧислоПрописом(гривні, True) & " грив" & Окончання(гривні, "ня", "ні", "ень")
    End If

    Результат = Результат & " " & format(копійки, "00") & " копійок"
    СумаПрописом = trim(Результат)
End Function


Private Function ЧислоПрописом(ByVal n As Long, Optional жіночийРід As Boolean = False) As String
    Static Одиниці, ОдиниціЖ, Десятки, Сотні, Особливі, Розряди
    If IsEmpty(Одиниці) Then
        Одиниці = Array("", "один", "два", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять")
        ОдиниціЖ = Array("", "одна", "дві", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять")
        Десятки = Array("", "десять", "двадцять", "тридцять", "сорок", "п’ятдесят", "шістдесят", "сімдесят", "вісімдесят", "дев’яносто")
        Сотні = Array("", "сто", "двісті", "триста", "чотириста", "п’ятсот", "шістсот", "сімсот", "вісімсот", "дев’ятсот")
        Особливі = Array("десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", _
                         "п’ятнадцять", "шістнадцять", "сімнадцять", "вісімнадцять", "дев’ятнадцять")
        Розряди = Array("", "тисяч", "мільйон", "мільярд")
    End If

    Dim частини(3) As String
    Dim група, степінь As Long
    Dim частина As Long, Результат As String

    ' Обробка груп по 3 цифри (тисячі, мільйони, мільярди)
    Do While n > 0
        частина = n Mod 1000
        If частина > 0 Then
            Dim слова As String
            слова = ТрицифровеЧислоПрописом(частина, (степінь = 1))
            Select Case степінь
                Case 1 ' тисячі
                    слова = слова & " тисяч" & Окончання(частина, "а", "і", "")
                Case 2 ' мільйони
                    слова = слова & " мільйон" & Окончання(частина, "", "и", "ів")
                Case 3 ' мільярди
                    слова = слова & " мільярд" & Окончання(частина, "", "и", "ів")
            End Select
            частини(степінь) = слова
        End If
        n = n \ 1000
        степінь = степінь + 1
    Loop

    Dim i As Long
    For i = UBound(частини) To 0 Step -1
        If частини(i) <> "" Then Результат = Результат & частини(i) & " "
    Next i

    ЧислоПрописом = trim(Результат)
End Function


Private Function ТрицифровеЧислоПрописом(ByVal n As Long, Optional жіночийРід As Boolean = False) As String
    Static Одиниці, ОдиниціЖ, Десятки, Сотні, Особливі
    If IsEmpty(Одиниці) Then
        Одиниці = Array("", "один", "два", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять")
        ОдиниціЖ = Array("", "одна", "дві", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять")
        Десятки = Array("", "десять", "двадцять", "тридцять", "сорок", "п’ятдесят", "шістдесят", "сімдесят", "вісімдесят", "дев’яносто")
        Сотні = Array("", "сто", "двісті", "триста", "чотириста", "п’ятсот", "шістсот", "сімсот", "вісімсот", "дев’ятсот")
        Особливі = Array("десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", _
                         "п’ятнадцять", "шістнадцять", "сімнадцять", "вісімнадцять", "дев’ятнадцять")
    End If

    Dim Результат As String
    Dim Сот, Дес, Од As Long
    Сот = (n \ 100) Mod 10
    Дес = (n \ 10) Mod 10
    Од = n Mod 10

    If Сот > 0 Then Результат = Сотні(Сот) & " "
    If Дес = 1 Then
        Результат = Результат & Особливі(Од)
    Else
        If Дес > 0 Then Результат = Результат & Десятки(Дес) & " "
        If Од > 0 Then
            If жіночийРід Then
                Результат = Результат & ОдиниціЖ(Од)
            Else
                Результат = Результат & Одиниці(Од)
            End If
        End If
    End If

    ТрицифровеЧислоПрописом = trim(Результат)
End Function


Private Function Окончання(ByVal n As Long, форма1 As String, форма2 As String, форма3 As String) As String
    Dim ост As Long
    ост = n Mod 100
    If ост >= 11 And ост <= 19 Then
        Окончання = форма3
    Else
        ост = n Mod 10
        Select Case ост
            Case 1: Окончання = форма1
            Case 2, 3, 4: Окончання = форма2
            Case Else: Окончання = форма3
        End Select
    End If
End Function


' робота з ПІБ
' Створено: Окланд, 110 ОМБр, травень 2023
Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

Dim i As Long, j As Long, string1_length As Long, string2_length As Long
Dim distance(0 To 60, 0 To 50) As Long, smStr1(1 To 60) As Long, smStr2(1 To 50) As Long
Dim min1 As Long, min2 As Long, min3 As Long, minmin As Long, MaxL As Long

string1 = Left(string1, 60)
string2 = Left(string2, 60)

string1_length = Len(string1):  string2_length = Len(string2)

distance(0, 0) = 0
For i = 1 To string1_length:    distance(i, 0) = i: smStr1(i) = Asc(LCase(Mid$(string1, i, 1))): Next
For j = 1 To string2_length:    distance(0, j) = j: smStr2(j) = Asc(LCase(Mid$(string2, j, 1))): Next
For i = 1 To string1_length
    For j = 1 To string2_length
        If smStr1(i) = smStr2(j) Then
            distance(i, j) = distance(i - 1, j - 1)
        Else
            min1 = distance(i - 1, j) + 1
            min2 = distance(i, j - 1) + 1
            min3 = distance(i - 1, j - 1) + 1
            If min2 < min1 Then
                If min2 < min3 Then minmin = min2 Else minmin = min3
            Else
                If min1 < min3 Then minmin = min1 Else minmin = min3
            End If
            distance(i, j) = minmin
        End If
    Next
Next

' Levenshtein will properly return a percent match (100%=exact) based on similarities and Lengths etc...
MaxL = string1_length: If string2_length > MaxL Then MaxL = string2_length
Levenshtein = 100 - CLng((distance(string1_length, string2_length) * 100) / MaxL)

End Function

Function getLike(ByVal string1 As String) As String
Dim tempstr, rezstr, pos
rezstr = ""
pos = 1
For Each tempstr In Split(string1, " ")
    If pos = 1 And Len(tempstr) > 4 Then
        rezstr = rezstr & Left(tempstr, Len(tempstr) - 3) & "* "
    Else
        rezstr = rezstr & Left(tempstr, Len(tempstr) - 1) & "* "
    End If
    pos = pos + 1
Next
getLike = Left(rezstr, Len(rezstr) - 1)
End Function

' тестова процедура для роботи з відмінками
Sub vidmtest()
Dim gender
Dim nc As NCLNameCaseUa, ncU As NCLNameCaseUa_U, inRange As Range, outRange As Range, i, rez, rezU, pp
Dim PIB, sPIB, PIDR, zv, IPN, pos, tPIB, gPIB, tPIDR, tZV, tIPN, tPOS, tisSHT, RVK, tRVK, tPhone, Phone, tAddr, Addr, Lsht, tLsht, Lshtpp, Lsht3, tLsht3, Lshtpp3, tPochatok
Dim PIBsplit, rezsplit, rez4, tSex

    Set nc = New NCLNameCaseUa
    Set ncU = New NCLNameCaseUa_U
    Set inRange = ActiveWorkbook.Worksheets("Дані ПІБ,№,дати").Range("B4")
'    Set outRange = ActiveWorkbook.Worksheets("Відмінки").Range("B4:B10")
'    gender = -1
'    nc1.fullReset
'    nc1.splitFullName " Головій Олег Володимирович"
'    If gender > -1 Then nc1.setGender gender
'    nc.resultString = nc.getFormatted(1, "S N F")
'    Debug.Print nc.q("Шумак Олексій Анатолійович", 0, -1)
'    Debug.Print "we here"

    PIB = StrConv(Application.trim(inRange.Value), vbProperCase)
    inRange.Value = PIB

    OptimizeVBA True

      
    If Len(PIB) > 1 Then
            rez = nc.q(inRange.Value)
            rezU = ncU.q(inRange.Value)
        'For i = 1 To 7
            If rez(2) <> rezU(2) Then
                ActiveWorkbook.Worksheets("Дані ПІБ,№,дати").Range("B5") = rezU(2)
            Else
                ActiveWorkbook.Worksheets("Дані ПІБ,№,дати").Range("B5") = rez(2)
            End If
            
            If rez(3) <> rezU(3) Then
                ActiveWorkbook.Worksheets("Дані ПІБ,№,дати").Range("B6") = rezU(3)
            Else
                ActiveWorkbook.Worksheets("Дані ПІБ,№,дати").Range("B6") = rez(3)
            End If
            
        'Next
    End If
        
    OptimizeVBA False
End Sub


' робимо швидко
Sub OptimizeVBA(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    'Application.ScreenUpdating = Not (isOn)
    If isOn Then Application.CalculateFullRebuild
End Sub

Sub ШПС3()
    ШПС3make (True)
End Sub

Sub ШПС3blank()
    ШПС3make (False)
End Sub


' процедура формування ШПС, розділ 3
Sub ШПС3make(fill As Boolean)
    Dim nWorksheet As Worksheet, cWorksheet As Worksheet, lWorksheet As Worksheet
    Dim Interval As Date
    Dim n, cc, cLine, nn, PIB
    Dim parRange As Range, Line As Range
    
    check_workbook
    If wbfound = vbNullString Or wbfound = "cache" Then
        MsgBox ("Ви забули відкрити файл 'Штат'")
        Exit Sub
    End If
    
    OptimizeVBA True
    
    ' нова вкладка
    For Each cWorksheet In ActiveWorkbook.Worksheets
        Set lWorksheet = cWorksheet
    Next
    
    Interval = Now()
    Set nWorksheet = ActiveWorkbook.Worksheets.Add(after:=lWorksheet)
    
    With nWorksheet
    
        ' формуємо шапку
        .name = "ШПС розділ 3 (" & format(Interval, "yyyymmddnnss") & ")"
        .Activate
        
        .Range("A:A").ColumnWidth = 9.14
        .Range("B:B").ColumnWidth = 66.57
        .Range("C:C").ColumnWidth = 22.57
        .Range("D:D").ColumnWidth = 30.57
        .Range("E:E").ColumnWidth = 58.14
        .Range("F:F").ColumnWidth = 32.14
        .Range("G:G").ColumnWidth = 30.43
        
        .Range("A1:G1").Merge Across:=True
        .Range("A1:G1").Value = "3. Штатно-посадовий облік"
        .Range("A1:G1").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("A2:A3").Merge
        .Range("A2:A3").Value = "№ з/п"
        .Range("A2:A3").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("B2:B3").Merge
        .Range("B2:B3").Value = "Посада за штатом"
        .Range("B2:B3").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("C2").Value = "ВОС згідно зі" & vbLf & "штатом"
        .Range("C2").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("C3").Value = "Тарифний розряд" & vbLf & "(посадовий оклад)"
        .Range("C3").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("D2").Value = "Військове звання" & vbLf & "згідно зі штатом"
        .Range("D2").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("D3").Value = "Військове звання" & vbLf & "фактичне"
        .Range("D3").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("E2:E3").Merge
        .Range("E2:E3").Value = "Прізвище, власне ім'я, по батькові"
        .Range("E2:E3").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("F2:F3").Merge
        .Range("F2:F3").Value = "Дата закінчення строку" & vbLf & "контракту, період і рік призову" & vbLf & "(дата і рік народження)"
        .Range("F2:F3").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Range("G2:G3").Merge
        .Range("G2:G3").Value = "Примітки"
        .Range("G2:G3").BorderAround LineStyle:=xlContinuous, Weight:=xlThin

        .Range("A:G").Font.Size = 12
        .Range("A:G").Font.name = "Times New Roman"
        .Range("A:G").VerticalAlignment = xlCenter
        .Range("A:G").HorizontalAlignment = xlCenter
        
        .Range("A1:G6").Font.Bold = True
        .Range("A1:G1").Font.Size = 14
        
        With .Range("A4:G4")
            For n = 1 To 7
                .Cells(n).Value = n
                .Cells(n).Font.Bold = True
                .Cells(n).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            Next
        End With
        
        
        ' йдемо по штатці
        cLine = 5
        nn = 0
        Set parRange = Workbooks(wbfound).Worksheets(wsShtat).Range("A3:AU10000")
        cc = 0
        For Each Line In parRange.Rows
        
            If Line.Range("N1").Value <> vbNullString Then
                ' непорожня стрічка
                cc = 0
            
                If Line.Range("J1").Value <> vbNullString Then
                    ' посада
                    nn = nn + 1
                    .Range("A" & cLine & ":G" & cLine).RowHeight = 37.25
                    .Range("A" & cLine + 1 & ":G" & cLine + 1).RowHeight = 37.25
                    .Range("A" & cLine & ":A" & cLine + 1).Merge
'                    .Range("A" & cLine & ":A" & cLine + 1).Font.Color = Line.Range("J1").Font.Color
                    .Range("A" & cLine & ":A" & cLine + 1).Value = nn
                    .Range("A" & cLine & ":A" & cLine + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range("B" & cLine & ":B" & cLine + 1).Merge
                    .Range("B" & cLine & ":B" & cLine + 1).Font.Color = Line.Range("J1").Font.Color
                    .Range("B" & cLine & ":B" & cLine + 1).Value = Line.Range("J1").Value
                    .Range("B" & cLine & ":B" & cLine + 1).HorizontalAlignment = xlLeft
                    .Range("B" & cLine & ":B" & cLine + 1).WrapText = True
                    .Range("B" & cLine & ":B" & cLine + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range("C" & cLine).NumberFormat = "@"
                    .Range("C" & cLine).Font.Color = Line.Range("J1").Font.Color
                    .Range("C" & cLine).Value = Line.Range("L1").Text
                    .Range("C" & cLine).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range("C" & cLine + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range("D" & cLine).Font.Color = Line.Range("J1").Font.Color
                    .Range("D" & cLine).Value = Line.Range("K1").Value
                    .Range("D" & cLine).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range("D" & cLine + 1).Font.Italic = True
                    If fill Then .Range("D" & cLine + 1).Value = Line.Range("M1").Value
                    .Range("D" & cLine + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range("E" & cLine & ":E" & cLine + 1).Merge
                    .Range("E" & cLine & ":E" & cLine + 1).Font.Italic = True
                    If fill Then
                        PIB = trim(Line.Range("N1").Value)
                        If StrComp(PIB, "вакант", vbTextCompare) = 0 Then
                            PIB = ""
                        Else
                            PIB = StrConv(Left(PIB, InStr(PIB, " ")), vbUpperCase) & " " & Right(PIB, Len(PIB) - InStr(PIB, " "))
                        End If
                        .Range("E" & cLine & ":E" & cLine + 1).Value = PIB
                    End If
                    .Range("E" & cLine & ":E" & cLine + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range("F" & cLine & ":F" & cLine + 1).Merge
                    .Range("F" & cLine & ":F" & cLine + 1).Font.Italic = True
                    If fill And Line.Range("X1").Text <> vbNullString Then .Range("F" & cLine & ":F" & cLine + 1).Value = vbLf & "(" & Line.Range("V1").Text & ")"
                    .Range("F" & cLine & ":F" & cLine + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range("G" & cLine & ":G" & cLine + 1).Merge
                    .Range("G" & cLine & ":G" & cLine + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    cLine = cLine + 2
                Else
                    ' підрозділ
                    .Range("A" & cLine & ":G" & cLine).Merge Across:=True
                    .Range("A" & cLine & ":G" & cLine).Value = Line.Range("N1").Value
                    .Range("A" & cLine & ":G" & cLine).Font.Bold = True
                    .Range("A" & cLine & ":G" & cLine).Font.Underline = Line.Range("N1").Font.Underline
                    .Range("A" & cLine & ":G" & cLine).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    cLine = cLine + 1
                End If
            
            Else
                ' порожня стрічка
                cc = cc + 1
                If cc > 9 Then
                    Exit For
                End If
            
            End If
        
        Next
        
    
    End With
    
    OptimizeVBA False
    
    MsgBox ("It's all.")
    
    
End Sub

' процедура шукає помилки у штаті
Sub findBug()
Dim s, spis1, spis2, spis3, spis4, founded, tfounded, founded2, tfounded2, founded3, tfounded3, pp, pp1
Dim endTime As Date, startTime As Date, Interval As Date, fractionDone, dateN, dateC, ddiff
Dim booStatusBarState As Boolean

startTime = Now()

check_workbook
If wbfound = vbNullString Then
    MsgBox ("Ви забули відкрити файл 'Штат'")
    Exit Sub
End If

Range("O2") = "Працюємо..."
booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
OptimizeVBA True

s = 9

Range("I9:I1000").Clear
Range("J9:J1000").Clear
Range("K9:K1000").Clear
Range("L9:L1000").Clear

Range("O9:O1000").Value = ""
Range("P9:P1000").Value = ""
Range("Q9:Q1000").Value = ""
Range("R9:R1000").Value = ""

Range("AA9:AA1000").Clear
Range("AB9:AB1000").Clear
Range("AC9:AC1000").Clear
Range("AD9:AD1000").Clear
Range("AE9:AD1000").Clear

Range("AG9:AB1000").Clear
Range("AH9:AC1000").Clear
Range("AI9:AD1000").Clear
Range("AJ9:AD1000").Clear

spis1 = 0
spis2 = 0
spis3 = 0
spis4 = 0

For pp = 0 To Cache.ppmax - 1

    fractionDone = pp / (Cache.ppmax - 1)
    Application.StatusBar = format(fractionDone, "0%") & " done..."
    Range("O2").Value = "Працюємо... " & format(fractionDone, "0%")
    Interval = Now() - startTime
    Range("O3").Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    DoEvents

    founded = 0
    tfounded = ""
    founded2 = 0
    tfounded2 = ""
    founded3 = 0
    tfounded3 = ""
    
    If Cache.getSHT(pp) = 1 Or Cache.getSHT(pp) = 2 Then
        If IsNumeric(Cache.getIPN(pp)) And Cache.getIPN(pp) > 100000 Then
            dateN = CDate(Cache.getIPN(pp) / 100000 + 1)
            dateC = DateAdd("yyyy", 60, dateN)
            'Debug.Print dateN, dateC
            ddiff = DateDiff("d", Date, dateC, vbMonday, vbFirstJan1)
            'Debug.Print ddiff
            If ddiff < 31 Then
                Range("AG" & s + spis4).Value = spis4 + 1
                Range("AH" & s + spis4).Value = Cache.getPIB(pp)
                Range("AI" & s + spis4).Value = IIf(ddiff < 0, 0 - ddiff & " днів ПІСЛЯ", ddiff & " днів до")
                Range("AJ" & s + spis4).Value = format(dateN, "dd.mm.yyyy")
                'MsgBox ("Alarm!" & vbLf & Cache.getPIB(pp) & " - " & ddiff & " днів до ДН60 " & Format(dateN, "dd.mm.yyyy"))
                spis4 = spis4 + 1
            End If
        End If
    End If

    If StrComp(StrConv(Cache.getsPIB(pp), vbLowerCase), "вакант", vbTextCompare) <> 0 Then

        For pp1 = pp + 1 To Cache.ppmax - 1
        
            If Cache.getsPIB(pp) = Cache.getsPIB(pp1) Then
                If Cache.getIPN(pp) = Cache.getIPN(pp1) Then
                    ' повний дубль - ІПНи і ПІБи збігаються
                    tfounded = tfounded & ", " & Cache.getPIDR(pp1)
                    founded = founded + 1
                Else
                    ' тезки - ПІБи однакові, ІПН різні
                End If
                tfounded2 = tfounded2 & ", " & Cache.getPIDR(pp1)
                founded2 = founded2 + 1
            Else
                If Cache.getIPN(pp) > 100000 And StrComp(Cache.getIPN(pp), "без ІПН", 1) <> 0 And StrComp(Cache.getIPN(pp), "", 1) <> 0 And Cache.getIPN(pp) = Cache.getIPN(pp1) Then
                    ' помилка ІПН - у різних людей однаковий
                    founded3 = founded3 + 1
                    tfounded3 = tfounded3 & ", " & Cache.getPIB(pp1)
                End If
            End If
    
        Next
        If founded > 0 Then
            Range("AA" & s + spis1).Value = spis1 + 1
            Range("AB" & s + spis1).Value = Cache.getPIB(pp)
            Range("AC" & s + spis1).Value = Cache.getIPN(pp)
            Range("AD" & s + spis1).Value = founded + 1
            tfounded = Cache.getPIDR(pp) & tfounded
            Range("AE" & s + spis1).Value = tfounded
            spis1 = spis1 + 1
        End If

        If founded2 > 0 Then
            Range("I" & s + spis2).Value = spis2 + 1
            Range("J" & s + spis2).Value = Cache.getPIB(pp)
            Range("K" & s + spis2).Value = founded2 + 1
            tfounded2 = Cache.getPIDR(pp) & tfounded2
            Range("L" & s + spis2).Value = tfounded2
            If founded > 0 Then
                Range("J" & s + spis2).Interior.Color = RGB(255, 200, 200)
            End If
            spis2 = spis2 + 1
        End If
        
        If founded3 > 0 Then
            Range("O" & s + spis3).Value = spis3 + 1
            Range("P" & s + spis3).Value = Cache.getIPN(pp)
            Range("Q" & s + spis3).Value = founded3 + 1
            Range("R" & s + spis3).Value = Cache.getPIB(pp) & tfounded3
            spis3 = spis3 + 1
        End If
    End If
Next

Range("J8:L100").Sort key1:=Range("J9"), order1:=xlAscending, Header:=xlYes
Range("AB8:AE100").Sort key1:=Range("AB9"), order1:=xlAscending, Header:=xlYes

Range("O2") = "Ok"
endTime = Now()
Interval = endTime - startTime
Range("O3") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")

Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False

OptimizeVBA False

Beep

'    findDoublePIB
'    findWrongIPN
'    findFullNamesake
End Sub

' процедура пошуку дублів ПІБ/ІПН - коли знайдено 2 чи більше пар ПІБ/ІПН
Sub findDoublePIB()
Dim founded, n, s, p, PIB, pPIB, IPN, fractionDone, w, wmax, tfounded, gofind, tRange
Dim endTime As Date, startTime As Date, Interval As Date
Dim booStatusBarState As Boolean
Dim parRange As Range, found As Range, Line As Range
Dim fAddress As String
Dim cWorksheet As Worksheet

check_workbook
If wbfound = vbNullString Or wbfound = "cache" Then
    MsgBox ("Ви забули відкрити файл 'Штат'")
    Exit Sub
End If

startTime = Now()
Range("AA4") = "Працюємо..."
booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
OptimizeVBA True

Range("AA9:AA100").Clear
Range("AB9:AB100").Clear
Range("AC9:AC100").Clear
Range("AD9:AD100").Clear
Range("AE9:AD100").Clear

Set parRange = Workbooks(wbfound).Worksheets(wsShtat).Range("B2:AU10000")

    n = -1
    s = 9
    p = 0
    w = 0
    
    Workbooks(wbfound).Worksheets(wsShtat).AutoFilterMode = False
    parRange.AutoFilter Field:=9, Criteria1:="<>"

    wmax = parRange.Range("J5").CurrentRegion.Rows.count

    For Each Line In parRange.Rows
    
        If w > wmax Then
            Exit For
        End If
    
        w = w + 1
        fractionDone = w / wmax
        Application.StatusBar = format(fractionDone, "0%") & " done..."
        Range("AA4").Value = "Працюємо... " & format(fractionDone, "0%")
        Interval = Now() - startTime
        Range("AA5").Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
        DoEvents
    
        If Not Line.Hidden And Line.Cells(9) <> vbNullString Then
          
          n = n + 1
          founded = 0
          tfounded = ""
          
          If n > 0 Then
            ' ПІБи
            PIB = trim(Line.Cells(13).Text)
            IPN = trim(Line.Cells(23).Text)
            If StrComp(PIB, "вакант", 1) <> 0 And IPN <> vbNullString And StrComp(IPN, "без ІПН", 1) <> 0 Then
            
                For Each cWorksheet In Workbooks(wbfound).Worksheets
                
                    gofind = 0
                    If cWorksheet.name Like wsShtat Or cWorksheet.name Like rozpTempl Or cWorksheet.name = "Призуп.в.сл." Or cWorksheet.name = "Загинувші" Or cWorksheet.name = "в 173 бат" Then
                        tRange = "X3:X10000"
                        pPIB = "N1"
                        gofind = 1
                    End If
                    If cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then
                        tRange = "Z3:Z10000"
                        pPIB = "P1"
                        gofind = 1
                    End If
                
                If gofind = 1 Then
                  With cWorksheet.Range(tRange)
                    Set found = .Find(IPN, LookIn:=xlFormulas)
                    If Not found Is Nothing Then
                      fAddress = found.Address
                      Do
                        If StrComp(found.Text, IPN, 1) = 0 Then
                            If StrComp(trim(found.EntireRow.Range(pPIB).Text), PIB, 1) = 0 Then
                                founded = founded + 1
                            End If
                            If Len(tfounded) > 0 Then
                                tfounded = tfounded & ", "
                            End If
                            If cWorksheet.name Like wsShtat Then _
                                tfounded = tfounded & found.EntireRow.Range("B1").Value
                            If cWorksheet.name Like rozpTempl Then _
                                tfounded = tfounded & "У розп.(" & found.EntireRow.Range("B1").Value & ")"
                            If cWorksheet.name = "Призуп.в.сл." Then _
                                tfounded = tfounded & "Призуп.(" & found.EntireRow.Range("B1").Value & ")"
                            If cWorksheet.name = "Загинувші" Then _
                                tfounded = tfounded & "Загинувші(" & found.EntireRow.Range("B1").Value & ")"
                            If cWorksheet.name = "в 173 бат" Then _
                                tfounded = tfounded & "в 173 бат(" & found.EntireRow.Range("B1").Value & ")"
                            If cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then _
                                tfounded = tfounded & cWorksheet.name & "(" & found.EntireRow.Range("D1").Value & ")"
                        End If
                        Set found = .FindNext(found)
                      Loop While Not found Is Nothing And found.Address <> fAddress
                    End If
                  End With
                End If
                
                Next
                
                  If founded > 1 Then
                    Set found = Range("P9:P100").Find(IPN, LookIn:=xlFormulas)
                    If found Is Nothing Then
                        Range("AA" & s + p).Value = p + 1
                        Range("AB" & s + p).Value = PIB
                        Range("AC" & s + p).Value = IPN
                        Range("AD" & s + p).Value = founded
                        Range("AE" & s + p).Value = tfounded
                        p = p + 1
                    End If
                  End If
            
            
            End If
            
          End If
          
        Else
          If Not Line.Hidden Then
            Exit For
          End If
        End If
    
    Next
    
    parRange.AutoFilter Field:=9

    Range("AB8:AE100").Sort key1:=Range("AB9"), order1:=xlAscending, Header:=xlYes

    OptimizeVBA False

Range("AA4") = "Ok"
endTime = Now()
Interval = endTime - startTime
Range("AA5") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")

Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False

End Sub


' процедура пошуку помилок ІПН - коли 2 чи більше ПІБів мають однаковий ІПН
Sub findWrongIPN()
Dim founded, n, s, p, PIB, pPIB, IPN, fractionDone, w, wmax, tfounded, gofind, tRange
Dim endTime As Date, startTime As Date, Interval As Date
Dim booStatusBarState As Boolean
Dim parRange As Range, found As Range, Line As Range
Dim fAddress As String
Dim cWorksheet As Worksheet

check_workbook
If wbfound = vbNullString Or wbfound = "cache" Then
    MsgBox ("Ви забули відкрити файл 'Штат'")
    Exit Sub
End If

startTime = Now()
Range("O4") = "Працюємо..."
booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
OptimizeVBA True

Range("O9:O100").Value = ""
Range("P9:P100").Value = ""
Range("Q9:Q100").Value = ""
Range("R9:R100").Value = ""
Set parRange = Workbooks(wbfound).Worksheets(wsShtat).Range("B2:AU10000")

    n = -1
    s = 9
    p = 0
    w = 0
    
    Workbooks(wbfound).Worksheets(wsShtat).AutoFilterMode = False
    parRange.AutoFilter Field:=9, Criteria1:="<>"

    wmax = parRange.Range("J5").CurrentRegion.Rows.count

    For Each Line In parRange.Rows
    
        If w > wmax Then
            Exit For
        End If
    
        w = w + 1
        fractionDone = w / wmax
        Application.StatusBar = format(fractionDone, "0%") & " done..."
        Range("O4").Value = "Працюємо... " & format(fractionDone, "0%")
        Interval = Now() - startTime
        Range("O5").Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
        DoEvents
    
        If Not Line.Hidden And Line.Cells(9) <> vbNullString Then
          
          n = n + 1
          founded = 0
          tfounded = ""
          
          If n > 0 Then
            ' ПІБи
            PIB = trim(Line.Cells(13).Text)
            IPN = trim(Line.Cells(23).Text)
            If StrComp(PIB, "вакант", 1) <> 0 And IPN <> vbNullString And StrComp(IPN, "без ІПН", 1) <> 0 Then
            
                For Each cWorksheet In Workbooks(wbfound).Worksheets
                
                    gofind = 0
                    If cWorksheet.name Like wsShtat Or cWorksheet.name Like rozpTempl Or cWorksheet.name = "Призуп.в.сл." Or cWorksheet.name = "Загинувші" Or cWorksheet.name = "в 173 бат" Then
                        tRange = "X3:X10000"
                        pPIB = "N1"
                        gofind = 1
                    End If
                    If cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then
                        tRange = "Z3:Z10000"
                        pPIB = "P1"
                        gofind = 1
                    End If
                
                If gofind = 1 Then
                  With cWorksheet.Range(tRange)
                    Set found = .Find(IPN, LookIn:=xlFormulas)
                    If Not found Is Nothing Then
                      fAddress = found.Address
                      Do
                        If StrComp(found.Text, IPN, 1) = 0 Then
                            If StrComp(trim(found.EntireRow.Range(pPIB).Text), PIB, 1) <> 0 Then
                                founded = founded + 1
                            End If
                            If Len(tfounded) > 0 Then
                                tfounded = tfounded & ", "
                            End If
                            tfounded = tfounded & trim(found.EntireRow.Range(pPIB).Text)
                        End If
                        Set found = .FindNext(found)
                      Loop While Not found Is Nothing And found.Address <> fAddress
                    End If
                  End With
                End If
                
                Next
                
                  If founded > 0 Then
                    Set found = Range("P9:P100").Find(IPN, LookIn:=xlFormulas)
                    If found Is Nothing Then
                        Range("O" & s + p).Value = p + 1
                        Range("P" & s + p).Value = IPN
                        Range("Q" & s + p).Value = founded + 1
                        Range("R" & s + p).Value = tfounded
                        p = p + 1
                    End If
                  End If
            
            
            End If
            
          End If
          
        Else
          If Not Line.Hidden Then
            Exit For
          End If
        End If
    
    Next
    
    parRange.AutoFilter Field:=9

    'Range("P8:R100").Sort key1:=Range("P9"), order1:=xlAscending, Header:=xlYes

    OptimizeVBA False

Range("O4") = "Ok"
endTime = Now()
Interval = endTime - startTime
Range("O5") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")

Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False

End Sub


' процедура для пошуку повних тезок у Штаті
Sub findFullNamesake()
Dim founded, n, s, p, PIB, fractionDone, w, wmax, tfounded, gofind, tRange
Dim endTime As Date, startTime As Date, Interval As Date
Dim booStatusBarState As Boolean
Dim parRange As Range, found As Range, Line As Range
Dim fAddress As String
Dim cWorksheet As Worksheet

check_workbook
If wbfound = vbNullString Or wbfound = "cache" Then
    MsgBox ("Ви забули відкрити файл 'Штат'")
    Exit Sub
End If

startTime = Now()
Range("I4") = "Працюємо..."
booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
OptimizeVBA True

Range("I9:I100").Clear
Range("J9:J100").Clear
Range("K9:K100").Clear
Range("L9:L100").Clear
Set parRange = Workbooks(wbfound).Worksheets(wsShtat).Range("B2:AU10000")

    n = -1
    s = 9
    p = 0
    w = 0
    
    Workbooks(wbfound).Worksheets(wsShtat).AutoFilterMode = False
    parRange.AutoFilter Field:=9, Criteria1:="<>"

    wmax = parRange.Range("J5").CurrentRegion.Rows.count

    For Each Line In parRange.Rows
    
        If w > wmax Then
            Exit For
        End If
    
        w = w + 1
        fractionDone = w / wmax
        Application.StatusBar = format(fractionDone, "0%") & " done..."
        Range("I4").Value = "Працюємо... " & format(fractionDone, "0%")
        Interval = Now() - startTime
        Range("I5").Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
        DoEvents
    
        If Not Line.Hidden And Line.Cells(9) <> vbNullString Then
          
          n = n + 1
          founded = 0
          tfounded = ""
          
          If n > 0 Then
            ' ПІБи
            PIB = trim(Line.Cells(13).Text)
            If StrComp(PIB, "вакант", 1) <> 0 And PIB <> vbNullString Then
            
            
                For Each cWorksheet In Workbooks(wbfound).Worksheets
                
                    gofind = 0
                    If cWorksheet.name Like wsShtat Or cWorksheet.name Like rozpTempl Or cWorksheet.name = "Призуп.в.сл." Or cWorksheet.name = "Загинувші" Or cWorksheet.name = "в 173 бат" Then
                        tRange = "N3:N10000"
                        gofind = 1
                    End If
                    If cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then
                        tRange = "P3:P10000"
                        gofind = 1
                    End If
                
                If gofind = 1 Then
                  With cWorksheet.Range(tRange)
                    Set found = .Find(PIB, LookIn:=xlFormulas)
                    If Not found Is Nothing Then
                      fAddress = found.Address
                      Do
                        If StrComp(trim(found.Text), PIB, 1) = 0 Then
                            founded = founded + 1
                            If founded > 1 Then tfounded = tfounded & ", "
                            If cWorksheet.name Like wsShtat Then _
                                tfounded = tfounded & found.EntireRow.Range("B1").Value
                            If cWorksheet.name Like rozpTempl Then _
                                tfounded = tfounded & "У розп.(" & found.EntireRow.Range("B1").Value & ")"
                            If cWorksheet.name = "Призуп.в.сл." Then _
                                tfounded = tfounded & "Призуп.(" & found.EntireRow.Range("B1").Value & ")"
                            If cWorksheet.name = "Загинувші" Then _
                                tfounded = tfounded & "Загинувші(" & found.EntireRow.Range("B1").Value & ")"
                            If cWorksheet.name = "в 173 бат" Then _
                                tfounded = tfounded & "в 173 бат(" & found.EntireRow.Range("B1").Value & ")"
                            If cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then _
                                tfounded = tfounded & cWorksheet.name & "(" & found.EntireRow.Range("D1").Value & ")"
                        End If
                        Set found = .FindNext(found)
                      Loop While Not found Is Nothing And found.Address <> fAddress
                    End If
                  End With
                End If
                
                Next
                
                  If founded > 1 Then
                    Set found = Range("J9:J100").Find(PIB, LookIn:=xlFormulas)
                    If found Is Nothing Then
                        Range("I" & s + p).Value = p + 1
                        Range("J" & s + p).Value = PIB
                        Range("K" & s + p).Value = founded
                        Range("L" & s + p).Value = tfounded
                        Set found = Range("AB9:AB100").Find(PIB, LookIn:=xlFormulas)
                        If Not found Is Nothing Then
                            Range("J" & s + p).Interior.Color = RGB(255, 200, 200)
                        End If
                        p = p + 1
                    End If
                  End If
            
            
            End If
            
          End If
          
        Else
          If Not Line.Hidden Then
            Exit For
          End If
        End If
    
    Next
    
    parRange.AutoFilter Field:=9

    Range("J8:L100").Sort key1:=Range("J9"), order1:=xlAscending, Header:=xlYes
    
    OptimizeVBA False

Range("I4") = "Ok"
endTime = Now()
Interval = endTime - startTime
Range("I5") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")

Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False

End Sub

' процедура приводит до нормального вигляду стовпці з ПІБ у файлі Штат
Sub normalShtat()
Dim w, wmax, fractionDone
Dim endTime As Date, startTime As Date, Interval As Date
Dim booStatusBarState As Boolean
Dim cWorksheet As Worksheet, wSheet As Worksheet
startTime = Now()

Set wSheet = ActiveSheet

check_workbook
If wbfound = vbNullString Or wbfound = "cache" Then
    MsgBox ("Ви забули відкрити файл 'Штат'")
    Exit Sub
End If

wSheet.Range("C4") = "Працюємо..."
booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
OptimizeVBA True

w = 0
wmax = Workbooks(wbfound).Worksheets.count

For Each cWorksheet In Workbooks(wbfound).Worksheets

    w = w + 1
    fractionDone = w / wmax
    Application.StatusBar = format(fractionDone, "0%") & " done..."
    wSheet.Range("C4").Value = "Працюємо... " & format(fractionDone, "0%")
    Interval = Now() - startTime
    wSheet.Range("C5").Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    DoEvents
    
    If cWorksheet.name Like wsShtat Or cWorksheet.name Like rozpTempl Or cWorksheet.name = "Призуп.в.сл." Or cWorksheet.name = "Загинувші" Or cWorksheet.name = "в 173 бат" Then
        TrimRange wbfound, cWorksheet.name, "N3:N10000"
    End If
    If cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then
        TrimRange wbfound, cWorksheet.name, "P3:P10000"
    End If

Next

OptimizeVBA False

wSheet.Range("C4") = "Ok"
endTime = Now()
Interval = endTime - startTime
wSheet.Range("C5") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")

Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False
Beep
End Sub


' процедура приводит до нормального вигляду стовпці з ПІБ у файлі Втрати
Sub normalVtraty()
Dim w, wmax, fractionDone
Dim endTime As Date, startTime As Date, Interval As Date
Dim booStatusBarState As Boolean, wSheet As Worksheet

check_vtratybook
If vtratyfound = vbNullString Then
    MsgBox ("Ви забули відкрити файл 'втрати'")
    Exit Sub
End If

startTime = Now()

Set wSheet = ActiveSheet

wSheet.Range("C11") = "Працюємо..."
booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
OptimizeVBA True

TrimRange vtratyfound, wsVtraty, "D4:D100000"
TrimRange vtratyfound, wsVtraty, "L4:L100000"
'TrimRange vtratyfound, wsVtraty, "N4:N100000"

OptimizeVBA False

wSheet.Range("C11") = "Ok"
endTime = Now()
Interval = endTime - startTime
wSheet.Range("C12") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")

Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False
Beep
End Sub


' процедура приводит до нормального вигляду стовпці з ПІБ у файлі Відпустки
Sub normalVac()
Dim w, wmax, fractionDone
Dim endTime As Date, startTime As Date, Interval As Date
Dim booStatusBarState As Boolean, wSheet As Worksheet

check_vacbook
If vacfound = vbNullString Then
    MsgBox ("Ви забули відкрити файл 'Відпустки'")
    Exit Sub
End If

startTime = Now()
Set wSheet = ActiveSheet
wSheet.Range("C18") = "Працюємо..."
booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
OptimizeVBA True

TrimRange vacfound, wsVac, "B2:B10000"
TrimRange vacfound, wsVac, "E2:E10000"
TrimRange vacfound, wsCom, "B2:B10000"

OptimizeVBA False

wSheet.Range("C18") = "Ok"
endTime = Now()
Interval = endTime - startTime
wSheet.Range("C19") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")

Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False
Beep
End Sub

' процедура оновлює колонку місцезнайходження вкладок "штат" і "у розп*" файла "Штат" із допоміжної
' таблиці актуальних втрат обособового складу за штатом (вкладка "Втрати" файла "Обробка-Олена") і таблиц відпусток і відряджень
Sub updateVtraty()
Dim m, mmax, Parts, fractionDone, isfound, PIB, gofind, tRange, fAddress, ovactype, tvactype, vacenddate
Dim endTime As Date, startTime As Date, Interval As Date
Dim booStatusBarState As Boolean
Dim curRange As Range, parRange As Range, vacRange As Range
Dim Line As Range, curLine As Range
Dim found As Range, sts As Range, sts2 As Range
Dim cWorksheet As Worksheet
Dim cStatus As Variant, statuses()

Set sts = ActiveSheet.Range("N13")
Set sts2 = ActiveSheet.Range("N14")
startTime = Now()

check_vacbook
If vacfound = vbNullString Then
    MsgBox ("Ви забули відкрити файл 'Відпустки'")
    Exit Sub
End If

check_workbook
If wbfound = vbNullString Or wbfound = "cache" Then
    MsgBox ("Ви забули відкрити файл 'Штат'")
    Exit Sub
End If

sts.Value = "Працюємо..."
booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = "0% done..."

OptimizeVBA True


' видаляємо старі статуси
'statuses = Array("відп", "відр", "сзч", "шпит")
'For Each cStatus In statuses
    For Each cWorksheet In Workbooks(wbfound).Worksheets
    
        sts.Value = "Видаляємо старі статуси... "
        Interval = Now() - startTime
        sts2.Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
        DoEvents
    
        gofind = 0
        If cWorksheet.name Like wsShtat Or cWorksheet.name Like rozpTempl Then
            gofind = 1
        End If
        If gofind = 1 Then
            For Each Line In cWorksheet.Range("Q3:Q10000").Rows
                Set Line = Line.EntireRow
                If Line.Range("A1").Value <> "" Then
                    gofind = 1
                    If InStr("відп відр сзч шпит", StrConv(trim(Line.Range("Q1").Value), vbLowerCase)) > 0 Then
                        If StrConv(Left(Line.Range("R1").Value, 3), vbLowerCase) <> "влк" Then
                            Line.Range("Q1").Value = "оос"
                            Line.Range("R1").Value = ""
                        End If
                    End If
                    If StrConv(trim(Line.Range("N1").Value), vbLowerCase) = "вакант" Then
                        Line.Range("Q1").Value = ""
                        Line.Range("R1").Value = ""
                    End If
                Else
                    If gofind > 10 Then Exit For
                    gofind = gofind + 1
                End If
            Next
        End If
    Next
'Next

'If False Then
' вносимо втрати
Set curRange = Range("A2:M100000")

m = 0
mmax = curRange.CurrentRegion.Rows.count

For Each curLine In curRange.Rows
    m = m + 1

    fractionDone = m / mmax
    Application.StatusBar = format(fractionDone, "0%") & " done..."
    sts.Value = "Працюємо... " & format(fractionDone, "0%")
    Interval = Now() - startTime
    sts2.Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    DoEvents
    
    PIB = trim(curLine.Cells(1).Text)
    
'    If PIB Like "Білик Сергій Олександрович*" Then
'    If PIB Like "Груша Сергій Валентинович*" Then
'        Debug.Print "we here"
'    End If
    
    If PIB <> vbNullString Then
        
      Set found = Workbooks(wbfound).Worksheets(wsShtat).Range("N:N").Find(PIB, LookIn:=xlFormulas)
        
      isfound = 0
      If Not found Is Nothing Then
          fAddress = found.Address
          Do
      
            If StrComp(trim(found.Text), PIB, 1) = 0 Then
                isfound = 1
                
                If StrComp(trim(found.EntireRow.Range("B1").Value), trim(curLine.Cells(3).Value), 1) <> 0 Or StrComp(trim(found.EntireRow.Range("X1").Value), trim(curLine.Cells(2).Value), 1) <> 0 Then
                    ' підрозділ або ІПН не збігаються
                    If curLine.EntireRow.Range("A1").Interior.Color = RGB(255, 255, 0) Then
                        ' ПІБ має дубль у штатці - треба пошукати по ІПН
                        If StrComp(trim(found.EntireRow.Range("X1").Value), trim(curLine.Cells(2).Value), 1) <> 0 Then
                            ' ІПН не збігається - треба шукати інший варіант
                            isfound = 0
                        Else
                            ' ІПН збігається - наш клієнт
                            curLine.Cells(13).Value = trim(found.EntireRow.Range("B1").Value) & " (" & trim(curLine.Cells(3).Value) & ")"
                            curLine.Cells(13).Interior.Color = RGB(255, 200, 200)
                        End If
                    Else
                        ' підрозділ інший, але дубля немає
                        curLine.Cells(13).Value = trim(found.EntireRow.Range("B1").Value) & " (" & trim(curLine.Cells(3).Value) & ")"
                        curLine.Cells(13).Interior.Color = RGB(255, 200, 200)
                    End If
                Else
                    curLine.Cells(13).Value = found.EntireRow.Range("B1").Value
                End If
                
                If isfound = 1 Then
                    If StrConv(Left(found.EntireRow.Range("R1").Value, 3), vbLowerCase) <> "влк" Then
                        If curLine.Cells(7).Text = 1 Then
                            found.EntireRow.Range("Q1").Value = "оос"   ' У строю
                            found.EntireRow.Range("R1").Value = ""
                        End If
                        If curLine.Cells(8).Text = 1 Then
                            found.EntireRow.Range("Q1").Value = "шпит"   ' Хворий
                            found.EntireRow.Range("R1").Value = StrConv(curLine.Cells(5).Text, vbLowerCase) & vbLf & "з " & curLine.Cells(4).Text
                        End If
                        If curLine.Cells(9).Text = 1 Then
                            found.EntireRow.Range("Q1").Value = "шпит"   ' 300
                            found.EntireRow.Range("R1").Value = StrConv(curLine.Cells(5).Text, vbLowerCase) & vbLf & "з " & curLine.Cells(4).Text
                        End If
                        If curLine.Cells(10).Text = 1 Then
                            found.EntireRow.Range("Q1").Value = "загинув"  ' 200
                            found.EntireRow.Range("R1").Value = curLine.Cells(4).Text
                        End If
                        If curLine.Cells(11).Text = 1 Then
                            found.EntireRow.Range("Q1").Value = "зниклий безвісти" ' 500
                            found.EntireRow.Range("R1").Value = curLine.Cells(4).Text
                        End If
                        If curLine.Cells(12).Text = 1 Then
                            found.EntireRow.Range("Q1").Value = "сзч" ' сзч
                            found.EntireRow.Range("R1").Value = "з " & curLine.Cells(4).Text
                        End If
                    Else
                        curLine.Cells(13).Value = "**"
                    End If
                    Exit Do
                End If
            End If
        
            Set found = Workbooks(wbfound).Worksheets(wsShtat).Range("N:N").FindNext(found)
        Loop While Not found Is Nothing And found.Address <> fAddress
        
      End If
      If isfound = 0 Then

            For Each cWorksheet In Workbooks(wbfound).Worksheets
                
                gofind = 0
                If cWorksheet.name Like rozpTempl Then
                    gofind = 1
                End If
                
                If gofind = 1 Then
                  With cWorksheet.Range("N:N")
                    Set found = .Find(PIB, LookIn:=xlFormulas)
                    If Not found Is Nothing Then
                        If StrComp(trim(found.Text), PIB, 1) = 0 Then
                            If StrConv(Left(found.EntireRow.Range("R1").Value, 3), vbLowerCase) <> "влк" Then
                                If curLine.Cells(7).Text = 1 Then
                                    found.EntireRow.Range("Q1").Value = "оос"   ' У строю
                                    found.EntireRow.Range("R1").Value = ""
                                End If
                                If curLine.Cells(8).Text = 1 Then
                                    found.EntireRow.Range("Q1").Value = "шпит"   ' Хворий
                                    found.EntireRow.Range("R1").Value = StrConv(curLine.Cells(5).Text, vbLowerCase) & vbLf & "з " & curLine.Cells(4).Text
                                End If
                                If curLine.Cells(9).Text = 1 Then
                                    found.EntireRow.Range("Q1").Value = "шпит"   ' 300
                                    found.EntireRow.Range("R1").Value = StrConv(curLine.Cells(5).Text, vbLowerCase) & vbLf & "з " & curLine.Cells(4).Text
                                End If
                                If curLine.Cells(10).Text = 1 Then
                                    found.EntireRow.Range("Q1").Value = "загинув"  ' 200
                                    found.EntireRow.Range("R1").Value = curLine.Cells(4).Text
                                End If
                                If curLine.Cells(11).Text = 1 Then
                                    found.EntireRow.Range("Q1").Value = "зниклий безвісти" ' 500
                                    found.EntireRow.Range("R1").Value = curLine.Cells(4).Text
                                End If
                                If curLine.Cells(12).Text = 1 Then
                                    found.EntireRow.Range("Q1").Value = "сзч" ' сзч
                                    found.EntireRow.Range("R1").Value = "з " & curLine.Cells(4).Text
                                End If
                                curLine.Cells(13).Value = "У розп. (" & trim(curLine.Cells(3).Value) & ")"
                                curLine.Cells(13).Interior.Color = RGB(255, 200, 200)
                            End If
                            Exit For
                        End If
                    End If
                  End With
                End If
                
            Next
        End If
    Else
        Exit For
    End If
Next

'End If

' відпустки
Set vacRange = Workbooks(vacfound).Worksheets(wsVac).Range("A1:K10000")

Workbooks(vacfound).Worksheets(wsVac).AutoFilterMode = False
vacRange.AutoFilter Field:=6, Criteria1:="<=" & CDbl(Date)
'vacRange.AutoFilter Field:=7, Criteria1:=">=" & CDbl(Date)

  For Each Line In vacRange.Rows
  
    Application.StatusBar = "Відпустки..."
    sts.Value = "Відпустки... "
    Interval = Now() - startTime
    sts2.Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    DoEvents
  
    PIB = trim(Line.Cells(2).Text)
    vacenddate = trim(Line.Cells(11).Value)
    If vacenddate = vbNullString Then
        vacenddate = trim(Line.Cells(7).Value)
    End If
    If IsDate(vacenddate) Then
        If IsNumeric(vacenddate) Then
          vacenddate = DateValue(format(vacenddate, "dd/mm/yyyy"))
        Else
          vacenddate = DateValue(vacenddate)
        End If
    Else
        vacenddate = trim(Line.Cells(7).Value)
        If IsNumeric(vacenddate) Then
          vacenddate = DateValue(format(vacenddate, "dd/mm/yyyy"))
        Else
            If IsDate(vacenddate) Then
              vacenddate = DateValue(vacenddate)
            Else
              vacenddate = ""
            End If
        End If
    End If
    ' Debug.Print vacenddate

    If Not Line.Hidden And PIB <> vbNullString And PIB <> "ПІБ" And vacenddate <> vbNullString And Not trim(StrConv(Line.Cells(11).Value, vbLowerCase)) Like "*анульовано*" Then
  
        If vacenddate >= DateValue(format(Now(), "dd/mm/yyyy")) Then

            Set ovactype = Line.Cells(5)
            If ovactype Like "*ля лікування*" Then
                tvactype = "СЗ"
            Else
                If ovactype Like "*щорічн*" Then
                    tvactype = "ЩР"
                Else
                    tvactype = "СО"
                End If
            End If
      
            Set found = Workbooks(wbfound).Worksheets(wsShtat).Range("N:N").Find(PIB, LookIn:=xlFormulas)
        
            isfound = 0
            If Not found Is Nothing Then
                If StrComp(trim(found.Text), PIB, 1) = 0 Then
                    isfound = 1
                    Select Case tvactype
                        Case "СЗ"
                            found.EntireRow.Range("Q1").Value = "відп"
                            found.EntireRow.Range("R1").Value = "для лікування по " & Line.Cells(7).Text
                        Case "ЩР"
                            found.EntireRow.Range("Q1").Value = "відп"
                            found.EntireRow.Range("R1").Value = "щорічна по " & Line.Cells(7).Text
                        Case "СО"
                            found.EntireRow.Range("Q1").Value = "відп"
                            found.EntireRow.Range("R1").Value = "по сімейним по " & Line.Cells(7).Text
                    End Select
                End If
            End If
            If isfound = 0 Then

                For Each cWorksheet In Workbooks(wbfound).Worksheets
                    
                    gofind = 0
                    If cWorksheet.name Like rozpTempl Then
                        gofind = 1
                    End If
                    
                    If gofind = 1 Then
                      With cWorksheet.Range("N:N")
                        Set found = .Find(PIB, LookIn:=xlFormulas)
                        If Not found Is Nothing Then
                            If StrComp(trim(found.Text), PIB, 1) = 0 Then
                            
                                Select Case tvactype
                                    Case "СЗ"
                                        found.EntireRow.Range("Q1").Value = "відп"
                                        found.EntireRow.Range("R1").Value = "для лікування по " & Line.Cells(7).Text
                                    Case "ЩР"
                                        found.EntireRow.Range("Q1").Value = "відп"
                                        found.EntireRow.Range("R1").Value = "щорічна по " & Line.Cells(7).Text
                                    Case "СО"
                                        found.EntireRow.Range("Q1").Value = "відп"
                                        found.EntireRow.Range("R1").Value = "по сімейним по " & Line.Cells(7).Text
                                End Select
                                Exit For
                            
                            End If
                        End If
                      End With
                    End If
                    
                Next
            End If
        End If
    Else
      If PIB = vbNullString Then
        Exit For
      End If
    End If
  Next


' відрядження
Set vacRange = Workbooks(vacfound).Worksheets(wsCom).Range("A2:J10000")

Workbooks(vacfound).Worksheets(wsCom).AutoFilterMode = False
vacRange.AutoFilter Field:=6, Criteria1:="<=" & CDbl(Date)
'vacRange.AutoFilter Field:=7, Criteria1:=">=" & CDbl(Date), Operator:=xlOr, Criteria2:="="

  For Each Line In vacRange.Rows
  
    Application.StatusBar = "Відрядження..."
    sts.Value = "Відрядження... "
    Interval = Now() - startTime
    sts2.Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    DoEvents
  
    PIB = trim(Line.Cells(2).Text)
    vacenddate = trim(Line.Cells(8).Value)
    If vacenddate = vbNullString Then
        vacenddate = trim(Line.Cells(7).Value)
    End If
    If IsDate(vacenddate) Then
        If IsNumeric(vacenddate) Then
          vacenddate = DateValue(format(vacenddate, "dd/mm/yyyy"))
        Else
          vacenddate = DateValue(vacenddate)
        End If
    Else
        vacenddate = trim(Line.Cells(7).Value)
        If IsNumeric(vacenddate) Then
          vacenddate = DateValue(format(vacenddate, "dd/mm/yyyy"))
        Else
            If IsDate(vacenddate) Then
              vacenddate = DateValue(vacenddate)
            Else
              vacenddate = ""
            End If
        End If
    End If
    
'    If PIB = "Волочій Андрій Васильович" Then
'        Debug.Print "we here"
'    End If
  
    If Not Line.Hidden And PIB <> vbNullString And PIB <> "ПІБ" And (vacenddate >= DateValue(format(Now(), "dd/mm/yyyy")) Or vacenddate = vbNullString) And Not trim(StrConv(Line.Cells(8).Value, vbLowerCase)) Like "*анульовано*" Then
  
      Set found = Workbooks(wbfound).Worksheets(wsShtat).Range("N:N").Find(PIB, LookIn:=xlFormulas)
        
      isfound = 0
      If Not found Is Nothing Then
        If StrComp(trim(found.Text), PIB, 1) = 0 Then
            isfound = 1
            found.EntireRow.Range("Q1").Value = "відр"
            If trim(Line.Cells(7).Text) = vbNullString Then
                found.EntireRow.Range("R1").Value = Line.Cells(10).Text
            Else
                found.EntireRow.Range("R1").Value = Line.Cells(10).Text & " по " & Line.Cells(7).Text
            End If
        End If
      End If
      If isfound = 0 Then

            For Each cWorksheet In Workbooks(wbfound).Worksheets
                
                gofind = 0
                If cWorksheet.name Like rozpTempl Then
                    gofind = 1
                End If
                
                If gofind = 1 Then
                  With cWorksheet.Range("N:N")
                    Set found = .Find(PIB, LookIn:=xlFormulas)
                    If Not found Is Nothing Then
                        If StrComp(trim(found.Text), PIB, 1) = 0 Then
                            found.EntireRow.Range("Q1").Value = "відр"
                            If trim(Line.Cells(7).Text) = vbNullString Then
                                found.EntireRow.Range("R1").Value = Line.Cells(10).Text
                            Else
                                found.EntireRow.Range("R1").Value = Line.Cells(10).Text & " по " & Line.Cells(7).Text
                            End If
                            Exit For
                        End If
                    End If
                  End With
                End If
                
            Next
        End If
      
    Else
      If PIB = vbNullString Then
        Exit For
      End If
    End If
  Next

OptimizeVBA False

sts.Value = "Ok"
endTime = Now()
Interval = endTime - startTime
sts2.Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")

Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False

End Sub



' процедура заповнює допоміжню таблицю актуальних втрат обособового складу за штатом (вкладка "Втрати" файла "Олена-обробка")
' незакриті втрати намагаємося закривати наступними відпустками
' version 2
Sub fillVtraty2()
Dim n, m, mmax, PIB, fractionDone, invac, z, iRow, r2, r2f, vtrdate, vacdate, cpos, ppos, smax, smin, speed
Dim booStatusBarState As Boolean
Dim endTime As Date, startTime As Date, Interval As Date
Dim prevtime As Date
Dim vtrArray As Variant
Dim vacArray As Variant
Dim tarray As Variant
Dim shtatArray(10000) As Variant
Dim vtr2Array(100000) As Variant
Dim curRange As Range
Dim parRange As Range
Dim vacRange As Range
Dim found As Range

check_vtratybook
If vtratyfound = vbNullString Then
    MsgBox ("Ви забули відкрити файл 'втрати'")
    Exit Sub
End If

check_vacbook
If vacfound = vbNullString Then
    MsgBox ("Ви забули відкрити файл 'Відпустки'")
    Exit Sub
End If

check_workbook
If wbfound = vbNullString Or wbfound = "cache" Then
    MsgBox ("Ви забули відкрити файл 'Штат'")
    Exit Sub
End If

startTime = Now()

Range("N5") = "Працюємо..."
DoEvents

booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True
OptimizeVBA True


Set curRange = Range("A2:M100000")

With Workbooks(vtratyfound).Worksheets(wsVtraty).Sort
     .SortFields.Add Key:=Range("N1"), Order:=xlDescending
     .SetRange Range("A3:AP100000")
     .Header = xlYes
     .Apply
End With

With Workbooks(vacfound).Worksheets(wsVac).Sort
     .SortFields.Add Key:=Range("F1"), Order:=xlDescending
     .SetRange Range("A1:K10000")
     .Header = xlYes
     .Apply
End With

Set parRange = Workbooks(vtratyfound).Worksheets(wsVtraty).Range("A4:P100000")
Set vacRange = Workbooks(vacfound).Worksheets(wsVac).Range("A2:K10000")

curRange.Clear
m = 0
mmax = parRange.CurrentRegion.Rows.count

vtrArray = Workbooks(vtratyfound).Worksheets(wsVtraty).Range("A4:P" & mmax).Value2
vacArray = Workbooks(vacfound).Worksheets(wsVac).Range("A2:K10000").Value2
tarray = Workbooks(wbfound).Worksheets(wsShtat).Range("N:N").Value2
n = 0
For iRow = 3 To 100000
  If tarray(iRow, 1) <> vbNullString Then
    shtatArray(iRow - 3) = tarray(iRow, 1)
    n = 0
  Else
    n = n + 1
    If n > 10 Then
      Exit For
    End If
  End If
Next

n = 0
ppos = 0
prevtime = Now()
smax = 0
smin = 100
'For Each parLine In parRange.Rows
For iRow = LBound(vtrArray, 1) To UBound(vtrArray, 1)

    m = m + 1
    fractionDone = m / mmax
    Application.StatusBar = format(fractionDone, "0%") & " done..."
    Range("N5").Value = "Працюємо... " & format(fractionDone, "0%")
    Interval = Now() - startTime
    Range("N6").Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    
    Interval = Now() - prevtime
    If (format(Interval, "ss") > 5) Then
        cpos = m
        speed = (cpos - ppos) / format(Interval, "ss")
        Range("O7").Value = speed
        prevtime = Now()
        ppos = cpos
        If smin > speed Then smin = speed
        If smax < speed Then smax = speed
        Range("N7").Value = smin
        Range("P7").Value = smax
    End If
    
    DoEvents

    'If parLine.Cells(4).Text <> "" Then
    If vtrArray(iRow, 4) <> vbNullString Then
    
        PIB = StrConv(trim(Replace(vtrArray(iRow, 4), Chr(10), " ")), vbProperCase)
        
        'Set found = curRange.Range("A1:A100000").Find(PIB, LookIn:=xlFormulas)
        z = Filter(vtr2Array, PIB, , vbTextCompare)
        'If Not found Is Nothing Then
        If UBound(z) > -1 Then
            ' такий вже був - пропускаємо
        Else
        
            'Set found = Workbooks(wbfound).Worksheets(wsShtat).Range("N:N").Find(PIB, LookIn:=xlFormulas)
            z = Filter(shtatArray, PIB, , vbTextCompare)
            'If Not found Is Nothing Then
            If UBound(z) > -1 Then
              n = n + 1
              With curRange.Rows(n)
                Set found = Worksheets("Послідовність").Range("I9:L100").Find(PIB, LookIn:=xlFormulas)
                If Not found Is Nothing Then
                    .EntireRow.Range("A1:M1").Interior.Color = RGB(255, 255, 0)
                End If
                vtr2Array(n - 1) = PIB
                .Cells(1).Value = PIB
                .Cells(2).Value = trim(Replace(vtrArray(iRow, 5), Chr(10), " "))
                .Cells(2).HorizontalAlignment = xlHAlignCenter
                .Cells(3).Value = trim(Replace(vtrArray(iRow, 2), Chr(10), " "))
                If IsNumeric(vtrArray(iRow, 14)) Then
                  vtrdate = DateValue(format(vtrArray(iRow, 14), "dd/mm/yyyy"))
                Else
                  vtrdate = DateValue(vtrArray(iRow, 14))
                End If
                .Cells(4).Value = vtrdate
                .Cells(5).Value = StrConv(trim(Replace(vtrArray(iRow, 12), Chr(10), " ")), vbProperCase)
                .Cells(6).Value = StrConv(trim(Replace(vtrArray(iRow, 16), Chr(10), " ")), vbProperCase)
                
                If .Cells(6).Value <> vbNullString Then
                    If StrComp(.Cells(6).Value, "Загибель", 1) = 0 Then
                        .Cells(10).Value = 1
                    Else
                        .Cells(7).Value = 1
                    End If
                Else
                    ' втрата не закрита - спробуємо закрити відпусткою
                    invac = 0
                    'Set found = vacRange.Find(PIB, LookIn:=xlFormulas)
                    r2f = Array()
                    For r2 = LBound(vacArray, 1) To UBound(vacArray, 1)
                      If vacArray(r2, 2) = vbNullString Then
                        Exit For
                      End If
                      If StrComp(trim(vacArray(r2, 2)), PIB, 1) = 0 Then
                        If IsNumeric(vacArray(r2, 6)) Then vacdate = format(vacArray(r2, 6), "dd/mm/yyyy") Else vacdate = vacArray(r2, 6)
                        If DateValue(vacdate) > DateValue(vtrdate) Then
                          invac = 1
                          Exit For
                        End If
                      End If
                    Next
                    'If Not found Is Nothing Then
                    '    If DateValue(found.Offset(columnOffset:=4).Text) > DateValue(curRange.Rows(n).Cells(3).Text) Then
                    '        invac = 1
                    '    End If
                    'End If
                    If invac = 1 Then
                        .Cells(7).Value = 1
                        'curRange.Rows(n).Cells(5).Value = "закрито відпусткою " & found.Offset(columnOffset:=4).Text
                        .Cells(6).Value = "закрито відпусткою " & vacdate
                    Else
                        If StrComp(.Cells(5).Value, "Ушкодження", 1) = 0 Then .Cells(8).Value = 1 Else If StrComp(.Cells(5).Value, "Поранення", 1) = 0 Or StrComp(.Cells(5).Value, "Травмування", 1) = 0 Then .Cells(9).Value = 1 Else If StrComp(.Cells(5).Value, "Загибель", 1) = 0 Then .Cells(10).Value = 1 Else If StrComp(.Cells(5).Value, "Зниклий Безвісті", 1) = 0 Then .Cells(11).Value = 1 Else If StrComp(.Cells(5).Value, "сзч", 1) = 0 Then .Cells(12).Value = 1
                    End If
                End If
              End With
            End If
        End If

    Else
        Exit For
    End If
Next

OptimizeVBA False
Application.DisplayStatusBar = booStatusBarState
Application.StatusBar = False

Range("N5") = "Ok"
endTime = Now()
Interval = endTime - startTime
Range("N6") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    
End Sub

' видаляємо зайві пробели
Sub TrimRange(curWB, curSheet As String, curRange As String)
Dim rng As Range, Line As Range
Dim m
Set rng = Workbooks(curWB).Sheets(curSheet).Range(curRange)
'rng.Replace what:=Chr(10), lookat:=xlPart, replacement:=" "
m = 0
For Each Line In rng.Rows
    If Not Line.Hidden And Line.Cells(1).Value <> vbNullString Then
        Line.Cells(1).Value = Application.trim(Replace(Replace(Line.Cells(1).Value, Chr(160), " "), Chr(10), " "))
        m = 0
    Else
        If Not Line.Hidden And Line.Cells(1).Value = vbNullString Then
            m = m + 1
            If m > 10 Then
                Exit For
            End If
        End If
    End If
Next
End Sub

' процедура заповнення кешу
Sub fillArrays()
    Dim cWorksheet As Worksheet, parRange As Range, Line As Range

    OptimizeVBA True
    Cache.ppmax = 0
    For Each cWorksheet In Workbooks(wbfound).Worksheets
    
      If cWorksheet.name Like wsShtat Or cWorksheet.name Like rozpTempl Or cWorksheet.name = "Призуп.в.сл." Or cWorksheet.name = "Загинувші" Or cWorksheet.name = "в 173 бат" Or cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then
    
        Set parRange = cWorksheet.Range("A2:AV10000")
        For Each Line In parRange.Rows
    
            If Not Line.Hidden And (Line.Range("N1").Value <> vbNullString Or Line.Range("P1").Value <> vbNullString) And StrComp(Line.Range("N1").Value, "ПІБ", 1) <> 0 And StrComp(Line.Range("N1").Value, "Постійний склад", 1) <> 0 And StrComp(Line.Range("P1").Value, "Постійний склад", 1) <> 0 Then
                
                
                If cWorksheet.name Like wsShtat Or cWorksheet.name Like rozpTempl Or cWorksheet.name = "Призуп.в.сл." Or cWorksheet.name = "Загинувші" Or cWorksheet.name = "в 173 бат" Then
                    If Line.Range("J1").Value <> vbNullString Or Line.Range("M1").Value <> vbNullString Then
                        Cache.setPIB Cache.ppmax, Line.Range("N1").Value
                        Cache.setsPIB Cache.ppmax, clearstring(Line.Range("N1").Value)
                        Cache.setIPN Cache.ppmax, Line.Range("X1").Value
                        Cache.setSex Cache.ppmax, IIf(trim(Line.Range("AU1").Value) <> vbNullString, жінка, чоловік)
                        Cache.setZV Cache.ppmax, Application.trim(Line.Range("M1").Value)
                        Cache.setPOS Cache.ppmax, Application.trim(Line.Range("J1").Value)
                        Cache.setRVK Cache.ppmax, Application.trim(Line.Range("Y1").Value)
                        Cache.setAddr Cache.ppmax, Application.trim(Line.Range("AV1").Value)
                        Cache.setPhone Cache.ppmax, Application.trim(Line.Range("AD1").Text)
                        Cache.setPochatok Cache.ppmax, Application.trim(Line.Range("AA1").Text)
                        If cWorksheet.name Like wsShtat Then
                            Cache.setPIDR Cache.ppmax, Application.trim(Line.Range("B1").Value)
                            Cache.setSHT Cache.ppmax, 1  ' штат
                        Else
                            If cWorksheet.name = "Загинувші" Then
                                Cache.setSHT Cache.ppmax, 3  ' загінувші
                            ElseIf cWorksheet.name Like "Призуп*" Then
                                Cache.setSHT Cache.ppmax, 5  ' Призуп.в.сл.
                            ElseIf cWorksheet.name = "в 173 бат" Then
                                Cache.setSHT Cache.ppmax, 6  ' в 173 бат
                            ElseIf cWorksheet.name Like "*СЗЧ*" Then
                                Cache.setSHT Cache.ppmax, 7  ' У розпор. (СЗЧ)
                            Else
                                Cache.setSHT Cache.ppmax, 2  ' у розп.
                            End If
                            Cache.setPIDR Cache.ppmax, cWorksheet.name & "(" & Application.trim(Line.Range("B1").Value) & ")"
                        End If
                        If IsNumeric(Cache.getIPN(Cache.ppmax)) Then
                            Cache.setIPN Cache.ppmax, trim(Str(Cache.getIPN(Cache.ppmax)))
                        End If
                        Cache.ppmax = Cache.ppmax + 1
                    End If
                End If
                If cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then
                    If Line.Range("L1").Value <> vbNullString Or Line.Range("O1").Value <> vbNullString Then
                        Cache.setPIB Cache.ppmax, Line.Range("P1").Value
                        Cache.setsPIB Cache.ppmax, clearstring(Line.Range("P1").Value)
                        Cache.setIPN Cache.ppmax, Line.Range("Z1").Value
                        Cache.setSex Cache.ppmax, IIf(trim(Line.Range("AW1").Value) <> vbNullString, жінка, чоловік)
                        Cache.setZV Cache.ppmax, Application.trim(Line.Range("O1").Value)
                        Cache.setPOS Cache.ppmax, Application.trim(Line.Range("L1").Value)
                        Cache.setRVK Cache.ppmax, Application.trim(Line.Range("AA1").Value)
                        Cache.setPIDR Cache.ppmax, cWorksheet.name & "(" & Application.trim(Line.Range("D1").Value) & ")"
                        Cache.setSHT Cache.ppmax, 4  ' звільнені або переведені
                        Cache.setAddr Cache.ppmax, Application.trim(Line.Range("AX1").Value)
                        Cache.setPhone Cache.ppmax, Application.trim(Line.Range("AF1").Value)
                        Cache.setPochatok Cache.ppmax, Application.trim(Line.Range("AC1").Text)
                        If IsNumeric(Cache.getIPN(Cache.ppmax)) Then
                            Cache.setIPN Cache.ppmax, trim(Str(Cache.getIPN(Cache.ppmax)))
                        End If
                        Cache.ppmax = Cache.ppmax + 1
                    End If
                End If
                

            End If
        Next
      End If
    Next
    OptimizeVBA False
End Sub

' процедура перевіряє чи відкрито файл зі Штатом за маскою
Sub check_workbook()
Dim WB As Workbook
OptimizeVBA True
Randomize
For Each WB In Workbooks
    If WB.name Like wsFile Then
        wbfound = WB.name
        If Cache Is Nothing Then
            Set Cache = New Cache
            fillArrays
            Cache.Creator = Int(100000 * Rnd + 1)
            WB.Worksheets(wsShtat).Range("A5700").Value = Cache.Creator
            'Cache.hwnd = WB.Signatures.Application.hwnd
            'WB.Saved = True
        Else
            'If WB.Signatures.Application.hwnd <> cache.hwnd Or WB.Saved = False Then
            'If WB.Signatures.Application.hwnd <> Cache.hwnd Then
            If WB.Worksheets(wsShtat).Range("A5700").Value <> Cache.Creator Then
                Set Cache = Nothing
                Set Cache = New Cache
                fillArrays
                Cache.Creator = Int(100000 * Rnd + 1)
                WB.Worksheets(wsShtat).Range("A5700").Value = Cache.Creator
                'Cache.hwnd = WB.Signatures.Application.hwnd
                'WB.Saved = True
            End If
        End If
        Exit Sub
    End If
Next WB
If Not Cache Is Nothing Then
    If Cache.ppmax > 100 Then
        wbfound = "cache"
        Exit Sub
    End If
End If
wbfound = ""
OptimizeVBA False
End Sub


' процедура перевіряє чи відкрито файл зі втратами за маскою
Sub check_vtratybook()
Dim WB As Workbook
For Each WB In Workbooks
    If WB.name Like wsFileVtraty Then
        vtratyfound = WB.name
        Exit Sub
    End If
Next WB
vtratyfound = ""
End Sub


' процедура перевіряє чи відкрито файл з відпустками за маскою
Sub check_vacbook()
Dim WB As Workbook
For Each WB In Workbooks
    If WB.name Like wsFileVac Then
        vacfound = WB.name
        Exit Sub
    End If
Next WB
vacfound = ""
End Sub

' процедура перевіряє чи відкрито файл банку за маскою
Sub check_bankbook()
Dim WB As Workbook
For Each WB In Workbooks
    If WB.name Like wsFileBank Then
        bankfound = WB.name
        Exit Sub
    End If
Next WB
bankfound = ""
End Sub

' функція видалення зайвих символів
Function clearstring(inputStr) As String
Dim SpecialCharacters As String, newstring As String, char
newstring = StrConv(inputStr, vbLowerCase)
SpecialCharacters = "!,.,',`,’,"",?," & Chr(160) & "," & vbLf & "," & vbCr
For Each char In Split(SpecialCharacters, ",")
    newstring = Replace(newstring, char, "")
Next
newstring = Replace(newstring, "c", "с")
newstring = Replace(newstring, "b", "в")
newstring = Replace(newstring, "t", "т")
newstring = Replace(newstring, "u", "и")
newstring = Replace(newstring, "h", "н")
newstring = Replace(newstring, "k", "к")
newstring = Replace(newstring, "e", "е")
newstring = Replace(newstring, "n", "п")
newstring = Replace(newstring, "m", "м")
newstring = Replace(newstring, "y", "у")
newstring = Replace(newstring, "i", "і")
newstring = Replace(newstring, "o", "о")
newstring = Replace(newstring, "p", "р")
newstring = Replace(newstring, "a", "а")
newstring = Replace(newstring, "x", "х")
newstring = Replace(newstring, "бекесюк", "бекасюк")
newstring = Replace(newstring, "констянтин", "костянтин")
newstring = Replace(newstring, "олекандр", "олександр")
newstring = Replace(newstring, "маргамович", "марганович")
newstring = Replace(newstring, "волоимир", "володимир")
newstring = Replace(newstring, "григрович", "григорович")
newstring = Replace(newstring, "олексанр", "олександр")
newstring = Replace(newstring, "стеанович", "степанович")
newstring = Replace(newstring, "зіновйович", "зіновійович")
newstring = Replace(newstring, "аанатолійович", "анатолійович")
newstring = Replace(newstring, "кирило", "кирил")
clearstring = Application.trim(newstring)
End Function

' процедура перевіряє поточний діапазон файла банку на наявність у Штаті
Sub Check_PIBs()
Dim cc, tPIB As String, sPIB As String, IPN, vybor, finding, isfound, m, mmax, fractionDone, nz, z, pIPN, p, p2, gofind, tRange, founded, tfounded As String, fAddress
Dim parRange As Range, Line As Range, found As Range
Dim cWorksheet As Worksheet
Dim booStatusBarState As Boolean
Dim endTime As Date, startTime As Date, Interval As Date
Dim PIBs(10000), IPNs(10000), sPIBs(10000), pp, ppmax
Dim array1(1000), max1, m1, array2(1000), max2, m2

cc = 0
nz = 0
z = 0

check_workbook
If wbfound = vbNullString Or wbfound = "cache" Then
    MsgBox ("Ви забули відкрити файл 'Штат'")
    Exit Sub
End If

check_bankbook
If bankfound = vbNullString Then
    MsgBox ("Ви забули відкрити файл 'BANK'")
    Exit Sub
End If

'vybor = MsgBox(prompt:="Перевірка усіх записів триватиме певний час!", Buttons:=vbOKCancel)
'If vybor = 2 Then Exit Sub

startTime = Now()

OptimizeVBA True

Range("R5") = "Працюємо..."
  Range("W3").Value = 0
  Range("W4").Value = 0
  Range("A3:C10000").Clear
  Range("F3:H10000").Clear
DoEvents

booStatusBarState = Application.DisplayStatusBar
Application.DisplayStatusBar = True

ppmax = 0

' набираємо масиви
For Each cWorksheet In Workbooks(wbfound).Worksheets
    If cWorksheet.name Like wsShtat Or cWorksheet.name Like rozpTempl Or cWorksheet.name = "Призуп.в.сл." Or cWorksheet.name = "Загинувші" Or cWorksheet.name = "в 173 бат" Then
        Set parRange = cWorksheet.Range("N2:X10000")
        For Each Line In parRange.Rows
            If Not Line.Hidden And Line.Cells(1) <> vbNullString And StrComp(Line.Cells(1), "ПІБ", 1) <> 0 Then
                PIBs(ppmax) = Line.Cells(1).Text
                IPNs(ppmax) = Line.Cells(11).Text
                sPIBs(ppmax) = clearstring(Line.Cells(1).Text)
                ppmax = ppmax + 1
            End If
        Next
    End If
    If cWorksheet.name = "Звільнені" Or cWorksheet.name = "Переведені" Then
        Set parRange = cWorksheet.Range("P2:Z10000")
        For Each Line In parRange.Rows
            If Not Line.Hidden And Line.Cells(1) <> vbNullString And StrComp(Line.Cells(1), "ПІБ", 1) <> 0 Then
                PIBs(ppmax) = Line.Cells(1).Text
                IPNs(ppmax) = Line.Cells(11).Text
                sPIBs(ppmax) = clearstring(Line.Cells(1).Text)
                ppmax = ppmax + 1
            End If
        Next
    End If
Next
    Application.StatusBar = "0% done..."
    Range("R5").Value = "Працюємо..."
    Interval = Now() - startTime
    Range("R6").Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    DoEvents
'TrimRange bankfound, wsBank, "B4:B10000"
'TrimRange bankfound, wsBank, "C4:C10000"

Workbooks(bankfound).Worksheets(wsBank).AutoFilterMode = False
Set parRange = Workbooks(bankfound).Worksheets(wsBank).Range("A4").Worksheet.UsedRange
m = 0
'mmax = parRange.CurrentRegion.Rows.Count
mmax = 10000

    p = 0
    p2 = 0
    max1 = 0
    max2 = 0

  For Each Line In parRange.Rows

    m = m + 1
    fractionDone = m / mmax
    Application.StatusBar = format(fractionDone, "0%") & " done..."
    Range("R5").Value = "Працюємо... " & format(fractionDone, "0%")
    Interval = Now() - startTime
    Range("R6").Value = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
    DoEvents
    
    tPIB = Line.Cells(2).Text
    sPIB = clearstring(tPIB)
    IPN = Line.Cells(3).Text
    
    If Not Line.Hidden And tPIB <> vbNullString And StrComp(tPIB, "ПІБ", 1) <> 0 Then
    
        founded = 0
        cc = 0
        tfounded = ""
        
                For pp = 0 To ppmax - 1

                    If StrComp(sPIBs(pp), sPIB, vbTextCompare) = 0 Then
                        founded = founded + 1
                        If founded > 1 Then tfounded = tfounded & ", "
                        If IPNs(pp) = "" Then
                            tfounded = tfounded & "порожньо"
                        Else
                            tfounded = tfounded & IPNs(pp)
                        End If
                    End If
                    
                Next
                
                If founded > 0 Then
                    z = z + 1
                    Range("W4").Value = z
                    Line.Cells(2).Interior.Color = RGB(200, 255, 200)
                Else
                    Line.Cells(2).Interior.Color = RGB(255, 100, 100)
                    nz = nz + 1
                    Range("W3").Value = nz
                End If

                If founded > 1 Then
                
                    founded = 0
                    For m1 = 0 To max1 - 1
                        If array1(m1) = sPIB & "-" & IPN Then
                            founded = 1
                            Exit For
                        End If
                    Next
                    If founded = 0 Then
                        Range("A" & 3 + p).Value = tPIB
                        Range("B" & 3 + p).Value = IPN
                        Range("C" & 3 + p).Value = tfounded
                        p = p + 1
                        array1(max1) = sPIB & "-" & IPN
                        max1 = max1 + 1
                    End If
                End If
    
      ' ще треба пошукати ІПН
        founded = 0
        tfounded = ""
                For pp = 0 To ppmax - 1
                
                    If IPN <> "" And IPNs(pp) = IPN Then
                        founded = founded + 1
                        If founded > 1 Then tfounded = tfounded & ", "
                        If PIBs(pp) = "" Then
                            tfounded = tfounded & "порожньо"
                        Else
                            tfounded = tfounded & PIBs(pp)
                        End If
                    End If
                    
                Next
                
                If founded > 1 Or tfounded <> "" And clearstring(tfounded) <> clearstring(tPIB) Then
                    founded = 0
                    For m2 = 0 To max2 - 1
                        If array2(m2) = IPN & "-" & sPIB Then
                            founded = 1
                            Exit For
                        End If
                    Next
                    If founded = 0 Then
                        Range("F" & 3 + p2).Value = IPN
                        Range("G" & 3 + p2).Value = tPIB
                        Range("H" & 3 + p2).Value = tfounded
                        p2 = p2 + 1
                        array2(max2) = IPN & "-" & sPIB
                        max2 = max2 + 1
                    End If
                End If
    Else
        cc = cc + 1
        If tPIB = vbNullString And cc > 10 Then
            Exit For
        End If
    End If
  Next
  Range("W3").Value = nz
  Range("W4").Value = z
  Range("A2:C10000").Sort key1:=Range("A3"), order1:=xlAscending, Header:=xlYes
  
  OptimizeVBA False
  Application.DisplayStatusBar = booStatusBarState
  Application.StatusBar = False

  Range("R5") = "Ok"
  endTime = Now()
  Interval = endTime - startTime
  Range("R6") = Int(CSng(Interval * 24)) & ":" & format(Interval, "nn:ss")
End Sub

' Створено: Окланд, 110 ОМБр, травень 2023



