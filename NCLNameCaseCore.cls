VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NCLNameCaseCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' https://github.com/seagullua/NameCaseLib/blob/master/Library/NCL/NCLNameCaseCore.php
Option Explicit
Private ready As Boolean
Private finished As Boolean
Private words() As NCLNameCaseWord
Private workingWord As String
Private workingLastCache As Variant
Private lastRule As Integer
Private lastResult As Variant
Private index As New Collection
Public gender_koef As Integer


Sub reset()
    lastRule = 0
    lastResult = Array()
End Sub

Sub fullReset()
    Dim n As Integer
    words = Array()
    For n = 1 To index.count
        index.Remove 1
    Next
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

Sub setWorkingWord(Word As String)
    reset
    workingWord = Word
    workingLastCache = Array()
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

Function Last(Optional length As Integer, Optional stopAfter As Integer) As String
    Dim ccut
    If length Is Nothing Then
        length = 1
    End If
    If stopAfter Is Nothing Then
        stopAfter = 0
    End If
    If stopAfter = 0 Then ccut = length Else ccut = stopAfter
    If Not workindLastCache(length, stopAfter) Is Nothing Then
        workindLastCache(length, stopAfter) = NCLStr.substr(workingWord, Len(workingWord) - length, Cut)
    End If
    Last = workindLastCache(length, stopAfter)
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

Sub wordForms(Word As String, endings As Variant, Optional replaceLast As Integer)
    Dim result As Variant
    Dim pagegIndex As Integer
    If replaceLast Is Nothing Then
        replaceLast = 0
    End If
    result = Array(workingWord)
    Word = NCLStr.substr(Word, 0, NCLStr.strlen(Word) - replaceLast)
    For padegIndex = 1 To CaseCount
        result(padegIndex) = Word & endings(padegIndex - 1)
    Next
    lastResult = result
End Sub

Sub setFirstName(Optional firstname As String)
    Dim i As Integer
    Dim obj As NCLNameCaseWord
    If firstname Is Nothing Then
        i = count(words)
        ReDim words(i + 1)
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
        ReDim words(i + 1)
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
        ReDim words(i + 1)
        Set obj = New NCLNameCaseWord
        obj.Init (fathername)
        obj.setNamePart ("F")
        Set words(i) = obj
        notReady
    End If
End Sub

Sub setGender(Optional gender As Integer)
    Dim iword As NCLNameCaseWord
    If gender Is Nothing Then
        gender = 0
    End If
    For Each iword In words
        iword.setTrueGender (gender)
    Next
End Sub

Sub setFullName(secondname As String, firstname As String, fathername As String)
    If secondname Is Nothing Then
        secondname = ""
    End If
    If firstname Is Nothing Then
        firstname = ""
    End If
    If fathername Is Nothing Then
        fathername = ""
    End If
    setFirstName (firstname)
    setSecondName (secondname)
    setFatherName (fathername)
End Sub

Sub setName(Optional firstname As String)
    If firstname Is Nothing Then
        firstname = ""
    End If
    setFirstName (firstname)
End Sub

Sub setLastName(Optional secondname As String)
    If secondname Is Nothing Then
        secondname = ""
    End If
    setSecondName (secondname)
End Sub
' ????
Sub setSirName(Optional fathername As String)
    If fathername Is Nothing Then
        fathername = ""
    End If
    setFatherName (fathername)
End Sub

Sub prepareNamePart(iword As NCLNameCaseWord)
    If iword.getNamePart() <> "" Then
        detectNamePart (iword)
    End If
End Sub

Sub prepareAllNameParts()
    Dim iword
    For Each iword In words
        prepareNamePart (iword)
    Next
End Sub

Sub prepareGender(iword As NCLNameCaseWord)
    Dim namepart As String
    If Not iword.isGenderSolved() Then
        namepart = iword.getNamePart()
        Select Case namepart
            Case "N"
                GenderByFirstName (iword)
            Case "F"
                GenderByFatherName (iword)
            Case "S"
                GenderBySecondName (iword)
        End Select
    End If
End Sub

Function solveGender() As Boolean
    Dim iword As NCLNameCaseWord
    Dim man As Integer, woman As Integer
    For Each iword In words
        If iword.isGenderSolved() Then
            setGender (iword.gender())
            solveGender = True
            Exit Function
        End If
    Next
    man = 0
    woman = 0
    For Each iword In words
        prepareGender (iword)
        gender = iword.getGender()
        man = man + gender(NCL.чоловік)
        woman = woman + gender(NCL.жінка)
    Next
    If (man > woman) Then
        setGender (NCL.чоловік)
    Else
        setGender (NCL.жінка)
    End If
    solveGender = True
End Function

Sub generateIndex()
    Dim m As Integer, tarray()
    Dim namepart As String, Word As NCLNameCaseWord
    index.Add Item:=Array(), Key:="N"
    index.Add Item:=Array(), Key:="S"
    index.Add Item:=Array(), Key:="F"
    For m = 0 To UBound(words)
        Set Word = words(m)
        namepart = Word.getNamePart()
        tarray = index.Item(namepart)
        index.Remove (namepart)
        ReDim tarray(count(tarray) + 1)
        tarray(count(tarray) - 1) = m
        index.Add Item:=tarray, Key:=namepart
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
            End If
        End If
    End If
    genderAutoDetect = False
End Function

Sub splitFullName(fullname As String)
    Dim list() As String, Word, iword As NCLNameCaseWord, formatArray(), n
    fullname = Application.trim(fullname)
    list = Split(fullname, " ")
    n = 0
    For Each Word In list
        n = n + 1
        ReDim words(1 To n)
    Next
    n = 0
    For Each Word In list
        Set iword = New NCLNameCaseWord
        iword.Init (Word)
        n = n + 1
        Set words(n) = iword
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

Sub WordCase(iword As NCLNameCaseWord)
    Dim gender As Integer, namepart As String, name_part_letter As String, method As String
    gender = IIf(iword.gender() = NCL.чоловік, "man", "woman")
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
    result = Array()
    last_rule = -1
    cnt = count(cur_words)
    For k = 0 To cnt - 1
        cur_word = cur_words(k)
        is_norm_rules = True
        o_ncw = New NCLNameCaseWord
        o_ncw.Init (cur_word)
'                    if ( $name_part_letter=='S' && $cnt>1 && $k<$cnt-1 ){
'                        //если первая часть фамилии тоже фамилия, то склоняем по общим правилам
'                        //иначе не склоняется
'                        $exclusion=array('тулуз');//исключения
'                        $cur_word_=mb_strtolower($cur_word);
'                        if ( !in_array($cur_word_, $exclusion ) ){
'                            $o_nc = new NCLNameCaseRu();
'                            $o_nc->detectNamePart( $o_ncw );
'                            $is_norm_rules=( $o_ncw->getNamePart()=='S' );
'                        }
'                        else {
'                            $is_norm_rules=false;
'                        }
'                    }
        setWorkingWord (cur_word)
        If is_norm_rules And Application.Run(method) Then
            result_tmp = lastResult
            last_rule = lastRule
        Else
            result_tmp = array_fill(0, CaseCount, cur_word)
            last_rule = -1
        End If
        o_ncw.setNameCases (result_tmp)
        ReDim o_cur_words(count(o_cur_words) + 1)
        Set o_cur_words(count(o_cur_words) - 1) = o_ncw
    Next
    For Each o_ncw In o_cur_words
        inameCases = o_ncw.getNameCases()
        For k = 0 To count(inameCases)
            namecase = inameCases(k)
            If Not Nothing Is result(k) Then
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
    If Not finished Then
        prepareEverything
        For Each iword In words
            WordCase iword
        Next
        finished = True
    End If
End Sub

Function getWordCase(iword As NCLNameCaseWord, Optional number As Integer) As Variant
    Dim icases() As String
    If Nothing Is number Then
        number = -1
    End If
    icases = iword.getNameCases
    If number < 0 Or number > (CaseCount - 1) Then
        getWordCase = icases
    Else
        getWordCase = icases(number)
    End If
End Function

Function getCasesConnected(indexArray, Optional number = 0)
Dim resultArray(CaseCount - 1), readyArray(), all, m, i, tmp
    For Each index In indexArray
        ReDim Preserve readyArray(0 To UBound(readyArray) - LBound(readyArray) + 1)
        readyArray(UBound(readyArray) - LBound(readyArray)) = getWordCase(words(index), number)
    Next
    all = UBound(readyArray) - LBound(readyArray) + 1
    If all > 0 Then
        If is_array(readyArray(0)) Then
            For m = 0 To CaseCount - 1
                tmp = ""
                For i = 0 To all - 1
                    tmp = tmp & readyArray(i, m) & " "
                Next
                resultArray(m) = Left(tmp, Len(tmp - 1))
            Next
            getCasesConnected = resultArray
            Exit Function
        End If
    Else
        getCasesConnected = Join(readyArray, " ")
        Exit Function
    End If
    getCasesConnected = ""
End Function

                $all = count($readyArr);
                if ($all)
                {
                        if (is_array($readyArr[0]))
                        {
                                //Масив нужно скелить каждый падеж
                                $resultArr = array();
                                for ($case = 0; $case < $this->CaseCount; $case++)
                                {
                                        $tmp = array();
                                        for ($i = 0; $i < $all; $i++)
                                        {
                                                $tmp[] = $readyArr[$i][$case];
                                        }
                                        $resultArr[$case] = implode(' ', $tmp);
                                }
                                return $resultArr;
                        }
                        Else
                        {
                                return implode(' ', $readyArr);
                        }
                }
                Return '';
        }

        /**
         * Функция ставит имя в нужный падеж.
         *
         * Если указан номер падежа <var>$number</var>, тогда возвращается строка с таким номером падежа,
         * если нет, тогда возвращается массив со всеми падежами текущего слова.
         * @param int $number номер падежа
         * @return mixed массив или строка с нужным падежом
         */
        public function getFirstNameCase($number=null)
        {
                $this->AllWordCases();

                return $this->getCasesConnected($this->index['N'], $number);
        }

        /**
         * Функция ставит фамилию в нужный падеж.
         *
         * Если указан номер падежа <var>$number</var>, тогда возвращается строка с таким номером падежа,
         * если нет, тогда возвращается массив со всеми падежами текущего слова.
         * @param int $number номер падежа
         * @return mixed массив или строка с нужным падежом
         */
        public function getSecondNameCase($number=null)
        {
                $this->AllWordCases();

                return $this->getCasesConnected($this->index['S'], $number);
        }

        /**
         * Функция ставит отчество в нужный падеж.
         *
         * Если указан номер падежа <var>$number</var>, тогда возвращается строка с таким номером падежа,
         * если нет, тогда возвращается массив со всеми падежами текущего слова.
         * @param int $number номер падежа
         * @return mixed массив или строка с нужным падежом
         */
        public function getFatherNameCase($number=null)
        {
                $this->AllWordCases();

                return $this->getCasesConnected($this->index['F'], $number);
        }

        /**
         * Функция ставит имя <var>$firstName</var> в нужный падеж <var>$CaseNumber</var> по правилам пола <var>$gender</var>.
         *
         * Если указан номер падежа <var>$CaseNumber</var>, тогда возвращается строка с таким номером падежа,
         * если нет, тогда возвращается массив со всеми падежами текущего слова.
         * @param string $firstName имя, которое нужно просклонять
         * @param int $CaseNumber номер падежа
         * @param int $gender пол, который нужно использовать
         * @return mixed массив или строка с нужным падежом
         */
        public function qFirstName($firstName, $CaseNumber=null, $gender=0)
        {
                $this->fullReset();
                $this->setFirstName($firstName);
                if ($gender)
                {
                        $this->setGender($gender);
                }
                return $this->getFirstNameCase($CaseNumber);
        }

        /**
         * Функция ставит фамилию <var>$secondName</var> в нужный падеж <var>$CaseNumber</var> по правилам пола <var>$gender</var>.
         *
         * Если указан номер падежа <var>$CaseNumber</var>, тогда возвращается строка с таким номером падежа,
         * если нет, тогда возвращается массив со всеми падежами текущего слова.
         * @param string $secondName фамилия, которую нужно просклонять
         * @param int $CaseNumber номер падежа
         * @param int $gender пол, который нужно использовать
         * @return mixed массив или строка с нужным падежом
         */
        public function qSecondName($secondName, $CaseNumber=null, $gender=0)
        {
                $this->fullReset();
                $this->setSecondName($secondName);
                if ($gender)
                {
                        $this->setGender($gender);
                }

                return $this->getSecondNameCase($CaseNumber);
        }

        /**
         * Функция ставит отчество <var>$fatherName</var> в нужный падеж <var>$CaseNumber</var> по правилам пола <var>$gender</var>.
         *
         * Если указан номер падежа <var>$CaseNumber</var>, тогда возвращается строка с таким номером падежа,
         * если нет, тогда возвращается массив со всеми падежами текущего слова.
         * @param string $fatherName отчество, которое нужно просклонять
         * @param int $CaseNumber номер падежа
         * @param int $gender пол, который нужно использовать
         * @return mixed массив или строка с нужным падежом
         */
        public function qFatherName($fatherName, $CaseNumber=null, $gender=0)
        {
                $this->fullReset();
                $this->setFatherName($fatherName);
                if ($gender)
                {
                        $this->setGender($gender);
                }
                return $this->getFatherNameCase($CaseNumber);
        }

        /**
         * Склоняет текущие слова во все падежи и форматирует слово по шаблону <var>$format</var>
         * <b>Формат:</b>
         * - S - Фамилия
         * - N - Имя
         * - F - Отчество
         * @param string $format строка формат
         * @return array массив со всеми падежами
         */
        public function getFormattedArray($format)
        {
                if (is_array($format))
                {
                        return $this->getFormattedArrayHard($format);
                }

                $length = NCLStr::strlen($format);
                $result = array();
                $cases = array();
                $cases['S'] = $this->getCasesConnected($this->index['S']);
                $cases['N'] = $this->getCasesConnected($this->index['N']);
                $cases['F'] = $this->getCasesConnected($this->index['F']);

                for ($curCase = 0; $curCase < $this->CaseCount; $curCase++)
                {
                        $line = "";
                        for ($i = 0; $i < $length; $i++)
                        {
                                $symbol = NCLStr::substr($format, $i, 1);
                                if ($symbol == 'S')
                                {
                                        $line.=$cases['S'][$curCase];
                                }
                                elseif ($symbol == 'N')
                                {
                                        $line.=$cases['N'][$curCase];
                                }
                                elseif ($symbol == 'F')
                                {
                                        $line.=$cases['F'][$curCase];
                                }
                                Else
                                {
                                        $line.=$symbol;
                                }
                        }
                        $result[] = $line;
                }
                return $result;
        }

        /**
         * Склоняет текущие слова во все падежи и форматирует слово по шаблону <var>$format</var>
         * <b>Формат:</b>
         * - S - Фамилия
         * - N - Имя
         * - F - Отчество
         * @param array $format массив с форматом
         * @return array массив со всеми падежами
         */
        public function getFormattedArrayHard($format)
        {

                $result = array();
                $cases = array();
                foreach ($format as $word)
                {
                        $cases[] = $word->getNameCases();
                }

                for ($curCase = 0; $curCase < $this->CaseCount; $curCase++)
                {
                        $line = "";
                        foreach ($cases as $value)
                        {
                                $line.=$value[$curCase] . ' ';
                        }
                        $result[] = trim($line);
                }
                return $result;
        }

        /**
         * Склоняет текущие слова в падеж <var>$caseNum</var> и форматирует слово по шаблону <var>$format</var>
         * <b>Формат:</b>
         * - S - Фамилия
         * - N - Имя
         * - F - Отчество
         * @param array $format массив с форматом
         * @return string строка в нужном падеже
         */
        public function getFormattedHard($caseNum=0, $format=array())
        {
                $result = "";
                foreach ($format as $word)
                {
                        $cases = $word->getNameCases();
                        $result.= $cases[$caseNum] . ' ';
                }
                return trim($result);
        }

        /**
         * Склоняет текущие слова в падеж <var>$caseNum</var> и форматирует слово по шаблону <var>$format</var>
         * <b>Формат:</b>
         * - S - Фамилия
         * - N - Имя
         * - F - Отчество
         * @param string $format строка с форматом
         * @return string строка в нужном падеже
         */
        public function getFormatted($caseNum=0, $format="S N F")
        {
                $this->AllWordCases();
                //Если не указан падеж используем другую функцию
                if (is_null($caseNum) or !$caseNum)
                {
                        return $this->getFormattedArray($format);
                }
                //Если формат сложный
                elseif (is_array($format))
                {
                        return $this->getFormattedHard($caseNum, $format);
                }
                Else
                {
                        $length = NCLStr::strlen($format);
                        $result = "";
                        for ($i = 0; $i < $length; $i++)
                        {
                                $symbol = NCLStr::substr($format, $i, 1);
                                if ($symbol == 'S')
                                {
                                        $result.=$this->getSecondNameCase($caseNum);
                                }
                                elseif ($symbol == 'N')
                                {
                                        $result.=$this->getFirstNameCase($caseNum);
                                }
                                elseif ($symbol == 'F')
                                {
                                        $result.=$this->getFatherNameCase($caseNum);
                                }
                                Else
                                {
                                        $result.=$symbol;
                                }
                        }
                        return $result;
                }
        }

        /**
         * Склоняет фамилию <var>$secondName</var>, имя <var>$firstName</var>, отчество <var>$fatherName</var>
         * в падеж <var>$caseNum</var> по правилам пола <var>$gender</var> и форматирует результат по шаблону <var>$format</var>
         * <b>Формат:</b>
         * - S - Фамилия
         * - N - Имя
         * - F - Отчество
         * @param string $secondName фамилия
         * @param string $firstName имя
         * @param string $fatherName отчество
         * @param int $gender пол
         * @param int $caseNum номер падежа
         * @param string $format формат
         * @return mixed либо массив со всеми падежами, либо строка
         */
        public function qFullName($secondName="", $firstName="", $fatherName="", $gender=0, $caseNum=0, $format="S N F")
        {
                $this->fullReset();
                $this->setFirstName($firstName);
                $this->setSecondName($secondName);
                $this->setFatherName($fatherName);
                if ($gender)
                {
                        $this->setGender($gender);
                }

                return $this->getFormatted($caseNum, $format);
        }

        /**
         * Склоняет ФИО <var>$fullname</var> в падеж <var>$caseNum</var> по правилам пола <var>$gender</var>.
         * Возвращает результат в таком же формате, как он и был.
         * @param string $fullname ФИО
         * @param int $caseNum номер падежа
         * @param int $gender пол человека
         * @return mixed либо массив со всеми падежами, либо строка
         */
        public function q($fullname, $caseNum=null, $gender=null)
        {
                $this->fullReset();
                $format = $this->splitFullName($fullname);
                if ($gender)
                {
                        $this->setGender($gender);
                }

                return $this->getFormatted($caseNum, $format);
        }

        /**
         * Определяет пол человека по ФИО
         * @param string $fullname ФИО
         * @return int пол человека
         */
        public function genderDetect($fullname)
        {
                $this->fullReset();
                $this->splitFullName($fullname);
                return $this->genderAutoDetect();
        }

        /**
         * Возвращает внутренний массив $this->words каждая запись имеет тип NCLNameCaseWord
         * @return array Массив всех слов в системе
         */
        Public Function getWordsArray()
        {
                return $this->words;
        }
