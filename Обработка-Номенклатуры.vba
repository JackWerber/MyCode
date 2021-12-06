' Подсчет определенных символов в ячейке
' =ЕСЛИ(ЕПУСТО(RC[1]);0;ДЛСТР(RC[1])-ДЛСТР(ПОДСТАВИТЬ(ПРОПИСН(RC[1]);",";""))+1)
' Найти позицию второго вхождения буквы "а" в строке "мама мыла раму"
' ПОИСК(",";RC8;ПОИСК(",";RC8)+1)
' Вывод строки до второго разделителя , включительно
' =ЕСЛИ(RC8="";"";ЕСЛИ(ЕОШ(ПОИСК(",";RC8;ПОИСК(",";RC8)+1));RC8;ЛЕВСИМВ(RC8;ПОИСК(",";RC8;ПОИСК(",";RC8)+1))))
' Подсчет ячеек длина которых больше 1
' =ЕСЛИ(СУММПРОИЗВ(Ч(ДЛСТР(R[3]C:R[27]C)>1));СУММПРОИЗВ(Ч(ДЛСТР(R[3]C:R[27]C)>1));"")
' Подсчет ячеек которые не пустые и не равны пробелу
' =СЧЁТЕСЛИМН(R[3]C:R[27]C;"<> ";R[3]C:R[27]C;"<>")

' Создать новый эксель файл в новом Application:
'   Set xlBook1 = xlAPP1.Workbooks.Add(dP.Item("DBPathSrc1"))
' Откроыть существующий файл в новом Application:
'   Set xlBook1 = xlAPP1.Workbooks.Open(Filename:=vCurrentFile, ReadOnly:=False)

' Array.IndexOf(myDictionary.Keys.ToArray(), "a") ' получить индекс ключа в словаре
' .Columns(i).Find("*", .Cells(1, i), , , xlByRows, xlPrevious).Row ' количество заполненных строк в определенном столбце

 ' =ЕСЛИ(ЕПУСТО(RC[-4]);R[-1]C[-4];RC[-4]) ' протяжка
 ' =ЕСЛИ(ЕПУСТО(RC[-4]);ЕСЛИ(ЕПУСТО(R[-1]C[-4]);R[-1]C;R[-1]C[-4]);RC[-4]) ' протяжка
'==============================================================
Function EVAL(strTextString As String)
	Application.Volatile
	EVAL = Evaluate(strTextString)
End Function

' лог процедуры в файл
Public vLogSessionTimeCurr As String
Public vLogSessionTimeLast As String

Function fLogToFile(ByVal vFileName As String, ByVal vTxt As String, Optional ByVal vFSuf As String)
  If vFileName = "this" Then
    vFileName = ThisWorkbook.Path & "\Logs\" & FilenameNoEXT(ThisWorkbook.Name) & "_" & vFSuf & ".txt" ' ThisWorkbook.FullName & "_" & vFSuf & ".txt"
  End If
  Set FileHndl = CreateObject("Scripting.FileSystemObject")
  On Error GoTo lCatch
  If Not FileHndl.FolderExists(ThisWorkbook.Path & "\Logs\") Then
    Set A = FileHndl.CreateFolder(ThisWorkbook.Path & "\Logs\")
  End If
  If vLogSessionTimeCurr = vLogSessionTimeLast Then
    Set vFile = FileHndl.OpenTextFile(vFileName, 8, True) ' 1 for reading, 2 for writing, 8 for appending
  Else
    vLogSessionTimeLast = vLogSessionTimeCurr
    Set vFile = FileHndl.CreateTextFile(vFileName, True)
    vFile.WriteLine " "
    vFile.WriteLine "===== " & vLogSessionTimeCurr & " ====="
  End If

	lFinally:
	  vFile.WriteLine vTxt
	  vFile.Close
	  Set FileHndl = Nothing
		Set vFile = Nothing
	Exit Function
	
	lCatch:
	  MsgBox Err.Description
	  Err.Clear
'	  GoTo lFinally
End Function

' имя файла без расширения
Function FilenameNoEXT(ByVal T As String)
  FilenameNoEXT = T
  R = fFindChr(T, ".", Len(T), 1, "", -1)
  If R Then
    FilenameNoEXT = Left(T, R - 1)
  End If
End Function

' Цифровое число прописью ' Автор MCH (Михаил Ч.), май 2012
Function fSumProp$(chislo#)
  Dim rub$, kop$, ed, des, sot, nadc, razr, i&, m$, vNegate
  vNegate = False
  If chislo >= 1E+15 Then Exit Function
  If chislo < 0 Then
    chislo = Abs(chislo)
    vNegate = True
  End If
  
  sot = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
  des = Array("", "", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
  nadc = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
  ed = Array("", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ", "", "одна ", "две ")
  razr = Array("триллион ", "триллиона ", "триллионов ", "миллиард ", "миллиарда ", "миллиардов ", "миллион ", "миллиона ", "миллионов ", "тысяча ", "тысячи ", "тысяч ", "рубль ", "рубля ", "рублей ")
  
  rub = Left(Format(chislo, "000000000000000.00"), 15)
  kop = Right(Format(chislo, "0.00"), 2)
  
  If CDbl(rub) = 0 Then m = "ноль "
  For i = 1 To Len(rub) Step 3
    If Mid(rub, i, 3) <> "000" Or i = Len(rub) - 2 Then
      m = m & sot(CInt(Mid(rub, i, 1))) & IIf(Mid(rub, i + 1, 1) = "1", nadc(CInt(Mid(rub, i + 2, 1))), _
      des(CInt(Mid(rub, i + 1, 1))) & ed(CInt(Mid(rub, i + 2, 1)) + IIf(i = Len(rub) - 5 And CInt(Mid(rub, i + 2, 1)) < 3, 10, 0))) & _
      IIf(Mid(rub, i + 1, 1) = "1" Or (Mid(rub, i + 2, 1) + 9) Mod 10 >= 4, razr(i + 1), IIf(Mid(rub, i + 2, 1) = "1", razr(i - 1), razr(i)))
    End If
  Next i
  If vNegate Then fSumProp = "- "
  fSumProp = fSumProp & UCase(Left(m, 1)) & Mid(m, 2) & kop & " копе" & IIf(kop \ 10 = 1 Or ((kop + 9) Mod 10) >= 4, "ек", IIf(kop Mod 10 = 1, "йка", "йки"))
End Function


' Четное нечетное число
Function fIsEven(ByVal nNumber As Long) As Boolean
  fIsEven = (nNumber Mod 2) = 0
End Function

' Сравнить два Range
Function fCmprRng(ByVal rng1 As Range, ByVal rng2 As Range) As Boolean
  Dim A As Application
  Set A = Application
  fCmprRng = Join(A.Transpose(A.Transpose(rng1.Value)), Chr(0)) = Join(A.Transpose(A.Transpose(rng2.Value)), Chr(0))
End Function

' Преобразовать Range в строку текста
Function getRangeText(Source As Range, Optional rowDelimiter As String = "@", Optional ColumnDelimiter As String = ",")
  Const CELLLENGTH = 255
  Dim Data()
  Dim Text As String
  Dim BufferSize As Double, length As Double, X As Long, y As Long
  BufferSize = CELLLENGTH * Source.Cells.Count
  Text = Space(BufferSize)

  Data = Source.Value

  For X = 1 To UBound(Data, 1)
    If X > 1 Then
      Mid(Text, length + 1, Len(rowDelimiter)) = rowDelimiter
      length = length + 1
    End If

    For y = 1 To UBound(Data, 2)
      If length + Len(Data(X, y)) + 2 > Len(Text) Then Text = Text & Space(CDbl(BufferSize / 4))
      If y > 1 Then
        Mid(Text, length + 1, Len(ColumnDelimiter)) = ColumnDelimiter
        length = length + 1
      End If

      Mid(Text, length + 1, Len(Data(X, y))) = Data(X, y)
      length = length + Len(Data(X, y))
    Next
  Next
  getRangeText = Left(Text, length) & rowDelimiter
End Function

Function GetSearchArray(ByVal strSearch, ByVal ShtList, ByVal vRange As Range, Optional ByVal ResultRow)
  Dim strResults As String
  Dim sht As Worksheet
  Dim rFND As Range
  Dim sFirstAddress
  For Each sht In ThisWorkbook.Worksheets
    If IsInArray(sht.Index, Split(ShtList, ";")) Then
      Set rFND = Nothing
      With vRange
        Set rFND = .Cells.Find(What:=strSearch, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not rFND Is Nothing Then
          sFirstAddress = rFND.Address
          Do
            If strResults = vbNullString Then
              If ResultRow Then
                strResults = rFND.Row
              Else
                strResults = "Worksheets(" & sht.Index & ").Range(" & Chr(34) & rFND.Address & Chr(34) & ")"
              End If
            Else
              If ResultRow Then
                strResults = strResults & "|" & rFND.Row
              Else
                strResults = strResults & "|" & "Worksheets(" & sht.Index & ").Range(" & Chr(34) & rFND.Address & Chr(34) & ")"
              End If
            End If
            Set rFND = .FindNext(rFND)
          Loop While Not rFND Is Nothing And rFND.Address <> sFirstAddress
        End If
      End With
    End If
  Next
  If strResults = vbNullString Then
    GetSearchArray = Null
  ElseIf InStr(1, strResults, "|", 1) = 0 Then
    GetSearchArray = Array(strResults)
  Else
    GetSearchArray = Split(strResults, "|")
  End If
End Function

Function SplitMultiDelims(ByVal Text As String, ByVal DelimChars As String) As String()
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' SplitMutliChar
  ' This function splits Text into an array of substrings, each substring
  ' delimited by any character in DelimChars. Only a single character
  ' may be a delimiter between two substrings, but DelimChars may
  ' contain any number of delimiter characters. If you need multiple
  ' character delimiters, use the SplitMultiDelimsEX function. It returns
  ' an unallocated array it Text is empty, a single element array
  ' containing all of text if DelimChars is empty, or a 1 or greater
  ' element array if the Text is successfully split into substrings.
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim Pos1 As Long
  Dim N As Long
  Dim m As Long
  Dim arr() As String
  Dim i As Long
  ''''''''''''''''''''''''''''''''
  ' if Text is empty, get out
  ''''''''''''''''''''''''''''''''
  If Len(Text) = 0 Then
    Exit Function
  End If
  ''''''''''''''''''''''''''''''''''''''''''''''
  ' if DelimChars is empty, return original text
  '''''''''''''''''''''''''''''''''''''''''''''
  If DelimChars = vbNullString Then
    SplitMultiDelims = Array(Text)
    Exit Function
  End If
  '''''''''''''''''''''''''''''''''''''''''''''''
  ' oversize the array, we'll shrink it later so
  ' we don't need to use Redim Preserve
  '''''''''''''''''''''''''''''''''''''''''''''''
  ReDim arr(1 To Len(Text))
  i = 0
  N = 0
  Pos1 = 1
  For N = 1 To Len(Text)
    For m = 1 To Len(DelimChars)
      If StrComp(Mid(Text, N, 1), Mid(DelimChars, m, 1), vbTextCompare) = 0 Then
        i = i + 1
        arr(i) = Mid(Text, Pos1, N - Pos1)
        Pos1 = N + 1
        N = N + 1
      End If
    Next m
  Next N
  If Pos1 <= Len(Text) Then
    i = i + 1
    arr(i) = Mid(Text, Pos1)
  End If
  ''''''''''''''''''''''''''''''''''''''
  ' chop off unused array elements
  ''''''''''''''''''''''''''''''''''''''
  ReDim Preserve arr(1 To i)
  SplitMultiDelims = arr
End Function

' Подсчет позиции Nой внешней скобки (если внутри открытой скобки есть подскобка - она не считается)
Function fCBrc(ByVal T, ByVal N, Optional ByVal i, Optional ByVal R, Optional ByVal MODE)
  Dim D
  Set D = CreateObject("Scripting.Dictionary")
  D.Add -1, ")"
  D.Add 1, "("

  LenT = Len(T)
  If LenT = 0 Or i > LenT Then
    fCBrc = "EEROR 1"
    Exit Function
  End If
  If i = "" Then
    i = 1
  End If
  If CInt(MODE) < 1 Or CInt(MODE) > 2 Then
    MODE = 1
  End If
  If CInt(R) <> 1 And CInt(R) <> -1 Then
    R = 1
  End If
  If R = 1 Then
    A = i
    B = LenT
  Else
    A = LenT
    B = i ' здесь надо подумать, может лучше длина минус позиция = I-тая позиция справа
  End If

  If MODE = 1 Then
    vTmp = fCBrc(T, N, i, R, 2)
    If vTmp = 0 Or InStr(vTmp, "ERROR 1") > 0 Or N > vTmp Then
      fCBrc = "0 0"
      Exit Function
    End If
    If N = 0 Then
      N = vTmp
    End If
  End If
  If CInt(N) < 1 Then
    N = 1
  End If

  BC = 0
  fCBrc = "0 0"
  j = 0
  For i = A To B Step R
    C = Mid(T, i, 1)
    If BC = 0 And C = D(R * (-1)) Then
      fCBrc = "ERROR 2"
      Exit Function
    End If
    If C = D(R) Then
      If BC = 0 Then
        j = i
      End If
      BC = BC + 1
    End If
    If C = D(R * (-1)) Then
      BC = BC - 1
      If BC = 0 Then
        If MODE = 1 Then
          NC = NC + 1
          If NC = N Then
            If R = 1 Then
              fCBrc = j & " " & i
            Else
              fCBrc = i & " " & j
            End If
            Exit Function
          End If
        Else
          fCBrc = fCBrc + 1
        End If
      End If
    End If
  Next
  If MODE = 2 Then
    Exit Function
  End If
  fCBrc = "ERROR 3"
End Function

' Функция Сортировка массива
Public Sub QuickSortMultiNaturalNum(strArray As Variant, intBottom As Long, intTop As Long, intSortIndex As Long, Optional intLowIndex As Long, Optional intHighIndex As Long = -1)
  Dim strPivot As String, strTemp As String
  Dim intBottomTemp As Long, intTopTemp As Long
  Dim i As Long
  
  intBottomTemp = intBottom
  intTopTemp = intTop

  If intHighIndex < intLowIndex Then
      If (intBottomTemp <= intTopTemp) Then
          intLowIndex = LBound(strArray, 2)
          intHighIndex = UBound(strArray, 2)
      End If
  End If

  strPivot = strArray((intBottom + intTop) \ 2, intSortIndex)

  While (intBottomTemp <= intTopTemp)
  
  ' < comparison of the values is a descending sort
    While (CompareNaturalNum(strArray(intBottomTemp, intSortIndex), strPivot) < 0 And intBottomTemp < intTop)
      intBottomTemp = intBottomTemp + 1
    Wend
    
    While (CompareNaturalNum(strPivot, strArray(intTopTemp, intSortIndex)) < 0 And intTopTemp > intBottom)
      intTopTemp = intTopTemp - 1
    Wend
    
    If intBottomTemp < intTopTemp Then
      For i = intLowIndex To intHighIndex
        strTemp = Var2Str(strArray(intBottomTemp, i))
        strArray(intBottomTemp, i) = Var2Str(strArray(intTopTemp, i))
        strArray(intTopTemp, i) = strTemp
      Next
    End If
    
    If intBottomTemp <= intTopTemp Then
      intBottomTemp = intBottomTemp + 1
      intTopTemp = intTopTemp - 1
    End If

  Wend
  
  'the function calls itself until everything is in good order
  If (intBottom < intTopTemp) Then QuickSortMultiNaturalNum strArray, intBottom, intTopTemp, intSortIndex, intLowIndex, intHighIndex
  If (intBottomTemp < intTop) Then QuickSortMultiNaturalNum strArray, intBottomTemp, intTop, intSortIndex, intLowIndex, intHighIndex
End Sub

Function CompareNaturalNum(string1 As Variant, string2 As Variant) As Long
'string1 is less than string2 -1
'string1 is equal to string2 0
'string1 is greater than string2 1
Dim n1 As Long, n2 As Long
Dim iPosOrig1 As Long, iPosOrig2 As Long
Dim iPos1 As Long, iPos2 As Long
Dim nOffset1 As Long, nOffset2 As Long

    If Not (IsNull(string1) Or IsNull(string2)) Then
        iPos1 = 1
        iPos2 = 1
        Do While iPos1 <= Len(string1)
            If iPos2 > Len(string2) Then
                CompareNaturalNum = 1
                Exit Function
            End If
            If isDigit(string1, iPos1) Then
                If Not isDigit(string2, iPos2) Then
                    CompareNaturalNum = -1
                    Exit Function
                End If
                iPosOrig1 = iPos1
                iPosOrig2 = iPos2
                Do While isDigit(string1, iPos1)
                    iPos1 = iPos1 + 1
                Loop

                Do While isDigit(string2, iPos2)
                    iPos2 = iPos2 + 1
                Loop

                nOffset1 = (iPos1 - iPosOrig1)
                nOffset2 = (iPos2 - iPosOrig2)

                n1 = Val(Mid(string1, iPosOrig1, nOffset1))
                n2 = Val(Mid(string2, iPosOrig2, nOffset2))

                If (n1 < n2) Then
                    CompareNaturalNum = -1
                    Exit Function
                ElseIf (n1 > n2) Then
                    CompareNaturalNum = 1
                    Exit Function
                End If

                ' front padded zeros (put 01 before 1)
                If (n1 = n2) Then
                    If (nOffset1 > nOffset2) Then
                        CompareNaturalNum = -1
                        Exit Function
                    ElseIf (nOffset1 < nOffset2) Then
                        CompareNaturalNum = 1
                        Exit Function
                    End If
                End If
            ElseIf isDigit(string2, iPos2) Then
                CompareNaturalNum = 1
                Exit Function
            Else
                If (Mid(string1, iPos1, 1) < Mid(string2, iPos2, 1)) Then
                    CompareNaturalNum = -1
                    Exit Function
                ElseIf (Mid(string1, iPos1, 1) > Mid(string2, iPos2, 1)) Then
                    CompareNaturalNum = 1
                    Exit Function
                End If
                iPos1 = iPos1 + 1
                iPos2 = iPos2 + 1
            End If
        Loop
        ' Everything was the same so far, check if Len(string2) > Len(String1)
        ' If so, then string1 < string2
        If Len(string2) > Len(string1) Then
            CompareNaturalNum = -1
            Exit Function
        End If
    Else
        If IsNull(string1) And Not IsNull(string2) Then
            CompareNaturalNum = -1
            Exit Function
        ElseIf IsNull(string1) And IsNull(string2) Then
            CompareNaturalNum = 0
            Exit Function
        ElseIf Not IsNull(string1) And IsNull(string2) Then
            CompareNaturalNum = 1
            Exit Function
        End If
    End If
End Function

Function isDigit(ByVal str As String, pos As Long) As Boolean
Dim iCode As Long
    If pos <= Len(str) Then
        iCode = Asc(Mid(str, pos, 1))
        If iCode >= 48 And iCode <= 57 Then isDigit = True
    End If
End Function

Public Function Var2Str(Value As Variant, Optional TrimSpaces As Boolean = True) As String
    If IsNull(Value) Then
        'Var2Str = vbNullString
        Exit Function
    End If
    If TrimSpaces Then
        Var2Str = Trim(Value)
    Else
        Var2Str = CStr(Value)
    End If
End Function

Sub Test()
Dim Target As Range
Dim vData 'as Variant
Dim Rows As Long
    ' Set Target to the CurrentRegion of cells around "A1"
    Set Target = Split("75, 87, 100, 83, 93, 81, 66, 73, 63, ", ", ") ' Range("A1").CurrentRegion
    ' Copy the values to a variant
    vData = Target.Value2
    ' Get the high/upper limit of the array
    Rows = Target.Rows.Count    'UBound(vData, 1)
    ' Sor The variant array, passing the variant, lower limit, upper limit and the index of the column to be sorted.
    QuickSortMultiNaturalNum vData, 1, Rows, 1
    ' Paste the values back onto the sheet.  For testing, you may want to paste it to another sheet/range
    Range("A1").Resize(Target.Rows.Count, Target.Columns.Count).Value = vData
End Sub

Function BubbleSort(ByRef arr As Variant) As Variant
  Dim strTemp As String
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  lngMin = LBound(arr)
  lngMax = UBound(arr)
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If arr(i) > arr(j) Then
        strTemp = arr(i)
        arr(i) = arr(j)
        arr(j) = strTemp
      End If
    Next j
  Next i
	BubbleSort = arr
End Function

' Процедурка Быстро отсортировать строки
Sub FastFix_sort()
  Dim arr As Variant
  With ActiveSheet
    vColumn = 6
    For vRow = 2 To 14
      arr = Split(.Cells(vRow, vColumn).Value, ", ")
      If Len(Join(arr)) = 0 Then Exit Sub
      .Cells(vRow, vColumn + 1).Interior.ColorIndex = 15
      Cells(vRow, vColumn + 1).Value = Join(BubbleSort(arr), ", ")
    Next
  End With
End Sub

'Описание: удаляет дубли одинаковые значения в массиве!
'DESCRIPTION: Removes duplicates from your array using the collection method.
'NOTES: (1) This function returns unique elements in your array, but
' it converts your array elements to strings.
'SOURCE: https://wellsr.com
'-----------------------------------------------------------------------
Function RemoveDupesColl(MyArray As Variant) As Variant
    Dim i As Long
    Dim arrColl As New Collection
    Dim arrDummy() As Variant
    Dim arrDummy1() As Variant
    Dim item As Variant
    If UBound(MyArray) = (-1) Or Len(Join(MyArray)) = 0 Then
      RemoveDupesColl = Array("")
      Exit Function
    End If
    ReDim arrDummy1(LBound(MyArray) To UBound(MyArray))

    For i = LBound(MyArray) To UBound(MyArray) 'convert to string
        arrDummy1(i) = CStr(MyArray(i))
    Next i
    On Error Resume Next
    For Each item In arrDummy1
       arrColl.Add item, item
    Next item
    Err.Clear
    ReDim arrDummy(LBound(MyArray) To arrColl.Count + LBound(MyArray) - 1)
    i = LBound(MyArray)
    For Each item In arrColl
       arrDummy(i) = item
       i = i + 1
    Next item
    RemoveDupesColl = arrDummy
End Function

' работает плохо - переполняет оперативку! исправить! она нужна
Function findInvisChar(sInput As String) As String
  Dim sSpecialChars As String
  Dim i As Long
  Dim sReplaced As String
  Dim ln As Integer

  sSpecialChars = "" & Chr(1) & Chr(2) & Chr(3) & Chr(4) & Chr(5) & Chr(6) & Chr(7) & Chr(8) & Chr(9) & Chr(10) & Chr(11) & Chr(12) & Chr(13) & Chr(14) & Chr(15) & Chr(16) & Chr(17) & Chr(18) & Chr(19) & Chr(20) & Chr(21) & Chr(22) & Chr(23) & Chr(24) & Chr(25) & Chr(26) & Chr(27) & Chr(28) & Chr(29) & Chr(30) & Chr(31) & Chr(32) & ChrW(&HA0) 'This is your list of characters to be removed
  'For loop will repeat equal to the length of the sSpecialChars string
  'loop will check each character within sInput to see if it matches any character within the sSpecialChars string
  For i = 1 To Len(sSpecialChars)
    ln = Len(sInput) 'sets the integer variable 'ln' equal to the total length of the input for every iteration of the loop
    sInput = Replace$(sInput, Mid$(sSpecialChars, i, 1), "")
    If ln <> Len(sInput) Then sReplaced = sReplaced & Mid$(sSpecialChars, i, 1)
    If ln <> Len(sInput) Then sReplaced = sReplaced & IIf(Mid$(sSpecialChars, i, 1) = Chr(10), "<Line Feed>", Mid$(sSpecialChars, i, 1)) & IIf(Mid$(sSpecialChars, i, 1) = Chr(1), "<Start of Heading>", Mid$(sSpecialChars, i, 1)) & IIf(Mid$(sSpecialChars, i, 1) = Chr(9), "<Character Tabulation, Horizontal Tabulation>", Mid$(sSpecialChars, i, 1)) & IIf(Mid$(sSpecialChars, i, 1) = Chr(13), "<Carriage Return>", Mid$(sSpecialChars, i, 1)) & IIf(Mid$(sSpecialChars, i, 1) = Chr(28), "<File Separator>", Mid$(sSpecialChars, i, 1)) & IIf(Mid$(sSpecialChars, i, 1) = Chr(29), "<Group separator>", Mid$(sSpecialChars, i, 1)) & IIf(Mid$(sSpecialChars, i, 1) = Chr(30), "<Record Separator>", Mid$(sSpecialChars, i, 1)) & IIf(Mid$(sSpecialChars, i, 1) = Chr(31), "<Unit Separator>", Mid$(sSpecialChars, i, 1)) & IIf(Mid$(sSpecialChars, i, 1) = ChrW(&HA0), "<Non-Breaking Space>", Mid$(sSpecialChars, i, 1)) 'Currently will remove all control character but only tell the user about Bell and Line Feed
  Next
  'MsgBox sReplaced & " These were identified and removed"
  findInvisChar = sInput
End Function

' Word предусматривает 5 параметров изменения регистра, а Excel всего три, и то в виде функций.
' Function ConvertRegistr позволяет изменять 5 параметров регистра, аналогично Word.
Function ConvertRegistr(ByVal sString As String, ByVal Tip As Byte) As String
  ' Tip = 1 - ВСЕ ПРОПИСНЫЕ
  ' Tip = 2 - все строчные
  ' Tip = 3 - Начинать С Прописных
  ' Tip = 4 - Как в предложениях
  ' Tip = 5 - иЗМЕНИТЬ рЕГИСТР
  Dim i&
  If Tip = 4 Then
    ConvertRegistr = StrConv(sString, 2)
    Mid$(ConvertRegistr, 1, 1) = UCase(Mid$(ConvertRegistr, 1, 1))
  ElseIf Tip > 4 Then
    For i = 1 To Len(sString)
      Mid$(sString, i, 1) = IIf(Mid$(sString, i, 1) = UCase(Mid$(sString, i, 1)), _
      LCase(Mid$(sString, i, 1)), UCase(Mid$(sString, i, 1)))
    Next
    ConvertRegistr = sString
  Else
    ConvertRegistr = StrConv(sString, Tip)
  End If
End Function

'==============================================================
'                       Строка                   Подстрока              С какой позиции      Какой элемент (0 - последн)   Cимвол остановки               Поиск в обратную сторону - Step?
Function fFindChr(ByVal vString As String, ByVal vPtrn As String, ByVal vI As Integer, ByVal vN As Integer, Optional ByVal vBrk As String, Optional ByVal vRev As Integer, Optional GetWithDelim As Boolean)
  If IsMissing(GetWithDelim) Then GetWithDelim = False
  vDelims = "_,.?!;:\|/<>(){}[]="
  If Len(vPtrn) = 0 Then
    fFindChr = 0
    Exit Function
  End If
  If InStr(vDelims, vPtrn) Then vDelims = Replace(vDelims, vPtrn, "")
  If InStr(vBrk, vPtrn) Then vBrk = Replace(vBrk, vPtrn, "")
  If vI < 1 Then vI = 1
  A = vI
  B = Len(vString)
  If vRev <> 1 And vRev <> -1 Then
    vRev = 1
  Else
    If vRev = -1 Then
      A = vI
      B = 1
    End If
  End If
  If vString = "" Then
    fFindChr = (-1)
    Exit Function
  End If
  Dim vChr As String
  Dim vJ As Integer
  vJ = 0
  fFindChr = 0
  For vI = A To B Step vRev
    vChr = Mid(vString, vI, 1)
    If (InStr(vBrk, vChr) Or vChr Like vBrk) And vBrk <> "" Then
      Exit Function
    End If
    If vChr Like vPtrn Or InStr(vDelims, vChr) Then
      If (InStr(vDelims, vChr) = 0 Or (GetWithDelim And vChr <> " ")) And vChr <> "" Then
        fFindChr = vI
        vJ = vJ + 1
      End If
    End If
    If Not vN = 0 Then
      If vN = vJ Then
        Exit Function
      End If
    End If
  Next
End Function

' Количество элементов в массиве
Function fArrCnt(arr As Variant) As Long
  fArrCnt = UBound(arr) - LBound(arr) + 1
End Function

' Добавить элемент в массив, если массива нет - сначала создает его
Function fPopArr(ByRef arr, ByVal data)
  'a = Array(): ReDim a(size)
  If Not IsArray(arr) Then
    arr = Array()
  End If
  If Len(Join(arr)) = 0 Then
    ReDim Preserve arr(0)
  Else
    ReDim Preserve arr(UBound(arr) + 1)
  End If
  arr(UBound(arr)) = data
End Function

' Наличие любой части строки в массиве
Function PartIsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

' Есть ли элемент в массиве
Function IsInArray(strSearch As String, arr As Variant) As Boolean
  If Len(Join(arr)) = 0 Then
    IsInArray = False
    Exit Function
  End If
  IsInArray = Not IsError(Application.Match(strSearch, arr, False))
End Function

Public Function ColorNom(Cell As Range)
  ColorNom = Cell.Interior.Color
End Function

' Получить текст между двумя позициями, S - сколько добавить символов вокруг, по-умолчанию 0
Function mGetWord(ByVal T, ByVal A, ByVal B, Optional ByVal S)
  A = CInt(A)
  B = CInt(B)
  S = CInt(S)
  Tmp = ""
  R = 0
  If S <> 1 Then
    S = 0
    R = 1
  End If
  If A > B Then
    Tmp = A
    A = B
    B = Tmp
  End If
  If A < 1 Or A > Len(T) Or B < 1 Or B > Len(T) Or Len(T) < (1 + S + S) Or B - A + R < 1 + S Then
    mGetWord = ""
    Exit Function
  Else
    mGetWord = Mid(T, A + S, B - (A + S - R))
  End If
End Function

' Удалить в строке T слово начинающееся с позиции А и заканчивающееся на позиции В, поиск с позиции S, удалется C раз
Function fRemWord(ByVal T, ByVal A, ByVal B, Optional ByVal S, Optional ByVal C)
'  MsgBox Len(T) & " " & A & " " & B & " M=" & Mid(T, A, B - (A - 1)) & " R=" & B - A - 1
  If A < 1 Or B < 1 Or Len(T) < CInt(A) Or Len(T) < CInt(B) Then
    fRemWord = -1
    Exit Function
  End If
  fRemWord = Application.Trim(Replace(T, Mid(T, A, B - (A - 1)), "", S, C))
End Function

Function WordRemove(ByVal str As String, ByVal RemoveWords As String, Optional ByVal vMult As String)
  If vMult = " " Then
    vMult = ""
  ElseIf vMult = "" Then
    vMult = "*"
  End If
  Dim RE As Object
  Set RE = CreateObject("vbscript.regexp")
    If InStr(RemoveWords, "[") Then RemoveWords = Replace(RemoveWords, "[", "\[")
    If InStr(RemoveWords, "]") Then RemoveWords = Replace(RemoveWords, "]", "\]")
  With RE
    .IgnoreCase = True
    .Global = True
    .Pattern = "(?:" & Join(Split(RemoveWords), "|") & ")\s" & vMult ' WorksheetFunction.Trim(RemoveWords)
    If .Test(str) Then
      WordRemove = .Replace(str, "")
    Else
      WordRemove = str
    End If
  End With
End Function

Function DiKey(ByVal D As Dictionary, SI As String) As String
  With D
  For Each Key In .Keys
    If .item(Key) = SI Then
      DiKey = Key
      Exit Function
    End If
  Next
  End With
End Function

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
  Dim sht As Worksheet

  If wb Is Nothing Then Set wb = ThisWorkbook
  On Error Resume Next
  Set sht = wb.Sheets(shtName)
  On Error GoTo 0
  SheetExists = Not sht Is Nothing
End Function

' Главная функция универсального поиска по паттерну в строке      ' str, Ptrn, pMatch, pSubMatch, pGlobal, pCase, pMod
Function fSearchRegEx(ByVal str As String, ByVal Ptrn As String, Optional ByVal pMatch As Integer, Optional ByVal pSubMatch As Integer, Optional ByVal pGlobal As Boolean, Optional ByVal pCase As Boolean, Optional ByVal pMod As Integer) As String
  fSearchRegEx = ""
  If IsMissing(pMatch) Then pMatch = 0
  If IsMissing(pSubMatch) Then pSubMatch = -1
  If IsMissing(pGlobal) Then pGlobal = False
  If IsMissing(pCase) Then pCase = False
  If IsMissing(pMod) Then pMod = 1
  Set RegExp = CreateObject("vbscript.regexp")
  With RegExp
    .IgnoreCase = pCase
    .Global = pGlobal
    .Pattern = Ptrn
    If .Test(str) Then
      Set vFoundstr = .Execute(str)
      If pSubMatch = -1 Then
        If vFoundstr.Count - 1 < pMatch Then Exit Function
        fSearchRegEx = Trim(vFoundstr(pMatch))
      Else
        If vFoundstr(pMatch).submatches().Count - 1 < pSubMatch Then Exit Function
        If pMod = 1 Then fSearchRegEx = Trim(CStr(vFoundstr(pMatch).Value))
        If pMod = 2 Then fSearchRegEx = Trim(CStr(vFoundstr(pMatch).submatches(pSubMatch)))
      End If
    End If
  End With
End Function

' выдает массив, 1 элемент - позиция вхождения, 2 элемент - само вхождение, 3 элемент - позиция следующего вхождения, 4 - следующее вхождение и т.д.
Function fSearchRegExArr(ByVal str As String, ByVal Ptrn As String, Optional ByVal pMatch As Integer, Optional ByVal pSubMatch As Integer, Optional ByVal pCase As Boolean, Optional ByVal pMod As Integer) As Variant
  If IsMissing(pMod) Then pMod = 1
  Dim X() As String
  ReDim X(0)
  bRes = False
  Set RegExp = CreateObject("VBScript.RegExp")
  With RegExp
    .Global = True
    .IgnoreCase = True
    .Pattern = Ptrn
    bRes = .Test(str)
    If bRes Then
      Set oMatches = .Execute(str)
      ReDim X(oMatches.Count * 2 - 1)
      j = 0
      For N = 0 To oMatches.Count - 1
        X(j) = CStr(oMatches(N).FirstIndex + 1)
        If pMod = 1 Then
          X(j + 1) = Mid(str, CStr(oMatches(N).FirstIndex + 1), 1)
        ElseIf pMod = 2 Then
          X(j + 1) = oMatches(N)
        End If
        j = j + 2
      Next
    End If
  End With
  fSearchRegExArr = X
End Function

'==============================================================
' ГТД Мега компаратор милионник
Private Function GetFileContent(ByRef fileStr As String) As Variant
  Dim GetFileContent_() As Variant
  Dim line As String
  ' Open fileStr For Input As #1
  With CreateObject("Scripting.FileSystemObject").GetFile(fileStr).OpenAsTextStream(1)
    i = 0
    ReDim Preserve GetFileContent_(i)
    Do While Not .AtEndOfStream
      GetFileContent_(i) = .ReadLine
      If Int(i / 100000) = i / 100000 Then
        If flagAbort = 6 Then
          Exit Function
        End If
        ' Application.Wait Now + TimeValue("0:00:01") ' Wait for 1 second
      End If
      i = i + 1
      ReDim Preserve GetFileContent_(i)
    Loop
    .Close
  End With
  GetFileContent = GetFileContent_
End Function

Private Function fGetDataParamsFromFile(ByRef arrFile As Variant, ByVal Data As String, ByVal Delim As String) As Variant
  Dim fGetDataParamsFromFile_() As Variant
  i = 0
  ReDim Preserve fGetDataParamsFromFile_(i)
  fGetDataParamsFromFile_(0) = Data
  For Each Row In arrFile
    If Len(Row) Then
      If Split(Row, Delim)(0) = Data Then
        For Each Param In Split(Row, Delim)
          i = i + 1
          ReDim Preserve fGetDataParamsFromFile_(i)
          fGetDataParamsFromFile_(i) = Param
        Next
        Exit For
      End If
    End If
  Next
  fGetDataParamsFromFile = (fGetDataParamsFromFile_)
End Function

Private Function fArrSearch(ByRef DBMain As Variant, _
                            ByVal pSearchMethod As String, _
                            ByVal pSearchData As Dictionary, _
                            Optional ByVal pSPattern As String, _
                            Optional ByVal indexI As Long, _
                            Optional ByVal indexN As Long, _
                            Optional ByVal pStep As Integer, _
                            Optional ByVal pElement As Integer) As Collection
  Dim Delims As String
  vBrkFChr = "_,.?!;:\|/<>(){}[]="
  vBrkFChrP = " _,.?!;:\|/<>(){}[]="
  Dim fArrSort_(3) As Variant
  Dim fArrSearch_(3) As Variant
  Set fArrSearch = New Collection
  vOccurrence = 0
  Set RegExp = CreateObject("vbscript.regexp")
  With RegExp
    .IgnoreCase = True
    .Global = True
    .Pattern = pSPattern
    If pStep <> 1 And pStep <> -1 Then pStep = 1
    If pStep = 1 Then
      If Not indexI > 0 Then indexI = 1
      If Not indexN > 0 Then indexN = UBound(DBMain)
      If indexI > indexN Then indexI = indexN
    Else
      If Not indexI > 0 Then indexI = UBound(DBMain)
      If Not indexN > 0 Then indexN = 1
      If indexI < indexN Then indexI = indexN
    End If
    For i = indexI - 1 To indexN - 1 Step pStep
      vResult = 0
      vFoundSubstr = ""
      vResultParticle = 0
      Select Case pSearchMethod
      ' -------------------------------- '
      Case "Equal"
        For Each vData In pSearchData.item("Val")
          If DBMain(i) = vData Then
            vResult = 1
            vFoundSubstr = vData
            vResultParticle = 100
            Exit For
          End If
        Next
      ' -------------------------------- '
      Case "InStr"
        For Each vData In pSearchData.item("Val")
          vResult = InStr(1, DBMain(i), vData, vbBinaryCompare)
          If vResult > 0 Then
            vFoundSubstr = Mid(DBMain(i), vResult, fFindChr(DBMain(i), "*", vResult, 0, vBrkFChr) - (vResult - 1))
            vResultParticle = 100
            Exit For
          End If
        Next
      ' -------------------------------- '
      Case "Particle"
        For Each vData In pSearchData.item("Val")
          If Len(vData) = 1 Then
            If DBMain(i) = vData Then
              vResult = 1
              vFoundSubstr = vData
              vResultParticle = 100
              Exit For
            End If
          Else
            vResult = InStr(1, DBMain(i), vData, vbBinaryCompare)
            If vResult > 0 Then
              vFoundSubstr = Mid(DBMain(i), vResult, fFindChr(DBMain(i), "*", vResult, 0, vBrkFChrP) - (vResult - 1))
              vResultParticle = 100
              Exit For
            Else
              For j = Len(vData) - 1 To 1 Step -1
                vResult = InStr(1, DBMain(i), Left(vData, j), vbBinaryCompare)
                If CStr(vResult) <> "0" Then
                  vResultParticle = InStr(1, DBMain(i + 1), "|" & Right(vData, Len(vData) - j), vbBinaryCompare)
                  If vResultParticle > 0 And vResultParticle < vResult + Len(vData) Then
                    vFoundSubstr = Left(vData, j) & Right(vData, Len(vData) - j)
                    vResultParticle = 100
                    Exit For
                  Else
                    vResult = 0
                  End If
                End If
              Next
              If vResult = 0 Then
                For K = 1 To 2
                  For j = (Len(vData) - 1) To WorksheetFunction.RoundUp(Len(vData) / 2, 0) Step -1
                    If K = 1 Then
                      vResultParticle = InStr(1, DBMain(i), Left(vData, j), vbBinaryCompare)
                    ElseIf K = 2 Then
                      vResultParticle = InStr(1, DBMain(i), Right(vData, j), vbBinaryCompare)
                    End If
                    If vResultParticle > 0 Then
                      vResult = vResultParticle
                      vFoundSubstr = Mid(DBMain(i), vResult, fFindChr(DBMain(i), "*", vResult, 0, vBrkFChrP) - (vResult - 1))
                      vResultParticle = 100 * (j / Len(vData))
                      GoTo ExitForResultParticle
                    End If
                  Next
                Next
ExitForResultParticle:
              End If
            End If
          End If
          If vResult > 0 Then
            Exit For
          End If
        Next
      ' -------------------------------- '
      Case "Pattern"
        If .Pattern = "" Then
          vResult = 0
        Else
          If .Test(DBMain(i)) Then
            vFoundSubstr = .Execute(DBMain(i))(0)
            vResult = .Execute(DBMain(i)).item(0).FirstIndex + 1
            vResultParticle = "100"
          End If
        End If
      End Select
      ' -------------------------------- '
      If CStr(vResult) <> "0" Then
        fArrSort_(0) = CStr(i + 1)
        fArrSort_(1) = vResult
        fArrSort_(2) = Trim(vFoundSubstr)
        fArrSort_(3) = WorksheetFunction.RoundUp(vResultParticle, 0)
        S = 0
        For Each Sort In pSearchData.item("Ord")
          fArrSearch_(S) = fArrSort_(Sort)
          S = S + 1
        Next
        fArrSearch.Add (fArrSearch_)
        vOccurrence = vOccurrence + 1
      End If
      If vOccurrence = pElement And pElement > 0 Then
        Exit Function
      End If
    Next
  End With
End Function

Private Function fDicCollect(ByRef DBMain As Variant, _
                             ByVal pSearchMethod As String, _
                             ByVal pSearchData As Dictionary, _
                    Optional ByVal pSPattern As String, _
                    Optional ByVal pSpread As Integer, _
                    Optional ByVal indexI As Long, _
                    Optional ByVal indexN As Long, _
                    Optional ByVal pStep As Long, _
                    Optional ByVal pElement As Long, _
                    Optional ByVal pStdOut As String) As Dictionary

  Set fDicCollect = New Dictionary
  Dim arrCollector() As Variant
  Dim cSrchResult As New Collection

  If IsMissing(pSpread) Then pSpread = 0
  If IsMissing(pStep) Then pStep = 1

  If pSearchData.Count = 0 Then Exit Function
  If Len(Join(pSearchData.Items()(0))) = 0 And IsMissing(pSPattern) Then Exit Function
  If Len(Join(pSearchData.Items()(0))) = 0 And pSPattern <> "" And pSearchMethod = "Pattern" Then
    Set cSrchResult = fArrSearch(DBMain, pSearchMethod, pSearchData, pSPattern)
    If cSrchResult.Count > 0 Then
      For Each Result In cSrchResult
        fDicCollect.Add Result(0), Result(1)
        If Len(pStdOut) > 0 Then
          vLog = fLogToFile("this", Result(1), pStdOut)
        End If
      Next
    End If
  Else
    Set cSrchResult = fArrSearch(DBMain, pSearchMethod, pSearchData, pSPattern, indexI, indexN) ' pSpread, не сделал
    If cSrchResult.Count > 0 Then
      For Each Result In cSrchResult
        If Not fDicCollect.Exists(Result(0)) Then fDicCollect.Add Result(0), Result(1) & ";" & Result(2) & ";" & Result(3)
      Next
    End If
  End If
End Function

Private Function fCheckDupes(ByRef dPairs As Variant, ByVal i As Long, ByVal Data As String) As Variant
  If Len(Data) = 0 Or IsInArray("", Split(Data, "|")) Then Exit Function
  Dim arrRowsInSheet_() As Variant
  If dPairs.Exists(Data) Then
    arrRowsInSheet_ = dPairs.item(Data)
    ReDim Preserve arrRowsInSheet_(UBound(arrRowsInSheet_) + 1)
    arrRowsInSheet_(UBound(arrRowsInSheet_)) = i
    dPairs.item(Data) = arrRowsInSheet_()
  Else
    dPairs.Add Data, Array(i)
  End If
  fCheckDupes = dPairs.item(Data)
End Function

Private Function fCollectDupesData(ByRef dPairs As Variant, ByVal i As Long, ByVal Data As String, ByVal pCollect As String) As Variant
  If Len(Data) = 0 Or IsInArray("", Split(Data, "|")) Then Exit Function
  Dim arrRowsInSheet_() As Variant
  If dPairs.Exists(Data) Then
    arrRowsInSheet_ = dPairs.item(Data)
    ReDim Preserve arrRowsInSheet_(UBound(arrRowsInSheet_) + 1)
    arrRowsInSheet_(UBound(arrRowsInSheet_)) = CStr(i) & "|" & pCollect
    dPairs.item(Data) = arrRowsInSheet_()
  Else
    dPairs.Add Data, Array(CStr(i) & "|" & pCollect)
  End If
  fCollectDupesData = dPairs.item(Data)
End Function

' Получить ключ словаря по условию итемов (каждый итем разделен ";") - результат ключ (String)
Private Function fGetCondKeyByItem(ByVal Dic As Dictionary, ByVal Cond As String, ByVal Index As Long) As String
  Dim Val As Long
  Dim ValCheck As Long
  With Dic
    ValCheck = CLng(Split(.Items()(0), ";")(Index))
    fGetCondKeyByItem = .Keys()(0)
    For Each Key In .Keys
      If Key = .Keys()(0) Then GoTo NextFor
      Val = CLng(Split(.item(Key), ";")(Index))
      Select Case Cond
      Case "MaxEq"
        Result = Val >= ValCheck
      Case "Max"
        Result = Val > ValCheck
      Case "MIn"
        Result = Val < ValCheck
      Case Else
        ValCheck = ""
        Exit Function
      End Select
      If Result Then
        ValCheck = Val
        fGetCondKeyByItem = Key
      End If
NextFor:
    Next
  End With
End Function

' ======================================================================================================
Sub C_GetNames()
  MsgBox "МОДУЛЬ " & Chr(13) & Application.VBE.ActiveCodePane.CodeModule.Name & Chr(13) & Chr(13) & "ПРОЦЕДУРА " & Chr(13) & _
  Application.VBE.ActiveVBProject.VBComponents(Application.VBE.ActiveCodePane.CodeModule.Name).CodeModule.ProcOfLine(NumLinProc, 0) & Chr(13) & _
  Chr(13) & "ОШИБКА " & Chr(13) & Err.Description & "(" & Err.Number & ")", vbCritical, "ПРОЕКТ " & Application.CurrentProject.Name
End Sub

'=======================================================================================================
' для файла закупок и прочих - объединение с сохранением значений в каждом столбце
' ======================================================================================================
Sub C_ReMerge()
  vDiConfirm = MsgBox("Ты выбрал нужный диапазон? <ДА> <НЕТ>", vbYesNo, "Переобъединение ячеек") ' 6 ДА, 7 НЕТ 2 ОТМЕНА
  If vDiConfirm <> 6 Then
    MsgBox "Отмена"
    Exit Sub
  End If
  Dim rRange As Range, rMrgRange As Range, wsTempSh As Worksheet, wsActSh As Worksheet
  Application.ScreenUpdating = False: Application.DisplayAlerts = False
  On Error GoTo CatchError
  If ActiveSheet.Name = "BackupReMerge" Then
    MsgBox "Окно бекапа, стоп"
    Exit Sub
  End If
  Set wsActSh = ActiveSheet
  If Not SheetExists("BackupReMerge") Then
    Set wsTempSh = Sheets.Add(, Sheets(Sheets.Count)): wsTempSh.Name = "BackupReMerge"
  Else
    Set wsTempSh = Sheets("BackupReMerge")
  End If
  wsActSh.Activate
  ' Set rRange = wsActSh.UsedRange: rRange.Copy wsTempSh.Range(rRange.Address)
  ' Set rMrgRange = wsTempSh.Range(rRange.Address)
  ' wsActSh.UsedRange.UnMerge
  ' wsActSh.UsedRange.SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
  ' rMrgRange.Copy: rRange.PasteSpecial xlPasteFormats: wsTempSh.Delete
  ' Application.ScreenUpdating = True: Application.DisplayAlerts = True
  Row = wsActSh.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Missing).Row ' макс строка
  col = wsActSh.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Missing).Column ' макс столбик
  For i = 1 To Row
    For y = 1 To col
      ExcelCells = Cells(i, y).Value ' получаем координаты ячейки
      ' проверяем объединенная ли ячейка
      If ExcelCells.MergeCells Then
        CellBeginY = ExcelCells.MergeArea.Row ' индекс строки крайней первой ячеки области
        CellBeginX = ExcelCells.MergeArea.Column ' индекс столбца
        ' индекс строки крайней последней ячейки области
        CellEndY = ExcelCells.MergeArea.Row + ExcelCells.MergeArea.Rows.Count - 1
        ' индекс столбца
        CellEndX = ExcelCells.MergeArea.Column + ExcelCells.MergeArea.Columns.Count - 1
      End If
    Next
  Next
  Exit Sub

CatchError:
    Msg = Chr(13) & "Error #" & str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Err.Description ' ControlChars.CrLf
    Debug.Print Msg
    Err.Clear
End Sub

' ======================================================================================================
' заполнить столбец ЕИ по ТН ВЭД из справочника
' ======================================================================================================
Sub A1_TNVED_UNIT_FROM_DIC()
  Application.ScreenUpdating = False
  Debug.Print String(65535, vbCr)
  Debug.Print "Start " & Now

  Dim dParam: On Error Resume Next
  Set dParam = CreateObject("Scripting.Dictionary")
      dParam.Add "startRow", 2
      dParam.Add "ColTNVED", 11
      dParam.Add "ColImporter", 18
      dParam.Add "ColTNVEDEI", 19
      dParam.Add "ColTransportMethod", 20

      dParam.Add "Importer", "БЕЛАРУСЬ"
      dParam.Add "TransportMethod", 99


  Dim dTNVEDmiss
  Set dTNVEDmiss = CreateObject("Scripting.Dictionary")

  Dim dTNVEDEI
  Set dTNVEDEI = CreateObject("Scripting.Dictionary")
      dTNVEDEI.Add "4010310000", ""
      dTNVEDEI.Add "4010350000", ""
      dTNVEDEI.Add "4010360000", ""
      dTNVEDEI.Add "4010390000", ""
      dTNVEDEI.Add "4016930005", ""
      dTNVEDEI.Add "4016995209", ""
      dTNVEDEI.Add "4016995709", ""
      dTNVEDEI.Add "7315119000", ""
      dTNVEDEI.Add "7320909008", ""
      dTNVEDEI.Add "8409100000", ""
      dTNVEDEI.Add "8409910008", ""
      dTNVEDEI.Add "8409990009", ""
      dTNVEDEI.Add "8413308008", "шт"
      dTNVEDEI.Add "8421230000", "шт"
      dTNVEDEI.Add "8421290009", "шт"
      dTNVEDEI.Add "8421310000", "шт"
      dTNVEDEI.Add "8421392009", "шт"
      dTNVEDEI.Add "8482109008", "шт"
      dTNVEDEI.Add "8483308007", "шт"
      dTNVEDEI.Add "8484100009", ""
      dTNVEDEI.Add "8708309109", ""
      dTNVEDEI.Add "8708803509", ""
      dTNVEDEI.Add "8708809909", ""
      dTNVEDEI.Add "8708949909", ""
      dTNVEDEI.Add "9403109809", ""
      dTNVEDEI.Add "8408201000", "шт"
      dTNVEDEI.Add "8409990001", ""
      dTNVEDEI.Add "6812999009", ""
      dTNVEDEI.Add "8708709109", ""
      dTNVEDEI.Add "8483508000", "шт"
      dTNVEDEI.Add "7326909808", ""
      dTNVEDEI.Add "7326909807", ""
      dTNVEDEI.Add "8708309909", "" ' пошлина 5%
      dTNVEDEI.Add "5911909000", "" ' 08.07.2017

  Dim vSeekRange() As Range

  Set vIntersect = Intersect(ThisWorkbook.Worksheets(1).UsedRange, Columns(dParam.item("ColTNVED"))).Cells
  For Each Cell In vIntersect
    If Cell.Row < dParam.item("startRow") Then
      GoTo SkipFor
    End If
    If Len(CStr(Cell.Value)) > 0 Then
      If dTNVEDEI.Exists(CStr(Cell.Value)) Then
        Cells(Cell.Row, dParam.item("ColTNVEDEI")).Value = dTNVEDEI.item(CStr(Cell.Value))
      Else
        dTNVEDmiss.Add CStr(Cell.Value), ""
        Debug.Print "В строке " & Cell.Row & " ТН ВЭД " & CStr(Cell.Value) & " которого нет в справочнике!"
      End If
    Else
      Debug.Print "В строке " & Cell.Row & " не указан ТН ВЭД!"
    End If

    Cells(Cell.Row, dParam.item("ColImporter")).Value = dParam.item("Importer")

    Cells(Cell.Row, dParam.item("ColTransportMethod")).Value = dParam.item("TransportMethod")
SkipFor:
  Next
  If dTNVEDmiss.Count > 0 Then
    With dTNVEDmiss
      Debug.Print "Добавь ТНВЭД к:"
      vdTxt = "Добавь ТНВЭД к:" & Chr(13)
      For Each vdKey In .Keys
        Debug.Print vdKey
        vdTxt = vdTxt & vdKey & Chr(13)
      Next
    End With
    MsgBox vdTxt
  Else
    MsgBox "Все ТН ВЭД есть в словаре"
  End If
  Debug.Print "Finish " & Now
  Set dParam = Nothing
  Set dTNVEDEI = Nothing
  Set dTNVEDmiss = Nothing
  MsgBox "Завершено"
End Sub

Sub A21_GET_CYR_ByREM()
vLogSessionTimeCurr = CStr(Date) & " " & CStr(Time())
  Debug.Print String(65535, vbCr)
  Debug.Print "Start " & Now

  flagSaveWS = True

  Dim dP
  Set dP = CreateObject("Scripting.Dictionary")
  dP.Add "startRow", 2 ' первая строка с данными

  dP.Add "ColSrc", 30 ' столбец с наименованием деталей кириллицей, движком и OEM
  dP.Add "ColBrand", 1
  dP.Add "ColCode", 2
  dP.Add "ColOEM", 17
  dP.Add "ColTarget", 26 ' куда записывать данные кириллицы
  dP.Add "Sequence", "ColBrand,ColCode,ColOEM"

  Set RegExp = CreateObject("vbscript.regexp")
  RegExp.Global = True
  RegExp.Pattern = " +" '+ означает, что заменяем и два и три и четыре и более пробелов подряд

  With Application
  .EnableEvents = False: .ScreenUpdating = False

  Set vIntersect = Intersect(ThisWorkbook.Worksheets(1).UsedRange, Columns(dP.item("ColTarget"))).Cells
  For Each Cell In vIntersect
    If Cell.Row < dP.item("startRow") Then
      GoTo SkipFor
    End If
    vTmp = Cells(Cell.Row, dP.item("ColSrc")).Value
    For i = 0 To 2
      vTmp = Trim(WordRemove(vTmp, Cells(Cell.Row, dP.item(Split(dP.item("Sequence"), ",")(i))).Value))
    Next

    If Len(vTmp) < Len(Cell.Value) Or Len(Cell.Value) = 0 Then Cell.Value = vTmp

SkipFor:
  Next

  .EnableEvents = True: .ScreenUpdating = True
  DoEvents ' ПРОВЕРИТЬ!!!!!!!!!!!!!!!!!!!! ЕВЕНТЫ ДО ИХ РАЗРЕШЕНИЯ!
  End With

  Debug.Print "Finish " & Now
End Sub

' ======================================================================================================
' разделить поле с наименованием деталей кириллицей, движком и OEM на 2 - в OffsetCyr_Col и в OEM_Col
' ======================================================================================================
Sub A2_GET_CYR_IN_COL()
  vLogSessionTimeCurr = CStr(Date) & " " & CStr(Time())
  Debug.Print String(65535, vbCr)
  Debug.Print "Start " & Now

  flagSaveWS = True

  Dim dParam
  Set dParam = CreateObject("Scripting.Dictionary")
  dParam.Add "Cyr_Col", 3 ' столбец с наименованием деталей кириллицей, движком и OEM
  dParam.Add "OEM_Col", 14 ' куда записывать данные по движкам и OEM без кириллицы
  dParam.Add "OffsetCyr_Col", 13 ' куда записывать данные кириллицы
  dParam.Add "startRow", 2 ' первая строка с данными
  dParam.Add "WSHeaderLength", 23

  Set RegExp = CreateObject("vbscript.regexp")
  RegExp.Global = True
  RegExp.Pattern = " +" '+ означает, что заменяем и два и три и четыре и более пробелов подряд

  ' Параметры отчета

  Dim xlAPP As Excel.Application
  Dim xlBook As Excel.Workbook
  Dim WS As Excel.Worksheet

  If flagSaveWS Then
    Dim dTmp: On Error Resume Next
    Set dTmp = CreateObject("Scripting.Dictionary")
  End If

  ' RegExp замена двойных пробелов
  '==============================================================
  With Application
  .EnableEvents = False: .ScreenUpdating = False

  Set vIntersect = Intersect(ThisWorkbook.Worksheets(1).UsedRange, Columns(dParam.item("Cyr_Col"))).Cells
  For Each Cell In vIntersect
    If Cell.Row < dParam.item("startRow") Then
      GoTo SkipFor
    End If
    ' Определение и выгрузка OEM в отдельный столбец

    ' vResultString = fCyrWSeparate(Cell.Value)

    vPosCyrF = fFindChr(Cell.Value, "*[А-Яа-я]*", 1, 1, , , True)
    vPosCyrL = fFindChr(Cell.Value, "*[А-Яа-я]*", vPosCyrF, 0, "*[A-Za-z0-9]*", 1, True) ' найти последний символ кириллицы до появления A-Z0-9
    vPosRevCyrF = fFindChr(Cell.Value, "*[А-Яа-я]*", Len(Cell.Value), 1, "", -1, True)
    vPosRevCyrL = fFindChr(Cell.Value, "*[А-Яа-я]*", vPosRevCyrF, 0, "*[A-Za-z0-9]*", -1, True)
    If vPosCyrF > 0 Then
      If vPosCyrL < vPosRevCyrF Then
        vLog = fLogToFile("this", "Кириллица в нескольких местах в строке " & Cell.Row & " CYR=" & vPosCyrL & " CYRREV=" & vPosRevCyrF & " SYMB=" & Mid(Cell.Value, vPosRevCyrF, 1) & " VALUE=" & Cell.Value)
      End If
      vAssignType = Trim(RegExp.Replace(Mid(Cell.Value, vPosCyrF, vPosCyrL - (vPosCyrF - 1)), " "))
      vAdditInfo = Trim(Replace(Cell.Value, vAssignType, ""))
    Else
      vAssignType = ""
      vLog = fLogToFile("this", "Нет названия детали " & Cell.Row & " Val=" & Cell.Value)
    End If

    If vAdditInfo = "" Then
      vAdditInfo = "OEM_N/A"
    Else
      vAdditInfo = Trim(RegExp.Replace(Join(SplitMultiDelims(vAdditInfo, "()"), " "), " "))
    End If
    'If Cell.Offset(0, dParam.Item("OEM_Col") - dParam.Item("Cyr_Col")).Value = "" Then
      Cells(Cell.Row, dParam.item("OffsetCyr_Col")) = ConvertRegistr(vAssignType, 4)
      If CLng(dParam.item("OEM_Col")) <> 0 Then
        Cells(Cell.Row, dParam.item("OEM_Col")).Value = ConvertRegistr(vAdditInfo, 1)
      End If
    'Else
    ' Debug.Print Cell.Row & " * ОЕМ уже заполнено"
    'End If

    '==============================================================
    If flagSaveWS Then dTmp.Add vAssignType, ""
    '==============================================================
SkipFor:
  Next

  .EnableEvents = True: .ScreenUpdating = True
  DoEvents ' ПРОВЕРИТЬ!!!!!!!!!!!!!!!!!!!! ЕВЕНТЫ ДО ИХ РАЗРЕШЕНИЯ!
  End With

  'Выгрузка содержимого
  '==============================================================
  If flagSaveWS Then
    Columns(7).HorizontalAlignment = xlLeft
    If Len(dParam.item("WSHeaderLength")) Then Cells(1, CInt(dParam.item("WSHeaderLength"))).Value = "Наименования деталей"
    If dTmp.Count > 0 Then
      For i = 0 To dTmp.Count - 1
        Cells(i + 1, dParam.item("WSHeaderLength")).Value = dTmp.Keys()(i)
      Next
    End If
    dTmp.RemoveAll ' Очистка содержимого контейнера
    dTmp = Nothing

  End If
TestExit:
  Debug.Print "Finish " & Now
End Sub

' не работает! в разработке
Function fCyrWSeparate(ByVal T As String) As Variant
  fCyrWSeparate = Array(T, "")
  Dim RT1 As String
  Dim RT2 As String
  With RegExp
    For Each W In Split(T, " ")
      .IgnoreCase = True
      .Global = True
      .Pattern = ""
      If .Test(W) Then
        vFoundSubstr = .Execute(DBMain(i))(0)
        vResult = .Execute(DBMain(i)).item(0).FirstIndex + 1
      End If
    Next
  End With
  fCyrWSeparate = Array(RT1, RT2)
End Function

'=======================================================================================================
' главная процедура исправлений номенклатуры во входящих ТТН Минска
' ======================================================================================================
Sub A3_PROC_FIX_MINSK()
  vLogSessionTimeCurr = CStr(Date) & " " & CStr(Time())
  Debug.Print String(65535, vbCr)
  Debug.Print "Start " & Now

  vActions = "1;2;5;6;7"
  flagSaveWS = True

  Dim dParam
  Set dParam = CreateObject("Scripting.Dictionary")
  dParam.Add "startRow", 2
  dParam.Add "ColsToFix", 14 ' все столбцы с исходными данными, далее идут генерируемые столбцы которые не надо править
  dParam.Add "IgnoreEmptyCellsInCol", 12 ' игнорить пустые ячейки в столбце
  dParam.Add "ColsToFixRegistr", "Vendor;Origin" ' столбцы в которых делать верхний регистр
  dParam.Add "ColsToTrim", "2;3;4;5;13;14"
  dParam.Add "HaveGTD", 1
  dParam.Add "WSHeaderLength", 2
  dParam.Add "NumberOfColumns", ThisWorkbook.Worksheets(1).UsedRange.Columns.Count
  ' Параметры отчета
  dParam.Add "arrLoadColumnColWidth", "40;16;11;24;24;24"
  dParam.Add "arrLoadColumnNames", "Название детали;Производитель;ТН ВЭД;ГТД;Страна;Импортер"
  dParam.Add "PCsPattern", "1SET=" ' паттерн в котором указано количество единиц в комплекте

  dParam.Add "SuffixList", "-025;-050;-075;-100;-STD;-A;-G;-AG"

  Dim xlAPP As Excel.Application
  Dim xlBook As Excel.Workbook
  Dim WS As Excel.Worksheet

  Dim dType: On Error Resume Next:      Set dType = CreateObject("Scripting.Dictionary")     ' Тип детали
  Dim dVendor:                          Set dVendor = CreateObject("Scripting.Dictionary")   ' Производитель
  Dim dTNVED:                           Set dTNVED = CreateObject("Scripting.Dictionary")    ' ТН ВЭД
  Dim dGTD:                             Set dGTD = CreateObject("Scripting.Dictionary")      ' ГТД
  Dim dOrigin:                          Set dOrigin = CreateObject("Scripting.Dictionary")   ' Страна происхождения
  Dim dImporter:                        Set dImporter = CreateObject("Scripting.Dictionary") ' Импортер
  Dim dTmp:                             Set dTmp = CreateObject("Scripting.Dictionary")      ' * временный словарь

  ' Тип данных по номерам столбцов
  Dim dColNames
  Set dColNames = CreateObject("Scripting.Dictionary")
  dColNames.Add "Index", 1       ' Порядковый номер в накладной
  dColNames.Add "Code", 2        ' Артикул
  dColNames.Add "Type", 3        ' Название (тип) детали
  dColNames.Add "Vendor", 4      ' Производитель
  dColNames.Add "Units", 5       ' Единица измерения
  dColNames.Add "Quantity", 6    ' Кол-во
  dColNames.Add "Price", 7       ' Цена
  dColNames.Add "Sum", 8         ' Сумма
  dColNames.Add "WghtUnit", 9    ' Вес одной единицы
  dColNames.Add "WghtTotal", 10  ' Общий вес
  dColNames.Add "TNVED", 11      ' ТН ВЭД
  dColNames.Add "GTD", 12        ' ГТД
  dColNames.Add "Origin", 13     ' Страна происхождения
  dColNames.Add "OEM", 14        ' Кросс на оригинал, движок, прочее
  dColNames.Add "Name", 15       ' Наименование генерируемое для 1С = "Vendor Type Code (OEM)"
  dColNames.Add "NameFull", 16   ' Полное наименование генерируемое для 1С = "Vendor Type Code"
  dColNames.Add "Comment", 17    ' Комментарий
  dColNames.Add "Importer", 18   ' Импортер
  dColNames.Add "PCs", 21        ' Количество в комплекте
  With dColNames
    For Each vdKey In .Keys
      If Not CInt(.item(vdKey)) = .item(vdKey) Then ' Проверка является ли числом
        MsgBox "Номер столбца указан неверно для " & .item(vdKey), vbInformation, "Ошибка в параметрах: тип данных по номерам столбцов"
        Exit Sub
      End If
      .item(vdKey) = .item(vdKey) - 1
      If IsInArray(CStr(vdKey), Split(dParam.item("ColsToFixRegistr"), ";")) Then dParam.item("ColsToFixRegistr") = Replace(dParam.item("ColsToFixRegistr"), vdKey, .item(vdKey))
    Next
    For Each vItem In Split(dParam.item("ColsToFixRegistr"), ";")
      If Not CInt(vItem) = vItem Then dParam.item("ColsToFixRegistr") = Replace(dParam.item("ColsToFixRegistr"), vItem, -1)
    Next
  End With

  ' RegExp замена двойных пробелов
  Set RegExp = CreateObject("vbscript.regexp")
  RegExp.Global = True
  RegExp.Pattern = " +" '+ означает, что заменяем и два и три и четыре и более пробелов подряд
  ' ===

  With Application
  .EnableEvents = False: .ScreenUpdating = False

  Set vIntersect = Intersect(ThisWorkbook.Worksheets(1).UsedRange, Columns(1)).Cells
  For Each Cell In vIntersect
    If Cell.Row < dParam.item("startRow") Then
      GoTo SkipFor
    End If

    If IsInArray("1", Split(vActions, ";")) Then
      ' Debug.Print Cell.Row & " Проверка нумерации"
      If CInt(Cell.Value) <> (Cell.Row - (dParam.item("startRow") - 1)) Then
        MsgBox "СТОП! Нумерация неверная в строке " & Cell.Row
        Exit Sub
      End If
    End If

    If IsInArray("2", Split(vActions, ";")) Then
      vLog = ""
      ' Debug.Print Cell.Row & " Удаление крайних и двойных пробелов, кривых невидимых символов в указанных столбцах в начальных параметрах"
      For i = 0 To dParam.item("ColsToFix") - 1
        ' Cells.Offset(0, i) = findInvisChar(cell.Offset(0, i))
        If IsInArray(CStr(i + 1), Split(dParam.item("ColsToTrim"), ";")) Then Cell.Offset(0, i) = Trim(RegExp.Replace(Cell.Offset(0, i), " "))
        ' Проверка на пустые строки
        If Len(Cell.Offset(0, i).Value) = 0 And Not IsInArray(CStr(i + 1), Split(dParam.item("IgnoreEmptyCellsInCol"), ";")) Then
          vLog = vLog & DiKey(dColNames, CStr(i)) & " (" & CStr(i + 1) & "), "
        End If
        ' исправить регистр
        If IsInArray(CStr(i), Split(dParam.item("ColsToFixRegistr"), ";")) Then Cell.Offset(0, i) = ConvertRegistr(Cell.Offset(0, i).Value, 1) ' 1 - все прописные
      Next
      If Len(vLog) > 0 Then Call fLogToFile("this", "пустые ячейки " & vLog & "строка " & Cell.Row)
    End If

    If IsInArray("3", Split(vActions, ";")) Then ' не работает
      ' Debug.Print Cell.Row & " Для поршней - проверить правильность написания размеров колец"
      If InStr(UCase(Cell.Offset(0, dColNames.item("Type")).Value), UCase("Поршни")) Or InStr(UCase(Cell.Offset(0, dColNames.item("Type")).Value), UCase("Поршень")) Or InStr(UCase(Cell.Offset(0, dColNames.item("Type")).Value), UCase("Вкладыш")) Then
        vCode = Cell.Offset(0, dColNames.item("Code")).Value
        vCodeSuffix = Split(vCode, "-")
        If IsError(Application.Match(vCodeSuffix(UBound(vCodeSuffix)), Split(dParam.item("SuffixList"), ";"), 0)) Then
          flagSuffixFound = -1
          i = 0
          For Each S In Split(dParam.item("SuffixList"), ";")
            If InStr(vCode, S) And Len(vCode) > 0 Then
              flagSuffixFound = i
            End If
            i = i + 1
          Next
          If flagSuffixFound = -1 Then
            Debug.Print "Отсутствует суффикс < ROW>" & Cell.Row & "< CODE>" & vCode & "< RESULT>" & CStr(flagSuffixFound)
          Else
            vCurrentSuf = Split(dParam.item("SuffixList"), ";")(flagSuffixFound)
            If InStr(vCode, vCurrentSuf) < Len(vCode) - Len(vCurrentSuf) Then
              Debug.Print "Суффикс не в конце < ROW>" & Cell.Row & "< CODE>" & vCode & "< SUF>" & vCurrentSuf & "< SUF NO>" & flagSuffixFound & "< SUF POS>" & InStr(vCode, vCurrentSuf) & "< LEN C-S>" & Len(vCode) & " " & Len(vCurrentSuf)
            Else
              Debug.Print "Исправить суффикс < ROW>" & Cell.Row & "< CODE>" & vCode & "< SUF>" & vCurrentSuf & "< SUF NO>" & flagSuffixFound & "< SUF POS>" & InStr(vCode, vCurrentSuf) & "< LEN C-S>" & Len(vCode) & " " & Len(vCurrentSuf)
            End If
          End If
        End If
      End If
    End If

    If IsInArray("4", Split(vActions, ";")) Then ' должно идти после исправления колец поршней
      ' Debug.Print Cell.Row & " Удалить дубль артикула из OEM"
      vCode = Cell.Offset(0, dColNames.item("Code")).Value
      For Each iCode In Split(dParam.item("SuffixList"), ";")
        If iCode Like Right(vCode, Len(iCode)) Then
          vCode = Left(vCode, Len(vCode) - Len(iCode))
        End If
      Next
      If InStr(Cell.Offset(0, dColNames.item("OEM")).Value, vCode) Then
        Cell.Offset(0, dColNames.item("OEM")).Value = Trim(WordRemove(Cell.Offset(0, dColNames.item("OEM")).Value, vCode))
      End If
    End If

    If IsInArray("5", Split(vActions, ";")) Then
      ' Debug.Print Cell.Row & " Заполнение комплектности, после заполнения OEM, т.к. данные берутся из этой строки"
      vOEM = Cell.Offset(0, dColNames.item("OEM")).Value
'      vPosSet = InStr(vOEM, dParam.Item("PCsPattern"))
'      If vPosSet > 0 Then
'        vPosSetEnd = fFindChr(vOEM, " ", vPosSet, 1)
'        vPCsPattern = Trim(mGetWord(vOEM, vPosSet, vPosSetEnd))
'        vOEM = WordRemove(vOEM, vPCsPattern)
'        Cell.Offset(0, dColNames.Item("PCs")).Value = vPCsPattern
'        Cell.Offset(0, dColNames.Item("OEM")).Value = vOEM
'        ' MsgBox ">" & vPosSet & "<" & Chr(13) & ">" & vPosSetEnd & "<" & Chr(13) & ">" & vPCsPattern & "<" & Chr(13) & ">" & vOEM & "<"
'      Else
        With RegExp
          .IgnoreCase = True
          .Global = True
          .Pattern = "((\s|^)|\d{1}SET=)(\d{1,2}PCS)(\s|$)"
          If .Test(vOEM) Then
            Set vFoundstr = .Execute(vOEM)
            vPrefToRemov = Trim(vFoundstr(0).submatches(0))
            vFoundSubstr = vFoundstr(0).submatches(2)
            vPosSet = InStr(vOEM, vFoundSubstr)
            If vPosSet > 0 Then
              vPosSetEnd = fFindChr(vOEM, " ", vPosSet, 1)
              If vPosSetEnd = 0 Then vPosSetEnd = Len(vOEM)
              vPCsPattern = Trim(mGetWord(vOEM, vPosSet, vPosSetEnd))
              vOEM = WordRemove(Replace(vOEM, vPrefToRemov, ""), vPCsPattern)
              Cell.Offset(0, dColNames.item("PCs")).Value = vPCsPattern
              Cell.Offset(0, dColNames.item("OEM")).Value = vOEM
              ' MsgBox ">" & vPosSet & "<" & Chr(13) & ">" & vPosSetEnd & "<" & Chr(13) & ">" & vPCsPattern & "<" & Chr(13) & ">" & vOEM & "<"
            End If
          End If
          RegExp.Pattern = " +" '+ означает, что заменяем и два и три и четыре и более пробелов подряд
        End With
      'End If
    End If

    If IsInArray("6", Split(vActions, ";")) Then
      ' Debug.Print Cell.Row & " Прописать OEM_N/A в поле OEM если отсутствует доп инфа из Минского артикула"
      'If cell.Offset(0, 2).Value Like "*[A-Za-z]*" Then
      '    vLog = fLogToFile("this", "ЯЧЕЙКА С ЛАТИНИЦЕЙ R " & cell.Row & " C5")
      'End If
      vOEM = Cell.Offset(0, dColNames.item("OEM")).Value
      If vOEM = "" Then Cell.Offset(0, dColNames.item("OEM")).Value = "OEM_N/A"
    End If

    If IsInArray("7", Split(vActions, ";")) Then
      vOEM = Cell.Offset(0, dColNames.item("OEM")).Value
      ' Debug.Print Cell.Row & " Наименование для 1С <Производитель тип_детали артикул (OEM)>"
      Cell.Offset(0, dColNames.item("Name")).Value = Cell.Offset(0, dColNames.item("Vendor")).Value & " " & Cell.Offset(0, dColNames.item("Type")).Value & " " & Cell.Offset(0, dColNames.item("Code")).Value & " (" & vOEM & ")"
      ' Полное наименование для 1С "Производитель артикул"
      Cell.Offset(0, dColNames.item("NameFull")).Value = Cell.Offset(0, dColNames.item("Vendor")).Value & " " & Cell.Offset(0, dColNames.item("Code")).Value
      ' Вставка комментария, например: Загрузка 03.02.2017 03:27:18 ТТН 3929118 №1
      Cell.Offset(0, dColNames.item("Comment")).Value = "Загрузка " & CStr(Date) & " " & CStr(Time()) & " ТТН " & ThisWorkbook.Worksheets(7).Cells(2, 2).Value & " №" & Cell.Value
    End If

    If flagSaveWS Then
      ' Debug.Print Cell.Row & " Выгрузка в словари"
      dType.Add Cell.Offset(0, dColNames.item("Type")).Value, ""
      dVendor.Add Cell.Offset(0, dColNames.item("Vendor")).Value, ""
      dTNVED.Add Cell.Offset(0, dColNames.item("TNVED")).Value, ""
      If dParam.item("HaveGTD") = 1 Then dGTD.Add Cell.Offset(0, dColNames.item("GTD")).Value, ""
      dOrigin.Add Cell.Offset(0, dColNames.item("Origin")).Value, ""
      dImporter.Add Cell.Offset(0, dColNames.item("Importer")).Value, ""
    End If

    'If cell.Row > 3 Then
    ' GoTo TestExit
    'End If
SkipFor:
  Next
  If dParam.item("HaveGTD") = 0 Then dGTD.Add "-", ""

  .EnableEvents = True: .ScreenUpdating = True
  DoEvents
  End With

  ' проверка что еще остались данные по комплектности в столбце OEM
  If IsInArray("5", Split(vActions, ";")) Then
    For i = dParam.item("startRow") To ThisWorkbook.Worksheets(1).UsedRange.Rows.Count
      If InStr(Cells(i, dColNames.item("OEM")).Value, "PC") Then
        Debug.Print "В строке " & CStr(i) & " остались данные по комплектности (в столбце OEM)  |  " & Cells(i, dColNames.item("OEM")).Value
        Call fLogToFile("this", "В строке " & CStr(i) & " остались данные по комплектности (в столбце OEM)  |  " & Cells(i, dColNames.item("OEM")).Value)
      End If
    Next
  End If

  'Проход по 3 массивам с выгрузкой содержимого в новую книгу
  '==============================================================
  If flagSaveWS Then
    Set xlAPP = CreateObject("Excel.Application")
    Set xlBook = xlAPP.Workbooks.Add
    Set WS = xlBook.Worksheets(1)

    For i = 1 To 6
      Select Case i
        Case 1
          Set dTmp = dType
        Case 2
          Set dTmp = dVendor
        Case 3
          Set dTmp = dTNVED
        Case 4
          Set dTmp = dGTD
        Case 5
          Set dTmp = dOrigin
        Case 6
          Set dTmp = dImporter
      End Select

      WS.Columns(i).ColumnWidth = Split(dParam.item("arrLoadColumnColWidth"), ";")(i - 1)
      WS.Cells(1, i).HorizontalAlignment = xlCenter
      WS.Cells(1, i).Font.Bold = True
      WS.Cells(1, i).Value = Split(dParam.item("arrLoadColumnNames"), ";")(i - 1)

      If dTmp.Count > 0 Then ' построчное заполнение столбца
        For j = 0 To dTmp.Count - 1
          WS.Cells(j + dParam.item("WSHeaderLength"), i) = dTmp.Keys()(j)
        Next
      End If
      With WS ' сортировка столбца
        TmpLastRow = .Columns(i).Find("*", .Cells(1, i), , , xlByRows, xlPrevious).Row ' количество заполненных строк в определенном столбце
        Set RngTmp = .Range(.Cells(CInt(dParam.item("WSHeaderLength")), i), .Cells(TmpLastRow, i))
        RngTmp.Sort Key1:=RngTmp(1, 1), Order1:=xlAscending, Header:=xlYes
      End With
      dTmp.RemoveAll ' Очистка содержимого контейнера
    Next
    dType = Nothing
    dVendor = Nothing
    dTNVED = Nothing
    dGTD = Nothing
    dOrigin = Nothing
    dImporter = Nothing
    dTmp = Nothing

    ' Сохранение новой книги
    '==============================================================
    MsgBox "Saving at " & ThisWorkbook.Path & "\" & FilenameNoEXT(ThisWorkbook.Name) & "-LoadData.xls"
    '51 = xlOpenXMLWorkbook (without macro's in 2007-2016, xlsx)
    '52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2016, xlsm)
    '50 = xlExcel12 (Excel Binary Workbook in 2007-2016 with or without macro's, xlsb)
    '56 = xlExcel8 (97-2003 format in Excel 2007-2016, xls)
    xlBook.SaveAs Filename:=ThisWorkbook.Path & "\" & FilenameNoEXT(ThisWorkbook.Name) & "-LoadData.xls", FileFormat:=56
    ' xlAPP.Visible = True
    xlBook.Close SaveChanges:=False
    xlAPP.Quit
  End If
TestExit:
  Debug.Print "Finish " & Now
End Sub

' ======================================================================================================
' МегаКомпаратор ГТД милионник
' ======================================================================================================
Sub A4_COMPARATOR_GTD_MILION()
  vLogSessionTimeCurr = CStr(Date) & " " & CStr(Time())
  Debug.Print String(65535, vbCr)
  Debug.Print Now & " * Start " & Application.VBE.ActiveCodePane.CodeModule
  Call fLogToFile("this", Now & " * Start " & Application.VBE.ActiveCodePane.CodeModule, "GTD_Comp")
  Dim vGTD As Long
  Dim DBMain As Variant
  Dim DBAdd1 As Variant
  Dim LibraryAdd1 As Variant
  ' Nest Dictionary Collectors
  Set dNestDicC = CreateObject("Scripting.Dictionary")
  Set dInnrDicC = CreateObject("Scripting.Dictionary")
  Set dTrnsDicC = CreateObject("Scripting.Dictionary")
  Set pSearchData = CreateObject("Scripting.Dictionary")
      pSearchData.Add "Val", Array()
      pSearchData.Add "Pref", ""
      pSearchData.Add "Ord", Array()
  ' Dictionary Vendor-Code pairs
  Set dPairs = New Dictionary
  ' Parameters General
  Set dPrmGnrl = New Dictionary
      dPrmGnrl.Add "DBMainPath", "S:\SharePoint\Документы\5-Поступление товара\_1ГТД\DB-GTD.txt" ' -Test
      dPrmGnrl.Add "DBAdd1Path", "S:\SharePoint\Документы\5-Поступление товара\_1ГТД\DB-VendorSyn.txt"
  ' ===================================================================
      dPrmGnrl.Add "StdOutGTD", "GTD_DUMP" ' !!!!!!!!!!!!!!!!!!!!!!!!!!
  ' ===================================================================
      dPrmGnrl.Add "StartRow", 2
      dPrmGnrl.Add "EndRow", Cells.Find("*", [A1], , , xlByRows, xlPrevious).Row
  ' Parameters of Dictionary Collectors
  Set dPrmDicC = CreateObject("Scripting.Dictionary")
      dPrmDicC.Add "ColumnToSearch", Array("", 4, 2)
      dPrmDicC.Add "ColumnToFill", Array("", 32, 34)
      dPrmDicC.Add "ColumnGTD", 12
      dPrmDicC.Add "Name", Array("DicGTD", "DicVendor", "DicCode")
      dPrmDicC.Add "SearchMethod", Array("Pattern", "InStr", "Particle")
      dPrmDicC.Add "Pattern", Array("[0-9]{5,9}/[0123][0-9][01][0-9][01][0-9]/[A-Za-zА-Яа-я]?[0-9]{5,8}", "", "")
      dPrmDicC.Add "Spread", Array(0, 0, 1)

  With Application
    .EnableEvents = False: .ScreenUpdating = False
    Set WS = ThisWorkbook.Sheets(1)
    With WS
      Debug.Print Now & " * Reading DB Main from " & dPrmGnrl.item("DBMainPath")
      Call fLogToFile("this", Now & " * Reading DB Main from " & dPrmGnrl.item("DBMainPath"), "GTD_Comp")
      DBMain = GetFileContent(dPrmGnrl.item("DBMainPath"))
      DBAdd1 = GetFileContent(dPrmGnrl.item("DBAdd1Path"))

'  ---------------------- Nest DicGTD -> Row -> GTD  ----------------------
      Debug.Print Now & " * fDicCollect() " & dPrmDicC.item("Name")(0)
      Call fLogToFile("this", Now & " * fDicCollect() " & dPrmDicC.item("Name")(0), "GTD_Comp")
      Set dInnrDicC = New Dictionary
      pSearchData.item("Val") = Array()
      pSearchData.item("Pref") = ""
      pSearchData.item("Ord") = Array(0, 2, 1, 3)
      Set dInnrDicC = fDicCollect(DBMain, dPrmDicC.item("SearchMethod")(0), pSearchData, dPrmDicC.item("Pattern")(0), , , , , , dPrmGnrl.item("StdOutGTD"))
      If Len(dPrmGnrl.item("StdOutGTD")) > 0 Then
        Debug.Print Now & " * " & dPrmDicC.item("Name")(0) & " ONLY STDOUT GTD. Stop"
        Call fLogToFile("this", Now & " * " & dPrmDicC.item("Name")(0) & " ONLY STDOUT GTD", "GTD_Comp")
        GoTo ExitSub
      End If
      If dInnrDicC.Count > 0 Then
        dNestDicC.Add dPrmDicC.item("Name")(0), dInnrDicC
      Else
        Debug.Print Now & " * " & dPrmDicC.item("Name")(0) & " no GTD found in DB. Stop"
        Call fLogToFile("this", Now & " * " & dPrmDicC.item("Name")(0) & " no GTD found in DB. Stop", "GTD_Comp")
        GoTo ExitSub
      End If

      Set dInnrDicC = New Dictionary
      dNestDicC.Add dPrmDicC.item("Name")(1), dInnrDicC
      dNestDicC.Add dPrmDicC.item("Name")(2), dInnrDicC

      Debug.Print Now & " * fDicCollect() ROWS=" & dPrmGnrl.item("StartRow") & "-" & dPrmGnrl.item("EndRow") & " " & dPrmDicC.item("Name")(1) & " & " & dPrmDicC.item("Name")(2)
      Call fLogToFile("this", Now & " * fDicCollect() ROWS=" & dPrmGnrl.item("StartRow") & "-" & dPrmGnrl.item("EndRow") & " " & dPrmDicC.item("Name")(1) & " & " & dPrmDicC.item("Name")(2), "GTD_Comp")
      For i = dPrmGnrl.item("StartRow") To dPrmGnrl.item("EndRow")
        If Len(.Cells(i, dPrmDicC.item("ColumnGTD")).Value) Then
          Debug.Print Now & " * row " & i & " GTD not empty"
          Call fLogToFile("this", Now & " * row " & i & " GTD not empty", "GTD_Comp")
          GoTo MarkSkipFor
        End If
        vDC1 = .Cells(i, dPrmDicC.item("ColumnToSearch")(1)).Value
        vDC2 = .Cells(i, dPrmDicC.item("ColumnToSearch")(2)).Value
        Set dInnrDicC = New Dictionary
'  ---------------------- Nest DicVendor -> Vendor -> Dic(Rows) ----------------------
        ' Debug.Print Now & " * " & vDC1 & " " & vDC2
        Call fLogToFile("this", Now & " * (" & CStr(i) & ") " & vDC1 & " " & vDC2, "GTD_Comp")
        .Cells(i, dPrmDicC.item("ColumnToFill")(1) + 1).Value = vDC1
        ' Set pSearchData Get Vendor synonyms
        pSearchData.item("Val") = fGetDataParamsFromFile(DBAdd1, vDC1, ";")
        pSearchData.item("Pref") = ""
        pSearchData.item("Ord") = Array(0, 1, 2, 3)

        If dNestDicC(dPrmDicC.item("Name")(1)).Exists(vDC1) Then GoTo MarkGetCode
        Set dInnrDicC = fDicCollect(DBMain, dPrmDicC.item("SearchMethod")(1), pSearchData, dPrmDicC.item("Pattern")(1), dPrmDicC.item("Spread")(1))

        If dInnrDicC.Count > 0 Then
          dNestDicC(dPrmDicC.item("Name")(1)).Add vDC1, dInnrDicC
          .Cells(i, dPrmDicC.item("ColumnToFill")(1) + 1).Value = vDC1
          vTxt = ""
          For Each Tmp In dNestDicC(dPrmDicC.item("Name")(1))(vDC1).Keys
            vTxt = vTxt & Split(Tmp, ";")(0) & ";"
            ' Debug.Print dNestDicC(dPrmDicC.Item("Name")(1))(vDC1).Item(tmp) & " 0=" & Split(tmp, ";")(0)
          Next
          .Cells(i, dPrmDicC.item("ColumnToFill")(1)).Value = vTxt
        Else
          Debug.Print Now & " * " & dPrmDicC.item("Name")(1) & " " & vDC1 & " not found"
          Call fLogToFile("this", Now & " * " & dPrmDicC.item("Name")(1) & " " & vDC1 & " not found", "GTD_Comp")
          GoTo MarkSkipFor
        End If

'  ---------------------- Nest DicCode -> Vendor|Code -> Dic(Rows) ----------------------
MarkGetCode:
        With dPrmDicC
          WS.Cells(i, .item("ColumnToFill")(2) + 1).Value = vDC2
          ' Debug.Print vDC1 & " " & vDC2
          ' Set pSearchData Get Codes
          pSearchData.item("Val") = Array(vDC2)
          pSearchData.item("Pref") = vDC1
          pSearchData.item("Ord") = Array(0, 3, 2, 1)
          vDupe = fCheckDupes(dPairs, i, vDC1 & "|" & vDC2)
          If UBound(vDupe) < 1 Then
            vTxt = ""
            Set dTrnsDicC = New Dictionary
            For Each vDC2Key In dNestDicC(.item("Name")(1))(vDC1).Keys
              Set dInnrDicC = New Dictionary
              Set dInnrDicC = fDicCollect(DBMain, .item("SearchMethod")(2), pSearchData, .item("Pattern")(2), .item("Spread")(2), vDC2Key - 1, vDC2Key + 1)
              If dInnrDicC.Count > 0 Then
                For Each Tmp In dInnrDicC.Keys
                  If Not dTrnsDicC.Exists(Tmp) Then
                    dTrnsDicC.Add Tmp, dInnrDicC.item(Tmp)
                    vTxt = vTxt & Tmp & ";"
                    'Debug.Print "row " & tmp & "; " & _
                    '            Split(dInnrDicC.Item(tmp), ";")(0) & "%; " & _
                    '            Split(dInnrDicC.Item(tmp), ";")(1) & "; " & _
                    '            Split(dInnrDicC.Item(tmp), ";")(2) ' dNestDicC(.Item("Name")(2))(vDC1 & "|" & vDC2).Item(tmp) & " A=" & tmp & " 0=" & Split(tmp, ";")(0)
                  End If
                Next
              End If
            Next
            WS.Cells(i, .item("ColumnToFill")(2)).Value = vTxt
            dNestDicC(.item("Name")(2)).Add vDC1 & "|" & vDC2, dTrnsDicC

'  ---------------------- Get position of latest best match of codes ----------------------
            If dNestDicC(dPrmDicC.item("Name")(2))(vDC1 & "|" & vDC2).Count > 0 Then
              For Each vDC2Key In dNestDicC(dPrmDicC.item("Name")(2))(vDC1 & "|" & vDC2).Keys
                Tmp = Tmp & "row " & vDC2Key & "; " & _
                             Split(dNestDicC(dPrmDicC.item("Name")(2))(vDC1 & "|" & vDC2).item(vDC2Key), ";")(0) & "%; " & _
                             Split(dNestDicC(dPrmDicC.item("Name")(2))(vDC1 & "|" & vDC2).item(vDC2Key), ";")(1) & "; " & _
                             Split(dNestDicC(dPrmDicC.item("Name")(2))(vDC1 & "|" & vDC2).item(vDC2Key), ";")(2) & Chr(13) ' dNestDicC(.Item("Name")(2))(vDC1 & "|" & vDC2).Item(tmp) & " A=" & tmp & " 0=" & Split(tmp, ";")(0)
              Next
              vMatchCodePos = CLng(fGetCondKeyByItem(dNestDicC(dPrmDicC.item("Name")(2))(vDC1 & "|" & vDC2), "MaxEq", 0))
              ' Debug.Print tmp & vMatchCodePos

'  ---------------------- Get GTD from best match ----------------------
              If vMatchCodePos > 0 Then
                vGTD = 0
                For Each KeyGTD In dNestDicC(dPrmDicC.item("Name")(0)).Keys
                  If IsNumeric(KeyGTD) Then
                    If CLng(KeyGTD) <= vMatchCodePos Then
                      vGTD = CLng(KeyGTD)
                    Else
                      Exit For
                    End If
                  Else
                    Call fLogToFile("this", Now & " * ЖОПА С РУЧКОЙ " & KeyGTD, "GTD_Comp")
                  End If
                Next
                WS.Cells(i, dPrmDicC.item("ColumnToFill")(2) + 2).Value = "R=" & vMatchCodePos & " P=" & Split(dNestDicC(dPrmDicC.item("Name")(2))(vDC1 & "|" & vDC2).item(CStr(vMatchCodePos)), ";")(0) & "% S=" & _
                                                                                                         Split(dNestDicC(dPrmDicC.item("Name")(2))(vDC1 & "|" & vDC2).item(CStr(vMatchCodePos)), ";")(1)
                WS.Cells(i, dPrmDicC.item("ColumnGTD")).Value = dNestDicC(dPrmDicC.item("Name")(0)).item(CStr(vGTD))
              End If
            End If
          Else
            ' Copy duplicate info into Row
            WS.Cells(i, dPrmDicC.item("ColumnToFill")(1)).Value = WS.Cells(vDupe(0), dPrmDicC.item("ColumnToFill")(1)).Value
            WS.Cells(i, dPrmDicC.item("ColumnToFill")(2)).Value = WS.Cells(vDupe(0), dPrmDicC.item("ColumnToFill")(2)).Value
            WS.Cells(i, dPrmDicC.item("ColumnGTD")).Value = WS.Cells(vDupe(0), dPrmDicC.item("ColumnGTD")).Value
            Debug.Print Now & " * Dupe found (" & CStr(i) & ") " & vDC1 & "|" & vDC2 & "  at " & Join(dPairs.item(vDC1 & "|" & vDC2), ", ")
            Call fLogToFile("this", Now & " * Dupe found (" & CStr(i) & ") " & vDC1 & "|" & vDC2 & "  at " & Join(dPairs.item(vDC1 & "|" & vDC2), ", "), "GTD_Comp")
          End If
        End With

MarkSkipFor:
      Next

ExitSub:
    End With
    .EnableEvents = True: .ScreenUpdating = True
    ' DoEvents
  End With

  Set dPrmDicC = Nothing
  Debug.Print Now & " * Finish"
  Call fLogToFile("this", Now & " * Finish", "GTD_Comp")
End Sub

'=======================================================================================================
Function fSearcher(ByVal S As String, ByVal arr As Variant, Optional ByVal vCASE As Integer, Optional ByVal pMod As Integer, Optional ByVal pDelimCase As String) As Collection
  If IsMissing(pMod) Then pMod = 1
  If IsMissing(pDelimCase) Or pDelimCase = "" Then pDelimCase = " "
  Set fSearcher = New Collection
  Dim tmpArr_ As Variant
  Dim vPtrn As String
  Dim A As String
  Dim B As String
  For i = LBound(arr) To UBound(arr)
    A = CStr(S)
    B = CStr(arr(i, 1))

    'If Len(A) = 0 Or Len(B) = 0 Then
    '  GoTo SkipForFSearcher
    'End If
    ' If Not InStr(A, " ") And Not InStr(B, " ") Then pMod = 1
      If i > 25 Then
        fuckcheck = 1
      End If
    ' --------------------------------
    Select Case pMod
    ' --------------------------------
      Case 1
        If vCASE = 1 Then
          If A = B Then
            fSearcher.Add (i)
          End If
        Else
          If UCase(A) = UCase(B) Then
            fSearcher.Add (i)
          ElseIf vCASE > 2 Then
            If Replace(Replace(UCase(A), "/", ""), "-", "") = Replace(Replace(UCase(B), "/", ""), "-", "") Then
              fSearcher.Add (i)
            ElseIf vCASE > 3 Then
              If (UCase(A) Like "*-[A-Za-z]0" Or UCase(A) Like "*-[A-Za-z]O") And (UCase(B) Like "*-[A-Za-z]0" Or UCase(B) Like "*-[A-Za-z]O") And Left(UCase(A), Len(UCase(A)) - 1) = Left(UCase(B), Len(UCase(B)) - 1) Then
                fSearcher.Add (i)
              ElseIf vCASE > 4 Then
                If InStr(UCase(A), UCase(B)) Or InStr(UCase(B), UCase(A)) Then
                  fSearcher.Add (i)
                End If
              End If
            End If
          End If
        End If
      ' --------------------------------
      Case 2
        tmpArr_ = Split(A, pDelimCase)
        For j = 0 To UBound(tmpArr_)
          vPtrn = CStr(tmpArr_(j))
          If vCASE = 1 Then
            If Not IsInArray(vPtrn, Split(B, pDelimCase)) Then GoTo SkipForFSearcher
          ElseIf vCASE = 2 Then
            If Not IsInArray(Replace(Replace(UCase(vPtrn), "/", ""), "-", ""), Split(Replace(Replace(UCase(B), "/", ""), "-", ""), pDelimCase)) Then GoTo SkipForFSearcher
          ElseIf vCASE = 3 Then
            If Not PartIsInArray(vPtrn, Split(B, pDelimCase)) Then GoTo SkipForFSearcher
          End If
        Next
        fSearcher.Add (i)
      ' --------------------------------
      Case 3
        If InStr(UCase(A), UCase(B)) Or InStr(UCase(B), UCase(A)) > 0 Then
          fSearcher.Add (i)
        End If
      ' --------------------------------
      Case 4
        If vCASE = 1 Then
          If IsInArray(B, Split(A, pDelimCase)) Then
            fSearcher.Add (i)
          End If
        End If
      ' --------------------------------
    End Select
SkipForFSearcher:
  Next
End Function

' проверка ColumnToCheck в текущей книге на листе № WorkSheet
' на соответствие SeekColumnTarget в открываемой книге на листе № WorkSheetTarget
' при совпадении Заполнение данных записать в эту книгу в ColumnToFill
' данные из другой книги из столбца ColumnTargetData
' ======================================================================================================

Function ColToArr(C As Collection) As Variant()
  Dim A() As Variant
  ReDim A(0 To C.Count - 1)
  Dim i As Integer
  For i = 1 To C.Count
    A(i - 1) = C.item(i)
  Next
  CollectionToArray = A
End Function

Sub B1_DUPESNDATA()
  ' как устранить ошибку Code execution has been interrupted
  ' Application.EnableCancelKey = xlDisabled

  Set dPFc = New Dictionary
'      dPFc.Add "1", "OEM_N/A;Name"
'      dPFc.Add "2", "=;NameF"
'      dPFc.Add "3", "=;Vend"
'      dPFc.Add "4", "=;Code"
'      dPFc.Add "5", "=;ID"
'      dPFc.Add "8", "=;TNVED"
'      dPFc.Add "9", "=;Country"
'      dPFc.Add "10", "DATE;Note"
'      dPFc.Add "11", "=;Units"
'      dPFc.Add "12", "=;GTD"
'      dPFc.Add "13", "=;NDS"
'      dPFc.Add "14", "=;Group"
'      dPFc.Add "15", "=;Cargo"
'      dPFc.Add "16", "=;Importer"
'      dPFc.Add "17", "=;PCs"
'      dPFc.Add "18", "=;Engine"
'      dPFc.Add "19", "=;CrossBrand"
'      dPFc.Add "20", "=;Info"
'      dPFc.Add "21", "=;Type"
'      dPFc.Add "22", "=;NameDocs"
'      dPFc.Add "23", "=;WghtN"
'      dPFc.Add "24", "=;WghtB"

'     dPFc.Add "1", "=;Vend"
'     dPFc.Add "4", "=;Code"
'     dPFc.Add "9", "=;Units"
'     dPFc.Add "10", "=;PCs"
'     dPFc.Add "12", "=;OEM"
'     dPFc.Add "25", "=;Name"
'     dPFc.Add "26", "=;Info"
'     dPFc.Add "27", "=;Size"
'     dPFc.Add "29", "=;Info2"
'     dPFc.Add "32", "=;Engine"
'     dPFc.Add "35", "=;ID"
'     dPFc.Add "49", "=;Type1"
'     dPFc.Add "50", "=;Type2"
'     dPFc.Add "51", "=;Type3"
'     'dPFc.Add "2", "=;NameF"
'     dPFc.Add "55", "=;TNVED"
'     dPFc.Add "56", "=;Country"
'     dPFc.Add "58", "DATE;Note"
'     dPFc.Add "59", "=;GTD"
'     dPFc.Add "60", "=;NDS"
'     dPFc.Add "63", "=;CrossBrand"
'     'dPFc.Add "15", "=;Cargo"
'     'dPFc.Add "16", "=;Importer"
'     'dPFc.Add "22", "=;NameDocs"
'     'dPFc.Add "23", "=;WghtN"
'     'dPFc.Add "24", "=;WghtB"
'=================
' исправление слияние дублей
'      'dPFc.Add "1", "<>;ID"
'      dPFc.Add "2", "=;Vend"
'      dPFc.Add "3", "=;Code"
'      dPFc.Add "4", "=;Units"
'      dPFc.Add "5", "=;PCs"
'      dPFc.Add "6", "=;OEM"
'      'dPFc.Add "7", "=;ProdNameSeller"
'      'dPFc.Add "8", "=;DescNameSeller"
'      'dPFc.Add "9", "=;NameENG"
'      dPFc.Add "10", "=;Name"
'      dPFc.Add "11", "=;Info"
'      dPFc.Add "12", "=;Size"
'      dPFc.Add "13", "=;Engine"
'      dPFc.Add "14", "=;Info2"
'      dPFc.Add "15", "=;CrossBrand"
'      'dPFc.Add "16", "=;Type3"
'      'dPFc.Add "17", "=;Type2"
'      'dPFc.Add "18", "=;Type1"
'      'dPFc.Add "19", "=;TNVED"
'      dPFc.Add "20", "=;Country"
'      'dPFc.Add "21", "DATE;Comment"
'      'dPFc.Add "22", "=;GTD"
'      'dPFc.Add "23", "=;NDS"
'      'dPFc.Add "24", "=;Group1C"
'      'dPFc.Add "25", "=;Vid1C"
'      'dPFc.Add "26", "=;Importer"
'      dPFc.Add "28", "=;NameOverall"
      
'=================
' копирование инфы между компл и шт
      'dPFc.Add "1", "<>;ID" ' НЕ ЗАБЫВАЙ КОММЕНТИТЬ ПРИ ПРИМЕРНОМ ПОИСКЕ ИНАЧЕ КОДЫ ДРУГИХ ЗАПЧАСТЕЙ НАЗНАЧАТСЯ И АХТУНГ!!!!!
      'dPFc.Add "2", "=;Vend"
      'dPFc.Add "3", "=;Code"
      'dPFc.Add "4", "=;Units"
      dPFc.Add "5", "=;PCs"
      dPFc.Add "6", "=;OEM"
      dPFc.Add "7", "=;ProdNameSeller"
      dPFc.Add "8", "=;DescNameSeller"
      dPFc.Add "9", "=;NameENG"
      dPFc.Add "10", "=;NameRUS"
      dPFc.Add "11", "=;Info"
      dPFc.Add "12", "=;Size"
      dPFc.Add "13", "=;Engine"
      dPFc.Add "14", "=;Info2"
      dPFc.Add "15", "=;CrossBrand"
      dPFc.Add "16", "=;MaterialENG"
      dPFc.Add "17", "=;MaterialRUS"
      dPFc.Add "18", "=;Type3"
      dPFc.Add "19", "=;Type2"
      dPFc.Add "20", "=;Type1"
      dPFc.Add "21", "=;TNVED"
      dPFc.Add "22", "=;Country"
      'dPFc.Add "23", "DATE;Comment"
      'dPFc.Add "24", "=;GTD"
      'dPFc.Add "25", "=;NDS"
      'dPFc.Add "26", "=;Group1C"
      'dPFc.Add "27", "=;Vid1C"
      'dPFc.Add "28", "=;Importer"
      dPFc.Add "30", "=;NameOverall"
      dPFc.Add "34", "=;SellQuant"
      dPFc.Add "35", "=;GasType"
      dPFc.Add "36", "=;Diameter"
      
'====================
      'dPFc.Add "15", "=;Cargo"

  ' DupesDelete не использовать - нужна сортировка строк от большего к меньшему перед удалением
  Set dP = New Dictionary
      dP.Add "Actions", "DupesCollect" ' "Name;DupesCollect;DupesSum;DupesFix;DupesName;DupesPair;DupesDelete"

  Dim cellmatch As Collection
  Set cellmatch = New Collection
  Dim tmpmatch As Collection
  Set tmpmatch = New Collection
  Dim itermatch As Collection
  Set itermatch = New Collection
  
  Dim paramDelimCase As String

  vLogSessionTimeCurr = CStr(Date) & " " & CStr(Time())
  vLogCurrentFileName = Replace(vLogSessionTimeCurr, ":", ".") & "-" & dP.item("Actions")
  Debug.Print String(65535, vbCr)
  Debug.Print Now & " * Start " & Application.VBE.ActiveCodePane.CodeModule
  Call fLogToFile("this", Now & " * Start " & Application.VBE.ActiveCodePane.CodeModule, vLogCurrentFileName)

  ' Parameters General
      dP.Add "SrcWS", 1
      dP.Add "SrcRowStart", 2
      dP.Add "SrcRowEnd", ThisWorkbook.Worksheets(CInt(dP.item("SrcWS"))).Cells.Find("*", [A1], , , xlByRows, xlPrevious).Row
      dP.Add "SrcColEnd", ThisWorkbook.Worksheets(CInt(dP.item("SrcWS"))).Cells.Find("*", [A1], , , xlByColumns, xlPrevious).Column
      If CLng(dP.item("SrcRowEnd")) < CLng(dP.item("SrcRowStart")) Then
        Debug.Print Now & " * Data range not recognized (turn off filtering). Stop at row " & dP.item("SrcRowEnd")
        Call fLogToFile("this", Now & " * Data range not recognized (turn off filtering). Stop at row " & dP.item("SrcRowEnd"), vLogCurrentFileName)
        Exit Sub
      Else
        flagConfirmRows = MsgBox("Обработать " & CStr(dP.item("SrcRowEnd")) & " строк?", vbYesNo, "Обработка XLSX основная (by BalRoG)") ' 6 ДА, 7 НЕТ 2 ОТМЕНА
        If flagConfirmRows <> 6 Then
          Debug.Print Now & " * Отмена по flagConfirmRows " & dP.item("SrcRowEnd")
          Call fLogToFile("this", Now & " * Отмена по flagConfirmRows " & dP.item("SrcRowEnd"), vLogCurrentFileName)
          Exit Sub
        End If
      End If

      dP.Add "SrcSearchMod", "1" ' SM1=простое соответствие, SM2=равенство какой либо части разделенной пробелом (CASE2), или части строки (CASE3)
      dP.Add "SrcSearchCase", "1" ' 1=регистровый поиск, 2=без-регистровый для SearchMod 1, SearchMod 2 только без "/" и "-", 3=PartIsInArray
      paramCopyHeader = 0
      paramDelimCase = "|"
      dP.Add "SrcFromFile", 1 ' для Name
      dP.Add "SrcColSkipCond", 0 ' для всех, положительное число = пропускать если столбец ЗАПОЛНЕН, отрицательное = пропускать если ПУСТОЙ, 0 = не проверяется
      dP.Add "SrcColSkipCond2", 0 ' для всех, положительное число = пропускать если столбец ЗАПОЛНЕН, отрицательное = пропускать если ПУСТОЙ, 0 = не проверяется

      dP.Add "DBMatchGetIndex", "ffilled" ' first, last, ffilled - какой найденный элемент поиска использовать (ffilled = первый заполненный)

      dP.Add "SrcColChk", "7" ' для Name, DupesName, DupesCollect
      dP.Add "SrcColCmpr", 0 ' для Name
      dP.Add "SrcColFillName", 16 ' КУДА ЗАКИДЫВАТЬ
      dP.Add "SrcColPair1", 5 ' для DupesPair
      dP.Add "SrcColPair2", 1 ' для DupesPair
      dP.Add "SrcColPair3", 8 ' для DupesPair
      dP.Add "SrcColCollect", 15 ' для DupesCollect

      dP.Add "DBWS", 1 ' все по DB для Name и DupesFix
      dP.Add "DBRowStart", 2 ' Name и DupesFix
      dP.Add "DBColSeek", "2" ' Name
      dP.Add "DBColDataName", "11;11" ' Name и DupesFix
      dP.Add "DBColCmpr", 0 ' относит знач из ренджа который выше!!!
      dP.Add "DBPathSrc1", ""

  Set dPF = New Dictionary ' для DupesFix
      vActFixDBRange = "6;40"
      dPF.Add "cID", 14 ' в этом столбце код 1С
      dPF.Add "cVend", 15 ' в этом столбце производитель
      dPF.Add "cCode", 16 ' в этом столбце артикул
      dPF.Add "cUnits", 17 ' в этом столбце ЕД
      dPF.Add "cType", 17 ' в этом столбце Тип

  ' DicDupes
  Set dPairs = CreateObject("Scripting.Dictionary")
  Set cDupes = New Collection
  Set RegExp = CreateObject("vbscript.regexp")
  RegExp.Global = True
  Dim vNameRangeRow As Range
  Dim kvaziRange(1 To 1, 1 To 1) As Variant
  Set cFixPairs = New Collection
  Set cDupesFixLog = New Collection
  RegExp.Pattern = " +" '+ означает, что заменяем и два и три и четыре и более пробелов подряд
  
  ' ------------------------------------------------------------------------------------------
  flagCopyHeader = 0
  With Application
  .EnableEvents = False: .ScreenUpdating = False

  If IsInArray("Name", Split(dP.item("Actions"), ";")) Then
    If CInt(dP.item("SrcFromFile")) > 0 Then
      If dP.item("DBPathSrc1") = "" Then
        With .FileDialog(msoFileDialogFilePicker)
          .Title = "Выберите файл данных"
          .Show
          If .SelectedItems.Count = 0 Then
            MsgBox "Файл не выбран"
            Exit Sub
            dP.item("Actions") = Replace(dP.item("Actions"), "Name", "")
            GoTo markPassName
          End If
          dP.item("DBPathSrc1") = .SelectedItems(1)
        End With
        ' MsgBox dP.Item("DBPathSrc1"), , "Открыт файл"
      End If
      Set xlAPP1 = CreateObject("Excel.Application")
      'xlAPP1.Visible = False 'Visible is False by default, so this isn't necessary
      Set xlBook1 = xlAPP1.Workbooks.Add(dP.item("DBPathSrc1"))
      Set Db = xlBook1.Worksheets(dP.item("DBWS"))
    Else
      MsgBox "Name из этого же листа, SEEK from " & Split(dP.item("DBColSeek"))(0) & " to " & Split(dP.item("DBColSeek"))(UBound(Split(dP.item("DBColSeek")))) & ", NAME " & CLng(Split(dP.item("DBColDataName"), ";")(0)) & " to " & CLng(Split(dP.item("DBColDataName"), ";")(1))
      Set Db = ThisWorkbook.Worksheets(dP.item("DBWS"))
    End If
    With Db
      Dim arrSeekRange As Variant
      Dim vSeekRange() As Range
      ReDim Preserve vSeekRange(UBound(Split(dP.item("SrcColChk"), ";")) + 1)
      For iSR = 0 To UBound(Split(dP.item("SrcColChk"), ";"))
        Set vSeekRange(iSR) = .Range(.Cells(dP.item("DBRowStart"), CLng(Split(dP.item("DBColSeek"), ";")(iSR))), .Cells(.UsedRange.Rows.Count, CLng(Split(dP.item("DBColSeek"), ";")(iSR))))
      Next
      Set vNameRange = .Range(.Cells(dP.item("DBRowStart"), CLng(Split(dP.item("DBColDataName"), ";")(0))), .Cells(.UsedRange.Rows.Count, CLng(Split(dP.item("DBColDataName"), ";")(1))))
    End With
  End If
markPassName:

  With ThisWorkbook.Worksheets(CInt(dP.item("SrcWS")))
  vPrevPair = ""
  flagHadErrors = False
  For i = 1 To dP.item("SrcRowEnd")
    If i < dP.item("SrcRowStart") Then
      If i = 1 And .Cells(i, CInt(dP.item("SrcColFillName"))).Value = "" Then
        .Cells(i, CInt(dP.item("SrcColFillName"))).Value = dP.item("Actions")
      End If
      If dP.item("Actions") = "Name" And i = 1 And paramCopyHeader = 1 Then
        flagCopyHeader = 1
        cellmatch.Add (0)
        GoTo CopyHeader
      Else
        GoTo SkipFor
      End If
    End If

    ' условие пропуска строки при заполненности ячейки
    If dP.item("SrcColSkipCond") > 0 Then
      If Len(CStr(.Cells(i, CInt(dP.item("SrcColSkipCond"))).Value)) > 0 Then GoTo SkipFor
    End If
    If dP.item("SrcColSkipCond") < 0 Then
      If Len(CStr(.Cells(i, Abs(CInt(dP.item("SrcColSkipCond")))).Value)) = 0 Then GoTo SkipFor
    End If
    ' проверка что ячейка объединена
    If dP.item("SrcColSkipCond") <> 0 Then
      If .Cells(i, Abs(CInt(dP.item("SrcColSkipCond")))).MergeCells Then GoTo SkipFor
    End If
    ' условие пропуска строки при заполненности ячейки2
    If dP.item("SrcColSkipCond2") > 0 Then
      If Len(CStr(.Cells(i, CInt(dP.item("SrcColSkipCond2"))).Value)) > 0 Then GoTo SkipFor
    End If
    If dP.item("SrcColSkipCond2") < 0 Then
      If Len(CStr(.Cells(i, Abs(CInt(dP.item("SrcColSkipCond2")))).Value)) = 0 Then GoTo SkipFor
    End If
    ' проверка что ячейка объединена2
    If dP.item("SrcColSkipCond2") <> 0 Then
      If .Cells(i, Abs(CInt(dP.item("SrcColSkipCond2")))).MergeCells Then GoTo SkipFor
    End If

    If IsInArray("DupesCollect", Split(dP.item("Actions"), ";")) Or IsInArray("DupesSum", Split(dP.item("Actions"), ";")) Then
      Call fCollectDupesData(dPairs, i, .Cells(i, CInt(dP.item("SrcColChk"))).Value, .Cells(i, CInt(dP.item("SrcColCollect"))).Value)
      'Call fLogToFile("this", Now & " * i > " & i & " < = > " & .Cells(i, CInt(dP.Item("SrcColChk"))).Value & " <", vLogCurrentFileName)
    End If

    If IsInArray("DupesName", Split(dP.item("Actions"), ";")) Then
      Call fCheckDupes(dPairs, i, .Cells(i, CInt(dP.item("SrcColChk"))).Value)
      Call fLogToFile("this", Now & " * i > " & i & " < = > " & .Cells(i, CInt(dP.item("SrcColChk"))).Value & " <", vLogCurrentFileName)
    End If

    If IsInArray("DupesPair", Split(dP.item("Actions"), ";")) Then
      Call fCheckDupes(dPairs, i, .Cells(i, CInt(dP.item("SrcColPair1"))).Value & "|" & .Cells(i, CInt(dP.item("SrcColPair2"))).Value & "|" & .Cells(i, CInt(dP.item("SrcColPair3"))).Value)
    End If

    If i > 22 Then
      FuckThisShit = 1
      ' GoTo mInterruptAll
    End If

    If IsInArray("Name", Split(dP.item("Actions"), ";")) Then
      Set cellmatch = New Collection

      iSrcColElem = 0
      flagSrcColFound = True
      Do Until (flagSrcColFound = False) Or (iSrcColElem > UBound(Split(dP.item("SrcColChk"), ";")))
        arrSeekRange = vSeekRange(iSrcColElem).Value ' рендж столбец со значениями
        vTmpC = .Cells(i, CInt(Split(dP.item("SrcColChk"), ";")(iSrcColElem))).Value ' значение для поиска по этому столбцу
        If iSrcColElem = 0 Then
          Set cellmatch = fSearcher(vTmpC, arrSeekRange, CInt(Split(dP.item("SrcSearchCase"), ";")(iSrcColElem)), CInt(Split(dP.item("SrcSearchMod"), ";")(iSrcColElem)), paramDelimCase)
        Else
          For Each vIndexMatch In cellmatch
            kvaziRange(1, 1) = arrSeekRange(vIndexMatch, 1) ' {"", arrSeekRange(vIndexMatch, 1); "", ""}]
            Set itermatch = fSearcher(vTmpC, kvaziRange, CInt(Split(dP.item("SrcSearchCase"), ";")(iSrcColElem)), CInt(Split(dP.item("SrcSearchMod"), ";")(iSrcColElem)), paramDelimCase)
            If itermatch.Count > 0 Then
              tmpmatch.Add (vIndexMatch)
            End If
          Next
          If tmpmatch.Count = 0 Then
            flagSrcColFound = False
          Else
            Set cellmatch = tmpmatch
          End If
        End If
        iSrcColElem = iSrcColElem + 1
        Set tmpmatch = New Collection
      Loop
      If cellmatch.Count And flagSrcColFound = True Then
CopyHeader:
        vNameRangeColumnsCount = vNameRange.Columns.Count
        If dP.item("DBMatchGetIndex") = "first" Then
          dP.item("DBMatchIndex") = 1
        ElseIf dP.item("DBMatchGetIndex") = "last" Then
          dP.item("DBMatchIndex") = cellmatch.Count
        ElseIf dP.item("DBMatchGetIndex") = "ffilled" Then
          dP.item("DBMatchIndex") = 1
          vIresult = 1
          Do
            If Not IsEmpty(vNameRange.Cells(cellmatch(vIresult), 1).Value) Then
              dP.item("DBMatchIndex") = vIresult
              vIresult = cellmatch.Count
            End If
            vIresult = vIresult + 1
          Loop Until vIresult > cellmatch.Count
        Else
          dP.item("DBMatchGetIndex") = "first"
          dP.item("DBMatchIndex") = 1
        End If
        Set vNameRangeRow = vNameRange.Rows(CLng(cellmatch.item(dP.item("DBMatchIndex"))))
        .Range(.Cells(i, CLng(dP.item("SrcColFillName"))), .Cells(i, CLng(dP.item("SrcColFillName")) + (vNameRangeColumnsCount - 1))).Value = vNameRangeRow.Value
        If i > 1 Then
          '.Range(.Cells(i, CLng(dP.Item("SrcColFillName"))), .Cells(i, CLng(dP.Item("SrcColFillName")) + (vNameRangeColumnsCount - 1))).Interior.Color = RGB(198, 89, 17)
        End If
        If flagCopyHeader = 1 Then
          flagCopyHeader = 0
          GoTo SkipFor
        End If
        'If Not .Range(.Cells(i, CLng(dP.Item("SrcColFillName"))), .Cells(i, CLng(dP.Item("SrcColFillName")) + (vNameRangeColumnsCount - 1))).Value = vNameRangeRow.Value Then
          ' .Cells(i, CInt(dP.Item("SrcColChk"))).Interior.Color = xlNone
        ' Call fLogToFile("this", Now & " * SrcColChk found diff placed R" & CStr(i) & " C" & CStr(dP.Item("SrcColChk")) & " >" & CStr(vTmpC) & " <", vLogCurrentFileName)
        'End If
        ' .Cells(i, CInt(dP.Item("SrcColFillName"))).Value = vNameRange.Cells(cellMatch.Item(1), 1).Value
        If CInt(dP.item("SrcColCmpr")) > 0 Then
          tmp1 = vNameRange.Cells(cellmatch.item(1), dP.item("DBColCmpr")).Value ' .Cells(i, CInt(dP.Item("SrcColFillName"))).Value
          tmp2 = .Cells(i, CInt(dP.item("SrcColCmpr"))).Value
          If tmp1 <> tmp2 Then
            .Cells(i, CInt(dP.item("SrcColFillName"))).Interior.Color = vbYellow
            Call fLogToFile("this", Now & " * Name diff > " & tmp1 & " < and > " & tmp2 & " <", vLogCurrentFileName)
          End If
        End If
      Else
        .Cells(i, CInt(Split(dP.item("SrcColChk"), ";")(0))).Interior.Color = vbYellow
        Debug.Print "R" & CStr(i) & " * Нету <" & vTmpC & "> среди " & CStr(arrSeekRange(1, 1)) & " | " & CStr(arrSeekRange(2, 1)) & " | " & CStr(arrSeekRange(3, 1))
        Call fLogToFile("this", Now & " * SrcColChk not found R" & CStr(i) & " C" & Split(dP.item("SrcColChk"), ";")(0) & " > " & CStr(vTmpC) & " <", vLogCurrentFileName)
      End If
    End If

    If IsInArray("DupesFix", Split(dP.item("Actions"), ";")) Then
      If Not dP.item("DBColDataName") = vActFixDBRange Then
        If dP.item("Actions") = "DupesFix" Then
          Debug.Print Now & " * не будет выполнен, переданы не все строки"
          Call fLogToFile("this", Now & " * не будет выполнен, переданы не все строки", vLogCurrentFileName)
          dP.item("Actions") = Replace(dP.item("Actions"), "DupesFix", "")
        End If
        GoTo markPassDupesFix
      End If
      If dPF.item("cID") > 0 Then
        If Len(.Cells(i, dPF.item("cID")).Value) < 11 & Len(.Cells(i, dPF.item("cID")).Value) > 0 Then
          Debug.Print Now & " * некорректный формат кода R" & CStr(i)
          Call fLogToFile("this", Now & " * некорректный формат кода R" & CStr(i), vLogCurrentFileName)
          GoTo markPassDupesFix
        End If
      End If
      If Len(.Cells(i, dPF.item("cVend")).Value) = 0 Or Len(.Cells(i, dPF.item("cCode")).Value) = 0 Then
        Debug.Print Now & " * ошибка - нет пары R" & CStr(i)
        Call fLogToFile("this", Now & " * ошибка - нет пары R" & CStr(i), vLogCurrentFileName)
        GoTo markPassDupesFix
      End If

      If cFixPairs.Count < 1 Then
        vPrevPair = .Cells(i, dPF.item("cVend")).Value & " " & .Cells(i, dPF.item("cCode")).Value & " " & .Cells(i, dPF.item("cUnits")).Value & " " & .Cells(i, dPF.item("cType")).Value
        cFixPairs.Add (i)
      Else
        vCellPair = .Cells(i, dPF.item("cVend")).Value & " " & .Cells(i, dPF.item("cCode")).Value & " " & .Cells(i, dPF.item("cUnits")).Value & " " & .Cells(i, dPF.item("cType")).Value
        If vPrevPair = vCellPair Then
          cFixPairs.Add (i)
        End If
        If vPrevPair <> vCellPair Or i >= dP.item("SrcRowEnd") Then
          If cFixPairs.Count > 1 Then
' =======================================================
            For vci = 1 To cFixPairs.Count
              If vci < cFixPairs.Count Then
                vcOffset = 1
              Else
                vcOffset = -1
              End If
              Set dDupesFixLog = New Dictionary
              For K = CLng(Split(dP.item("DBColDataName"), ";")(0)) To CLng(Split(dP.item("DBColDataName"), ";")(1))
                j = (K + CLng(dP.item("SrcColFillName"))) - 1
                If dPFc.Exists(CStr(K)) Then
                  If Not CStr(.Cells(cFixPairs.item(vci), j).Value) = CStr(.Cells(cFixPairs.item(vci + vcOffset), j).Value) _
                  Or Len(.Cells(cFixPairs.item(vci), j).Value) = 0 _
                  Or (Split(dPFc.item(CStr(K)), ";")(0) <> "=" And Split(dPFc.item(CStr(K)), ";")(0) <> "<>" And Split(dPFc.item(CStr(K)), ";")(0) <> "DATE") Then
                    If Len(.Cells(cFixPairs.item(vci), j).Value) = 0 Then
                      vTmpFilledCell = 0
                      For m = 1 To cFixPairs.Count
                        If Not m = vci And Len(.Cells(cFixPairs.item(m), j).Value) Then
                          .Cells(cFixPairs.item(vci), j).Value = .Cells(cFixPairs.item(m), j).Value
                          .Cells(cFixPairs.item(m), j).Interior.ColorIndex = 35
                          .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 43
                          vTmpFilledCell = 1
                          Exit For
                        End If
                      Next
                      If vTmpFilledCell = 1 Then
                        If Not dDupesFixLog.Exists(CStr(cFixPairs.item(m))) Then
                          dDupesFixLog.Add CStr(cFixPairs.item(m)), CStr(K) & ";" & dPFc.item(CStr(K))
                        Else
                          dDupesFixLog.item(CStr(cFixPairs.item(m))) = dDupesFixLog.item(CStr(cFixPairs.item(m))) & " | " & CStr(K) & ";" & dPFc.item(CStr(K))
                        End If
                      End If
                    Else
                      ' -----------------------------------------------
                      If Split(dPFc.item(CStr(K)), ";")(0) <> "=" And Split(dPFc.item(CStr(K)), ";")(0) <> "<>" And Split(dPFc.item(CStr(K)), ";")(0) <> "DATE" Then
                        vTmpFilledCell = 0
                        If InStr(.Cells(cFixPairs.item(vci), j).Value, Split(dPFc.item(CStr(K)), ";")(0)) Then
                          For vPtrnRow = 1 To cFixPairs.Count
                            If Not vPtrnRow = vci And InStr(.Cells(cFixPairs.item(vPtrnRow), j).Value, Split(dPFc.item(CStr(K)), ";")(0)) = 0 Then
                              .Cells(cFixPairs.item(vci), j).Value = .Cells(cFixPairs.item(vPtrnRow), j).Value
                              .Cells(cFixPairs.item(vPtrnRow), j).Interior.ColorIndex = 35
                              .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 43
                              vTmpFilledCell = 1
                              Exit For
                            End If
                          Next
                          If vTmpFilledCell = 1 Then
                            If Not dDupesFixLog.Exists(CStr(cFixPairs.item(vPtrnRow))) Then
                              dDupesFixLog.Add CStr(cFixPairs.item(vPtrnRow)), CStr(K) & ";" & dPFc.item(CStr(K))
                            Else
                              dDupesFixLog.item(CStr(cFixPairs.item(vPtrnRow))) = dDupesFixLog.item(CStr(cFixPairs.item(vPtrnRow))) & " | " & CStr(K) & ";" & dPFc.item(CStr(K))
                            End If
                          Else
                            vTmpFilledCell = 2
                          End If
                        ElseIf .Cells(cFixPairs.item(vci), j).Value <> .Cells(cFixPairs.item(vci + vcOffset), j).Value And InStr(.Cells(cFixPairs.item(vci + vcOffset), j).Value, Split(dPFc.item(CStr(j)), ";")(0)) = 0 Then
                          vTmpFilledCell = 2
                        End If
                      End If
                      ' -----------------------------------------------
                      If (Split(dPFc.item(CStr(K)), ";")(0) = "=" And .Cells(cFixPairs.item(vci), j).Value <> .Cells(cFixPairs.item(vci + vcOffset), j).Value) Or (vTmpFilledCell > 1 And Split(dPFc.item(CStr(K)), ";")(0) <> "<>" And Split(dPFc.item(CStr(K)), ";")(0) <> "DATE") Then
                        If 2 < 1 And (CStr(Split(dPFc.item(CStr(K)), ";")(1)) = "OEM" Or CStr(Split(dPFc.item(CStr(K)), ";")(1))) = "Engine" Then
                          If Len(.Cells(cFixPairs.item(vci), j).Value) < Len(.Cells(cFixPairs.item(vci + vcOffset), j).Value) Then
                            If InStr(.Cells(cFixPairs.item(vci + vcOffset), j).Value, .Cells(cFixPairs.item(vci), j).Value) Then
                              .Cells(cFixPairs.item(vci), j).Value = .Cells(cFixPairs.item(vci + vcOffset), j).Value
                              .Cells(cFixPairs.item(vci + vcOffset), j).Interior.Pattern = xlNone ' убираем заливку
                              .Cells(cFixPairs.item(vci), j).Interior.Pattern = xlNone ' убираем заливку
                            Else
                              .Cells(cFixPairs.item(vci), j).Value = .Cells(cFixPairs.item(vci), j).Value & "|" & .Cells(cFixPairs.item(vci + vcOffset), j).Value
                              .Cells(cFixPairs.item(vci + vcOffset), j).Value = .Cells(cFixPairs.item(vci), j).Value
                              .Cells(cFixPairs.item(vci + vcOffset), j).Interior.Pattern = xlNone ' убираем заливку
                              .Cells(cFixPairs.item(vci), j).Interior.Pattern = xlNone ' убираем заливку
                            End If
                          ElseIf Len(.Cells(cFixPairs.item(vci), j).Value) > Len(.Cells(cFixPairs.item(vci + vcOffset), j).Value) Then
                            If InStr(.Cells(cFixPairs.item(vci), j).Value, .Cells(cFixPairs.item(vci + vcOffset), j).Value) Then
                              .Cells(cFixPairs.item(vci + vcOffset), j).Value = .Cells(cFixPairs.item(vci), j).Value
                              .Cells(cFixPairs.item(vci + vcOffset), j).Interior.ColorIndex = 0 ' убираем заливку
                              .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 0 ' убираем заливку
                            Else
                              .Cells(cFixPairs.item(vci + vcOffset), j).Value = .Cells(cFixPairs.item(vci + vcOffset), j).Value & "|" & .Cells(cFixPairs.item(vci), j).Value
                              .Cells(cFixPairs.item(vci), j).Value = .Cells(cFixPairs.item(vci + vcOffset), j).Value
                              .Cells(cFixPairs.item(vci + vcOffset), j).Interior.ColorIndex = 0 ' убираем заливку
                              .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 0 ' убираем заливку
                            End If
                          Else
                            .Cells(cFixPairs.item(vci + vcOffset), j).Value = .Cells(cFixPairs.item(vci + vcOffset), j).Value & "|" & .Cells(cFixPairs.item(vci), j).Value
                            .Cells(cFixPairs.item(vci), j).Value = .Cells(cFixPairs.item(vci + vcOffset), j).Value
                            .Cells(cFixPairs.item(vci + vcOffset), j).Interior.ColorIndex = 0 ' убираем заливку
                            .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 0 ' убираем заливку
                          End If
                        End If
                        If vTmpFilledCell = 2 And K = 1 Then
                          If Replace(.Cells(cFixPairs.item(vci), j).Value, "-", "") = Replace(.Cells(cFixPairs.item(vci + vcOffset), j).Value, "-", "") Then
                            If Len(.Cells(cFixPairs.item(vci), j).Value) < Len(.Cells(cFixPairs.item(vci + vcOffset), j).Value) Then
                                   .Cells(cFixPairs.item(vci), j).Value = .Cells(cFixPairs.item(vci + vcOffset), j).Value
                                   .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 43
                                   .Cells(cFixPairs.item(vci + vcOffset), j).Interior.ColorIndex = 35
                              If Not dDupesFixLog.Exists(CStr(cFixPairs.item(vci))) Then
                                dDupesFixLog.Add CStr(cFixPairs.item(vci)), CStr(K) & ";" & dPFc.item(CStr(K))
                              Else
                                dDupesFixLog.item(CStr(cFixPairs.item(vci))) = dDupesFixLog.item(CStr(cFixPairs.item(vci))) & " | " & CStr(K) & ";" & dPFc.item(CStr(K))
                              End If
                            ElseIf Len(.Cells(cFixPairs.item(vci + vcOffset), j).Value) < Len(.Cells(cFixPairs.item(vci), j).Value) Then
                                       .Cells(cFixPairs.item(vci + vcOffset), j).Value = .Cells(cFixPairs.item(vci), j).Value
                                       .Cells(cFixPairs.item(vci + vcOffset), j).Interior.ColorIndex = 43
                                       .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 35
                              If Not dDupesFixLog.Exists(CStr(cFixPairs.item(vci + vcOffset))) Then
                                dDupesFixLog.Add CStr(cFixPairs.item(vci + vcOffset)), CStr(K) & ";" & dPFc.item(CStr(K))
                              Else
                                dDupesFixLog.item(CStr(cFixPairs.item(vci + vcOffset))) = dDupesFixLog.item(CStr(cFixPairs.item(vci + vcOffset))) & " | " & CStr(K) & ";" & dPFc.item(CStr(K))
                              End If
                            End If
                            GoTo markSkipDiff
                          End If
                        End If
                        Debug.Print Now & " * данные различаются R" & CStr(cFixPairs.item(vci)) & " <> R" & CStr(cFixPairs.item(vci + vcOffset)) & " <> C" & str(j) & " " & Split(dPFc.item(CStr(K)), ";")(1)
                        Call fLogToFile("this", Now & " * данные различаются R" & CStr(cFixPairs.item(vci)) & " <> R" & CStr(cFixPairs.item(vci + vcOffset)) & " <> C" & str(j) & " " & Split(dPFc.item(CStr(K)), ";")(1), vLogCurrentFileName)
                        .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 45
                        .Cells(cFixPairs.item(vci + vcOffset), j).Interior.ColorIndex = 45
                        If Not dDupesFixLog.Exists(CStr(cFixPairs.item(vci))) Then
                          dDupesFixLog.Add CStr(cFixPairs.item(vci)), CStr(K) & ";" & dPFc.item(CStr(K))
                        Else
                          dDupesFixLog.item(CStr(cFixPairs.item(vci))) = dDupesFixLog.item(CStr(cFixPairs.item(vci))) & " | " & CStr(K) & ";" & dPFc.item(CStr(K))
                        End If
                      End If
                      ' -----------------------------------------------
                      If Split(dPFc.item(CStr(K)), ";")(0) = "DATE" Then
                        vTmpFilledCell = 0
                        vTmpPrevComment = fSearchRegEx(.Cells(cFixPairs.item(vci), j).Value, "[0123][0-9].(0|1)[0-9].(19|20)[0-9]{2}", , 1, False, True, 1)
                        For vDrow = 1 To cFixPairs.Count
                          If Not vDrow = vci Then
                            vTmpLastComment = fSearchRegEx(.Cells(cFixPairs.item(vDrow), j).Value, "[0123][0-9].(0|1)[0-9].(19|20)[0-9]{2}", , 1, False, True, 1)
                            If Len(vTmpPrevComment) Or Len(vTmpLastComment) Then
                              If Len(vTmpPrevComment) And Len(vTmpLastComment) Then
                                If CDate(vTmpPrevComment) < CDate(vTmpLastComment) Then
                                  vTmpPrevComment = vTmpLastComment
                                  vTmpFilledCell = vDrow
                                End If
                              ElseIf Len(vTmpPrevComment) Then
                                vTmpFilledCell = vci
                              Else
                                vTmpPrevComment = vTmpLastComment
                                vTmpFilledCell = vDrow
                              End If
                            End If
                          End If
                        Next
                        If vTmpFilledCell > 0 Then
                          .Cells(cFixPairs.item(vci), j).Value = .Cells(cFixPairs.item(vTmpFilledCell), j).Value
                          .Cells(cFixPairs.item(vTmpFilledCell), j).Interior.ColorIndex = 35
                          .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 43
                          If Not dDupesFixLog.Exists(CStr(cFixPairs.item(vci))) Then
                            dDupesFixLog.Add CStr(cFixPairs.item(vci)), CStr(K) & ";" & dPFc.item(CStr(K))
                          Else
                            dDupesFixLog.item(CStr(cFixPairs.item(vci))) = dDupesFixLog.item(CStr(cFixPairs.item(vci))) & " | " & CStr(cFixPairs.item(vTmpFilledCell)) & " | " & CStr(K) & ";" & dPFc.item(CStr(j))
                          End If
                        End If
                      End If
                      ' -----------------------------------------------
                    End If
                  Else
                    If Split(dPFc.item(CStr(K)), ";")(0) = "<>" Then ' -----------------------------------------------
                      Debug.Print Now & " * данные равны но должны отличаться R" & CStr(cFixPairs.item(vci)) & " = " & CStr(cFixPairs.item(vci + vcOffset))
                      Call fLogToFile("this", Now & " * данные равны но должны отличаться R" & CStr(cFixPairs.item(vci)) & " <> " & CStr(cFixPairs.item(vci + vcOffset)), vLogCurrentFileName)
                      .Cells(cFixPairs.item(vci), j).Interior.ColorIndex = 46
                      .Cells(cFixPairs.item(vci + vcOffset), j).Interior.ColorIndex = 46
                      If Not dDupesFixLog.Exists(CStr(cFixPairs.item(vci))) Then
                        dDupesFixLog.Add CStr(cFixPairs.item(vci)), CStr(K) & ";" & dPFc.item(CStr(K))
                      Else
                        dDupesFixLog.item(CStr(cFixPairs.item(vci))) = dDupesFixLog.item(CStr(cFixPairs.item(vci))) & " | " & dDupesFixLog.item(CStr(m)) & " | " & CStr(K) & ";" & dPFc.item(CStr(K))
                      End If
                    End If
                  End If
                End If
markSkipDiff:
              Next
            Next
            For Each vKeyF In dDupesFixLog
              .Cells(CLng(vKeyF), dP.item("SrcColEnd") + 2).Value = vKeyF & " " & dDupesFixLog.item(vKeyF)
            Next
' =======================================================
          Else
            'Debug.Print Now & " * нет дубля R" & CStr(cFixPairs.Item(1))
            'Call fLogToFile("this", Now & " * нет дубля R" & CStr(cFixPairs.Item(1)), vLogCurrentFileName)
            '.Cells(cFixPairs.Item(1), CLng(Split(dP.Item("DBColDataName"), ";")(1)) + CLng(dP.Item("SrcColFillName"))).Value = .Cells(cFixPairs.Item(1), CLng(Split(dP.Item("DBColDataName"), ";")(1)) + CLng(dP.Item("SrcColFillName"))).Value & " НЕТ ДУБЛЯ!!!!!"
          End If
          vPrevPair = .Cells(i, dPF.item("cVend")).Value & " " & .Cells(i, dPF.item("cCode")).Value & " " & .Cells(i, dPF.item("cUnits")).Value & " " & .Cells(i, dPF.item("cType")).Value
          Set cFixPairs = New Collection
          cFixPairs.Add (i)
        End If
      End If
    End If
markPassDupesFix:

    If 2 < 1 Then
      vTxtCheckResult = ""
      flagTxtCheckLog = False
      vCurCellTargRow = Cell.Row - (dP.item("SrcRowStart") - 1)
      'vValCellTarg = vSeekRange.Cells(vCurCellTargRow, 1)

      For Each vColumnToCheck In Split(dP.item("SrcColChk"), ",")
        vColumnToCheck = CInt(vColumnToCheck)
        vValCellSource = Trim(RegExp.Replace(CStr(Cell.Offset(0, vColumnToCheck - 1)), " "))
        'решнил вместо этого юкейсами Ucase проверять
        'If vValCellSource Like "*[А-Яа-я]*" Then vValCellSource = ConvertRegistr(vValCellSource, 4)
        If InStr(UCase(vValCellTarg), UCase(vValCellSource)) = 0 Or Len(vValCellSource) = 0 Then
          vTxtCheckResult = vTxtCheckResult & Cell.Row & Chr(9) & Cells(1, vColumnToCheck) & Chr(9) & vValCellTarg & Chr(9) & vValCellSource & Chr(13)
          flagTxtCheckLog = True
          flagHadErrors = True
        Else
          'vTxtCheckResult = vTxtCheckResult & Cells(1, vColumnToCheck) & " = YES" & Chr(13)
        End If
      Next
      If flagTxtCheckLog Then
        vLog = fLogToFile("this", vTxtCheckResult)
        Cell.Offset(0, dP.item("SrcColFillName") - 1).Interior.Color = 255
      Else
        Cell.Offset(0, dP.item("SrcColFillName") - 1).Value = vDataRange.Cells(vCurCellTargRow, 1)
      End If
      'cell.Offset(0, dP.Item("SrcColFillName") - 1).Value = vValCellTarg
      'If cell.Row > 200 Then Goto ExitFor ' TEST
    End If
SkipFor:
  Next
ExitFor:

  If IsInArray("DupesName", Split(dP.item("Actions"), ";")) Or IsInArray("DupesPair", Split(dP.item("Actions"), ";")) Then
    Debug.Print Now & " * log list of dupes"
    Call fLogToFile("this", Now & " * log list of dupes", vLogCurrentFileName)

    With dPairs
      For Each Key In .Keys
        If UBound(.item(Key)) > 0 Then
          Debug.Print "Dupe " & Key & " at " & Join(.item(Key), ", ")
          Call fLogToFile("this", Now & " * dupe " & Key & " at " & Join(.item(Key), ", "), vLogCurrentFileName)
          Cells(.item(Key)(0), dP.item("SrcColFillName")).Value = Join(.item(Key), ", ")
          For i = 1 To UBound(.item(Key))
            Cells(.item(Key)(i), dP.item("SrcColFillName")).Value = CStr(.item(Key)(0))
          Next
        End If
      Next

      If IsInArray("DupesDelete", Split(dP.item("Actions"), ";")) Then
        Debug.Print Now & " * write dupes into sheet"
        Call fLogToFile("this", Now & " * write dupes into sheet", vLogCurrentFileName)
        For i = .Count - 1 To 0 Step -1
          If UBound(.Items()(i)) > 0 Then
            For j = UBound(.Items()(i)) To 1 Step -1
              ThisWorkbook.Worksheets(CInt(dP.item("SrcWS"))).Cells(.Items()(i)(j), dP.item("SrcColFillName")).Value = .Items()(i)(0)
              'ThisWorkbook.Worksheets(CInt(dP.Item("SrcWS"))).Rows(.Items()(i)(j)).Delete
              'Call fLogToFile("this", Now & " * delete " & j & " at " & i & " row " & .Items()(i)(j), vLogCurrentFileName)
            Next
          End If
        Next
      End If
    End With
  End If

  If IsInArray("DupesCollect", Split(dP.item("Actions"), ";")) Then
    With dPairs
      For Each Key In .Keys
        vTxt = ""
        For i = 0 To UBound(.item(Key))
          If Len(vTxt) = 0 Then
            vTxt = Split(.item(Key)(i), "|")(1)
          Else
            vTxt = vTxt & ", " & Split(.item(Key)(i), "|")(1)
          End If
        Next
        vTxt = Join(RemoveDupesColl(Split(vTxt, ", ")), ", ")
        For i = 0 To UBound(.item(Key))
          Cells(CLng(Split(.item(Key)(i), "|")(0)), dP.item("SrcColFillName")).Value = vTxt
        Next
      Next
    End With
  End If
  
  If IsInArray("DupesSum", Split(dP.item("Actions"), ";")) Then
    With dPairs
      For Each Key In .Keys
        vSum = ""
        For i = 0 To UBound(.item(Key))
          If Len(vSum) = 0 Then
            vSum = CLng(Split(.item(Key)(i), "|")(1))
          Else
            vSum = vSum + CLng(Split(.item(Key)(i), "|")(1))
          End If
        Next
        Cells(CLng(Split(.item(Key)(0), "|")(0)), dP.item("SrcColFillName")).Value = CStr(vSum)
      Next
    End With
  End If
mInterruptAll:
  If IsInArray("Name", Split(dP.item("Actions"), ";")) And CInt(dP.item("SrcFromFile")) = 1 Then
    xlBook1.Close SaveChanges:=False
    xlAPP1.Quit
  End If

  End With
  
  .EnableEvents = True: .ScreenUpdating = True
  DoEvents
  End With
  Debug.Print Now & " * Finish " & Application.VBE.ActiveCodePane.CodeModule
  Call fLogToFile("this", Now & " * Finish " & Application.VBE.ActiveCodePane.CodeModule, vLogCurrentFileName)

  ' как устранить ошибку Code execution has been interrupted
  ' Application.EnableCancelKey = xlEnabled
End Sub

Sub B2_1CRPT_DUPHANDL()
  vSubName = "1CRPT_DUPHANDL"
  vLogSessionTimeCurr = CStr(Date) & " " & CStr(Time())
  Debug.Print String(65535, vbCr)
  Debug.Print Now & " * Start " & vSubName & " (" & Application.VBE.ActiveCodePane.CodeModule & ")"
  Call fLogToFile("this", Now & " * Start " & vSubName & " (" & Application.VBE.ActiveCodePane.CodeModule & ")", vSubName)

  ' Parameters General
  Set dP = New Dictionary
      dP.Add "SrcWS", 1
      Set oTWS = ThisWorkbook.Worksheets(CInt(dP.item("SrcWS")))
      dP.Add "SrcRowF", 2
      dP.Add "SrcRowEnd", oTWS.Cells.Find("*", [A1], , , xlByRows, xlPrevious).Row
      dP.Add "SrcColChk", 1
      dP.Add "SrcPtrnSeek", "(?:(\s|^)Ячейка\[R)(\d*)(?:C*)"
      dP.Add "SrcColFillFrom", 2

      dP.Add "DBsToOpen", 2
      dP.Add "DBAWS", 1
      dP.Add "DBARowF", 2
      dP.Add "DBAColToSeek", Array(4, 2)

      dP.Add "DB1PathSrc", "H:\Ramax\Документы\5-Поступление товара\013 20170414 ТТН 3929145 СФ 24 поступление (Минск) (обработано Димой)\013 20170414 ТТН 3929145 СФ 24 (обработано Димой).xlsm"
      dP.Add "DB2PathSrc", "H:\Ramax\Документы\5-Поступление товара\014 20170414 ТТН 3929146 СФ 25 поступление (Минск) (обработано Димой)\014 20170414 ТТН 3929146 СФ 25 (обработано Димой).xlsm"

      dP.Add "Actions", "CopyRowsDB" ' "CopyRowsDB;MarkRows;HideRows"

  Set dRpt = New Dictionary
  Set dDBs = New Dictionary
  Set dPairs = New Dictionary

  With Application
  .EnableEvents = False: .ScreenUpdating = False


  For i = 1 To dP.item("DBsToOpen")
    If dP.item("DB" & CStr(i) & "PathSrc") = "" Then
      With .FileDialog(msoFileDialogFilePicker)
        .Title = "Выберите БД №" & CStr(i)
        .Show
        If .SelectedItems.Count = 0 Then
          MsgBox "Файл не выбран"
          Exit Sub
        End If
        dP.item("DB" & CStr(i) & "PathSrc") = .SelectedItems(1)
      End With
      MsgBox dP.item("DB" & CStr(i) & "PathSrc"), , "Открыт файл"
    End If
    Set oApp = CreateObject("Excel.Application")
    If IsInArray("MarkRows", Split(dP.item("Actions"), ";")) Or IsInArray("HideRows", Split(dP.item("Actions"), ";")) Then
      oApp.Visible = True 'Visible is False by default
    End If
    Set oBook = oApp.Workbooks.Add(dP.item("DB" & CStr(i) & "PathSrc"))
    Set oDB = oBook.Worksheets(dP.item("DBAWS"))
    dDBs.Add "oDB" & CStr(i) & "App", oApp
    dDBs.Add "oDB" & CStr(i) & "Book", oBook
    dDBs.Add "oDB" & CStr(i), oDB
    dDBs.Add "oDB" & CStr(i) & "RowL", dDBs.item("oDB" & CStr(i)).Cells.Find("*", dDBs.item("oDB" & CStr(i)).Cells(1, 1), , , xlByRows, xlPrevious).Row
    dDBs.Add "oDB" & CStr(i) & "ColL", dDBs.item("oDB" & CStr(i)).Cells.Find("*", dDBs.item("oDB" & CStr(i)).Cells(1, 1), , , xlByRows, xlPrevious).Column

'   With Db
'     Set vDbSeekRange = .Range(.Cells(dP.Item("DBRowStart"), dP.Item("DBColSeek")), .Cells(.UsedRange.Rows.Count, dP.Item("DBColSeek")))
'   End With
  Next

  j = 1
  With oTWS
    For i = 1 To dP.item("SrcRowEnd")
      vTmp = fSearchRegEx(.Cells(i, dP.item("SrcColChk")).Value, dP.item("SrcPtrnSeek"), , 1, False, True, 2)
      If Len(vTmp) Then
        Debug.Print CStr(j) & " = " & vTmp
        dRpt.Add CStr(j), Array(vTmp, i)
        j = j + 1
      End If
    Next
  End With

  For i = 1 To dP.item("DBsToOpen")
    Set dPairs.item("dPair" & CStr(i)) = New Dictionary
    With dDBs.item("oDB" & CStr(i))
      For j = dP.item("DBARowF") To dDBs.item("oDB" & CStr(i) & "RowL")
        vDC1 = .Cells(j, dP.item("DBAColToSeek")(0)).Value
        vDC2 = .Cells(j, dP.item("DBAColToSeek")(1)).Value
        Call fCheckDupes(dPairs.item("dPair" & CStr(i)), j, vDC1 & "|" & vDC2)
      Next
    End With
  Next

  If dRpt.Count = 0 Then GoTo flagSkipFor
  vCurRow = dP.item("SrcRowF")
  vRowInsOffset = 0
  For Each vKey In dRpt
    For i = 1 To dP.item("DBsToOpen")
      With dDBs.item("oDB" & CStr(i))
        N = 1
        If IsInArray("CopyRowsDB", Split(dP.item("Actions"), ";")) Then
          If i = 1 Then
            vDC1 = .Cells(dRpt.item(vKey)(0), dP.item("DBAColToSeek")(0)).Value
            vDC2 = .Cells(dRpt.item(vKey)(0), dP.item("DBAColToSeek")(1)).Value
            .Rows(CStr(dRpt.item(vKey)(0))).EntireRow.Interior.Color = xlNone
            oTWS.Range(oTWS.Cells(dRpt.item(vKey)(1) + vRowInsOffset, 2), oTWS.Cells(dRpt.item(vKey)(1) + vRowInsOffset, dDBs.item("oDB" & CStr(i) & "ColL") + 1)).Value = .Range(.Cells(dRpt.item(vKey)(0), 1), .Cells(dRpt.item(vKey)(0), dDBs.item("oDB" & CStr(i) & "ColL"))).Value
          Else
            If dPairs.item("dPair" & CStr(i)).Exists(vDC1 & "|" & vDC2) Then
              For Each PairRow In dPairs.item("dPair" & CStr(i)).item(vDC1 & "|" & vDC2)
                .Rows(PairRow).EntireRow.Interior.Color = xlNone
                oTWS.Rows(dRpt.item(vKey)(1) + N + vRowInsOffset).EntireRow.Insert
                oTWS.Range(oTWS.Cells(dRpt.item(vKey)(1) + N + vRowInsOffset, 2), oTWS.Cells(dRpt.item(vKey)(1) + N + vRowInsOffset, dDBs.item("oDB" & CStr(i) & "ColL") + 1)).Value = .Range(.Cells(PairRow, 1), .Cells(PairRow, dDBs.item("oDB" & CStr(i) & "ColL"))).Value
                For K = 2 To dDBs.item("oDB" & CStr(i) & "ColL") + 1
                  If oTWS.Cells(dRpt.item(vKey)(1) + N + vRowInsOffset, K) <> oTWS.Cells(dRpt.item(vKey)(1) + vRowInsOffset, K) Then oTWS.Cells(dRpt.item(vKey)(1) + N + vRowInsOffset, K).Interior.Color = vbYellow
                Next
                N = N + 1
                vRowInsOffset = vRowInsOffset + 1
              Next
            End If
          End If

          ' .Range(.Cells(dRpt.Item(vKey)(0), ), .Cells(dRpt.Item(vKey)(0), ))
        End If

        If IsInArray("MarkRows", Split(dP.item("Actions"), ";")) Then
          If i = 1 Then
            vDC1 = .Cells(dRpt.item(vKey)(0), dP.item("DBAColToSeek")(0)).Value
            vDC2 = .Cells(dRpt.item(vKey)(0), dP.item("DBAColToSeek")(1)).Value
            .Rows(CStr(dRpt.item(vKey)(0))).EntireRow.Interior.Color = vbYellow
          Else
            If dPairs.item("dPair" & CStr(i)).Exists(vDC1 & "|" & vDC2) Then
              For Each PairRow In dPairs.item("dPair" & CStr(i)).item(vDC1 & "|" & vDC2)
                .Rows(CStr(PairRow)).EntireRow.Interior.Color = vbYellow ' xlNone
              Next
            End If
          End If
        End If

        If IsInArray("HideRows", Split(dP.item("Actions"), ";")) Then
          If CLng(dRpt.item(vKey)(0)) > vCurRow Then
            vHidRowF = vCurRow
            vHidRowL = dRpt.item(vKey)(0) - 1
            .Rows(CStr(vHidRowF) & ":" & CStr(vHidRowL)).EntireRow.Hidden = True
            If CLng(vKey) = dRpt.Count And (CLng(dRpt.item(vKey)(0)) < CLng(dP.item("SrcRowEnd"))) Then
              .Rows(CStr(dRpt.item(vKey)(0) + 1) & ":" & dP.item("SrcRowEnd")).EntireRow.Hidden = True
            End If
          End If
          vCurRow = dRpt.item(vKey)(0) + 1
        End If

      End With
    Next
  Next

flagSkipFor:

  If IsInArray("CopyRowsDB", Split(dP.item("Actions"), ";")) And Not IsInArray("MarkRows", Split(dP.item("Actions"), ";")) And Not IsInArray("HideRows", Split(dP.item("Actions"), ";")) Then
    For i = 1 To dP.item("DBsToOpen")
      dDBs.item("oDB" & CStr(i) & "Book").Close SaveChanges:=False
      dDBs.item("oDB" & CStr(i) & "App").Quit
    Next
  End If

  .EnableEvents = True: .ScreenUpdating = True
  DoEvents
  End With
  Debug.Print Now & " * Done " & vSubName & " (" & Application.VBE.ActiveCodePane.CodeModule & ")"
  Call fLogToFile("this", Now & " * Done " & vSubName & " (" & Application.VBE.ActiveCodePane.CodeModule & ")", vSubName)
End Sub

'Error_Val  Error Value
'#NULL!     Error 2000
'#DIV/0!    Error 2007
'#VALUE!    Error 2015
'#REF!      Error 2023
'#NAME?     Error 2029
'#NUM!      Error 2036
'#N/A       Error 2042
'=======================================================================================================
' процедура добавления номенклатурного наименования из файла по номеру строки
'=======================================================================================================
Sub B9_GetNameByNumberFile()
  Application.ScreenUpdating = False
  Debug.Print String(65535, vbCr)
  Debug.Print "Start " & Now

  Dim dParam
  Set dParam = CreateObject("Scripting.Dictionary")
  dParam.Add "startRow", 2
  dParam.Add "BooksOffset", 23
  Dim vSeekRange As Range

  Set xlAPP1 = CreateObject("Excel.Application")
  'xlAPP1.Visible = False 'Visible is False by default, so this isn't necessary
  Set xlBook1 = xlAPP1.Workbooks.Add(ThisWorkbook.Path & "\2605.xls")
  Set w1 = xlBook1.Worksheets(1)
  Set xlAPP2 = CreateObject("Excel.Application")
  Set xlBook2 = xlAPP2.Workbooks.Add(ThisWorkbook.Path & "\2933.xls")
  Set w2 = xlBook2.Worksheets(1)
  Dim wC As Worksheet

  If xlBook1 Is Nothing Or xlBook2 Is Nothing Then
    MsgBox "Error opening file" & Chr(13) & "xlBook1=" & CStr(xlBook1) & " xlBook2=" & CStr(xlBook2)
    GoTo ExitFor
  End If

  Set vIntersect = Intersect(ThisWorkbook.Worksheets(1).UsedRange, Columns(1)).Cells

  For Each Cell In vIntersect
    If Cell.Row < dParam.item("startRow") Or (Cell.Value <> 2605 And Cell.Value <> 2933) Then
      If Cell.Row >= dParam.item("startRow") Then Debug.Print Cell.Row & " - error, no file specified"
      GoTo SkipFor
    End If
    cellmatch = ""

    If Cell = 2605 Then
      Set wC = w1
    ElseIf Cell = 2933 Then
      Set wC = w2
    End If

    If wC Is Nothing Then
      GoTo SkipFor:
    End If
    vSeekStartRow = dParam.item("startRow") + dParam.item("BooksOffset")
    vSeekLastRow = wC.UsedRange.Rows.Count
    With wC
      Set vSeekRange = .Range(.Cells(vSeekStartRow, 2), .Cells(vSeekLastRow, 2))
    End With
    vNomPos = Cell.Offset(0, 1)
    cellmatch = Application.Match(vNomPos, vSeekRange, 0)
    If IsError(cellmatch) Then
      MsgBox Cell.Row & " error " & " not found file " & Cell & " pos " & vNomPos
      Exit For
    Else
      cellmatch = cellmatch + (vSeekStartRow - 1)
      'MsgBox "Val " & cell.Offset(0, 1) & " at pos " & cellMatch
      Cell.Offset(0, 2).Value = wC.Cells(cellmatch, 3).Value
      Cell.Offset(0, 15).Value = wC.Cells(cellmatch, 42).Value
      Cell.Offset(0, 16).Value = wC.Cells(cellmatch, 43).Value
    End If

    'If cell.Row > 3 Then
    '  GoTo ExitFor:
    'End If

SkipFor:
  Next
ExitFor:

  xlBook1.Close SaveChanges:=False
  xlAPP1.Quit
  Set xlAPP1 = Nothing
  xlBook2.Close SaveChanges:=False
  xlAPP2.Quit
  Set xlAPP2 = Nothing

  Application.ScreenUpdating = True
  Debug.Print "Finish " & Now
End Sub

' ======================================================================================================
' разобраться
' ======================================================================================================
Sub B9_FIX_NOMEKL_SUN()
  vLogSessionTimeCurr = CStr(Date) & " " & CStr(Time())
  Debug.Print String(65535, vbCr)

  Dim Dic
  Set Dic = CreateObject("Scripting.Dictionary")

  Set RegExp = CreateObject("vbscript.regexp")
  RegExp.Global = True
  RegExp.Pattern = " +" '+ означает, что заменяем и два и три и четыре и более пробелов подряд

  vSeekStartRow = 2
  vColCode = 4

  With Application
  .EnableEvents = False: .ScreenUpdating = False

  Set wC = ThisWorkbook.Sheets(1)
  vSeekLastRow = wC.UsedRange.Rows.Count
  With wC
    Set vSeekRange = .Range(.Cells(vSeekStartRow, vColCode), .Cells(vSeekLastRow, vColCode))
  End With

  Debug.Print "Start " & Now

  Set vIntersect = Intersect(ThisWorkbook.Worksheets(1).UsedRange, Columns(1)).Cells
  For Each Cell In vIntersect
    vCurRow = Cell.Row
    If Cell.Row < vSeekStartRow Or Dic.Exists(CStr(Cell.Row)) Then
      GoTo SkipFor
    End If

    ' сделать чтобы полное наименование состояло только из производителя и артикула
    If Cells(Cell.Row, 2).Value Like "*[А-Яа-я]*" Then
      Debug.Print "OLD=" & Cells(Cell.Row, 2).Value & " NEW=" & Cells(Cell.Row, 3).Value & " " & Cells(Cell.Row, vColCode).Value
      Cells(Cell.Row, 2).Value = Cells(Cell.Row, 3).Value & " " & Cells(Cell.Row, vColCode).Value
    End If

    arrSrchR = GetSearchArray(Cells(Cell.Row, vColCode).Value, "1", vSeekRange, True)
    If UBound(arrSrchR) > 0 Then
      vOEMCode = ""
      vEmpty = ""
      For Each R In arrSrchR
        Dic.Add R, ""
        If Not Cells(R, 1).Value Like "*OEM_N/A*" Then
          If Right(Cells(R, 1).Value, 1) = ")" And fCBrc(Cells(R, 1).Value, 1, 1, -1) > 0 Then ' если есть артикул OEM (в скобках)
            vTmp = Split(fCBrc(Cells(R, 1).Value, 1, 1, -1), " ")
            If vOEMCode = vbNullString Then
              vOEMCode = mGetWord(Cells(R, 1).Value, vTmp(0), vTmp(1), 1) & "=" & R ' вытащить артикул OEM в скобках
            Else
              vOEMCode = vOEMCode & ";" & mGetWord(Cells(R, 1).Value, vTmp(0), vTmp(1), 1) & "=" & R  ' вытащить артикул OEM в скобках
            End If
            ' Debug.Print "C = " & Cell.Row & " R = "& r & " NOM = " & Cells(Cell.Row, 1).Value & " OEM = " & vOEMCode
          End If
        Else
          If vEmpty = vbNullString Then
            vEmpty = R
          Else
            vEmpty = vEmpty & ";" & R
          End If
        End If
      Next
      If Not vOEMCode = vbNullString Then
        If UBound(Split(vOEMCode, ";")) > 0 Then
          vOldR = ""
          For Each R In Split(vOEMCode, ";")
            If Split(R, "=")(0) <> vOldR And vOldR <> "" Then
              Debug.Print "C = " & Split(R, "=")(1) & " R = " & Split(R, "=")(0) & " NOM = " & Cells(Split(R, "=")(1), 1).Value
              GoTo SkipFor
            End If
            vOldR = Split(R, "=")(0)
          Next
        Else
          vOldR = Split(vOEMCode, "=")(0)
        End If
        For Each R In Split(vEmpty, ";")
          Debug.Print "REPL AT " & R & " = " & Cells(R, 1).Value & " | TO = " & Replace(Cells(R, 1).Value, "OEM_N/A", vOldR)
          Cells(R, 1).Value = Replace(Cells(R, 1).Value, "OEM_N/A", vOldR)
        Next
      End If
    End If
    'vNom = Cells(Cell.Row, vColCode).Value ' артикул
    'cellMatch = Application.Match(vNom, vSeekRange, 0)
    'If IsError(cellMatch) Then
    ' Debug.Print Cell.Row & " error " & " not found file " & vNom & " pos " & cellMatch
    'Else
    ' cellMatch = cellMatch + (vSeekStartRow - 1)
    ' 'MsgBox "Val " & cell.Offset(0, 1) & " at pos " & cellMatch
    ' Debug.Print "AT=" & cellMatch & " FOUND=" & Cells(cellMatch, vColCode).Value
    'End If

SkipFor:
  Next
ExitFor:
  Debug.Print "Finish " & Now

  .EnableEvents = True: .ScreenUpdating = True
  DoEvents
  End With
  Set Dic = Nothing
End Sub

' ======================================================================================================
' Получить OEM из наименования
' ======================================================================================================
Sub B9_GetOEM()
  Set RegExp = CreateObject("vbscript.regexp")
  RegExp.Global = True
  RegExp.Pattern = " +" '+ означает, что заменяем и два и три и четыре и более пробелов подряд
  ' ===
  pOverwrite = False
  vColSeek = 1
  vColInput = 30

  With Application
    .EnableEvents = False: .ScreenUpdating = False

    Set vIntersect = Intersect(ThisWorkbook.Worksheets(1).UsedRange, Columns(vColSeek)).Cells
    For Each Cell In vIntersect
      If Cell.Row < 2 Then
        GoTo SkipFor
      End If

      If Len(Cells(Cell.Row, vColInput).Value) = 0 Or (pOverwrite And Len(Cells(Cell.Row, vColInput).Value) > 0) Then
        If (Right(Cell.Value, 1) = ")" Or (InStr(Cell.Value, ")") And (InStr(Cell.Value, "SUN Ремень") Or InStr(Cell.Value, "DONGIL Ремень")))) And fCBrc(Cell.Value, 1, 1, -1) > 0 Then ' если есть артикул OEM (в скобках)
          vTmp = Split(fCBrc(Cell.Value, 1, 1, -1), " ")
          vOEMCode = mGetWord(Cell.Value, vTmp(0), vTmp(1), 1) ' вытащить артикул OEM в скобках
        Else
          vOEMCode = "OEM_N/A"
        End If
        Cells(Cell.Row, vColInput).Value = vOEMCode
      End If
SkipFor:
    Next

  End With
End Sub

Sub colors56()
'57 colors, 0 to 56
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual   'pre XL97 xlManual
Dim i As Long
Dim str0 As String, str As String
For i = 0 To 56
  Cells(i + 1, 1).Interior.ColorIndex = i
  Cells(i + 1, 1).Value = "[Color " & i & "]"
  Cells(i + 1, 2).Font.ColorIndex = i
  Cells(i + 1, 2).Value = "[Color " & i & "]"
  str0 = Right("000000" & Hex(Cells(i + 1, 1).Interior.Color), 6)
  'Excel shows nibbles in reverse order so make it as RGB
  str = Right(str0, 2) & Mid(str0, 3, 2) & Left(str0, 2)
  'generating 2 columns in the HTML table
  Cells(i + 1, 3) = "#" & str & "#" & str & ""
  Cells(i + 1, 4).Formula = "=Hex2dec(""" & Right(str0, 2) & """)"
  Cells(i + 1, 5).Formula = "=Hex2dec(""" & Mid(str0, 3, 2) & """)"
  Cells(i + 1, 6).Formula = "=Hex2dec(""" & Left(str0, 2) & """)"
  Cells(i + 1, 7) = "[Color " & i & ")"
Next i
done:
  Application.Calculation = xlCalculationAutomatic  'pre XL97 xlAutomatic
  Application.ScreenUpdating = True
End Sub

' Получить данные из краснодарского формата - в скобках гтд и страна
Sub fGet2LastWrdsFromBrc()
  Debug.Print String(65535, vbCr)

  Set dP = New Dictionary
      dP.Add "Src", 9
      dP.Add "Gtd", 14
      dP.Add "Cntry", 15

  With ActiveSheet
    For i = 405 To 452
      vTmp = Split(fCBrc(.Cells(i, dP.item("Src")).Value, 1, 1, -1), " ")
      vTxt = mGetWord(.Cells(i, dP.item("Src")).Value, vTmp(0), vTmp(1), 1)
      arrGTDCountry = Split(vTxt, ", ", 2)
      If Len(Join(arrGTDCountry, " ")) > 0 Then
        .Cells(i, dP.item("Gtd")).Value = arrGTDCountry(0)
        If UBound(arrGTDCountry) > 0 Then
          .Cells(i, dP.item("Cntry")).Value = arrGTDCountry(1)
        End If
        .Cells(i, dP.item("Src")).Value = fRemWord(.Cells(i, dP.item("Src")).Value, vTmp(0), vTmp(1), 1, 1)
      End If
    Next
  End With

End Sub

' Получить позиции Nое слово из строки
Function fGetNwordFromStr(ByVal S As String, ByVal N As Integer, Optional ByVal R As Integer)
  If R <> 1 And R <> -1 Then
    R = 1
  End If
  vTmp = fFindChr(S, " ", Len(S), 2, "", R) + 1 & " " & fFindChr(S, " ", Len(S), 1, "", R) - 1
  fGetNwordFromStr = mGetWord(S, Split(vTmp, " ")(0), Split(vTmp, " ")(1))
End Function

Sub GetNwordFromStr()
  Debug.Print String(65535, vbCr)

  Set dP = New Dictionary
      dP.Add "Src", 9
      dP.Add "Trg", 19

  With ActiveSheet
    For i = 414 To 452
      vTmp = fGetNwordFromStr(.Cells(i, dP.item("Src")).Value, 1, -1)
      .Cells(i, dP.item("Src")).Value = WordRemove(.Cells(i, dP.item("Src")).Value, vTmp)
      .Cells(i, dP.item("Trg")).Value = vTmp
      'vTxt = mGetWord(.Cells(i, dP.Item("Src")).Value, vTmp(0), vTmp(1), 1)
      'arrGTDCountry = Split(vTxt, ", ", 2)
      'If Len(Join(arrGTDCountry, " ")) > 0 Then
      '  .Cells(i, dP.Item("Gtd")).Value = arrGTDCountry(0)
       ' If UBound(arrGTDCountry) > 0 Then
       '   .Cells(i, dP.Item("Cntry")).Value = arrGTDCountry(1)
       ' End If
       ' .Cells(i, dP.Item("Src")).Value = fRemWord(.Cells(i, dP.Item("Src")).Value, vTmp(0), vTmp(1), 1, 1)
      'End If
    Next
  End With
End Sub

' Получить текущие D дату T время и F для формирования имени файла
Function fGetDT(pFlags)
  fGetDT = ""
  If InStr(pFlags, "D") Then fGetDT = CStr(Date)
  If InStr(pFlags, "T") Then
    If Len(fGetDT) Then fGetDT = fGetDT & " "
    If InStr(pFlags, "F") Then
      fGetDT = fGetDT & Replace(CStr(Time()), ":", ".")
    Else
      fGetDT = fGetDT & CStr(Time())
    End If
  End If
End Function

Function fDateDiff(ByVal vD1 As Date, Optional ByVal vD2 As Date) As String
  If IsMissing(vD2) Then vD2 = Now
  diff = Abs(vD1 - vD2)
  txt = "0s"
  If Abs(vD2 - vD1) >= 1 Then txt = txt & Abs(Int(vD2 - vD1)) & "d "
  If Abs(Hour(vD2 - vD1)) >= 1 Then txt = txt & Hour(Abs(vD2 - vD1)) & "h "
  If Abs(Minute(vD2 - vD1)) >= 1 Then txt = txt & Abs(Minute(vD2 - vD1)) & "m "
  If Abs(Second(vD2 - vD1)) >= 1 Then txt = txt & Abs(Second(vD2 - vD1)) & "s"
  fDateDiff = Trim(txt)
End Function

'===================================================================================================
' Главная процедура для итерации по текущему листу книги, к которому привязан модуль
Sub A0_MainIterationProc()
  With Application
  .EnableEvents = False: .ScreenUpdating = False

  Debug.Print String(65535, vbCr)
  vLogSessionTimeCurr = CStr(Date) & " " & CStr(Time())
  Debug.Print Now & " * Start MainIterationProc"
  flagExitSub = 0
  Dim D1 As Date
  Dim Ddiff As String
  Dim vStartSel As Range
  Set vStartSel = Selection
  ' Parameters General
  Set dP = New Dictionary
  '==== EDIT THESE ROWS
  dP.Add "Actions", "pBubbleSort" ' "pGenNomNoSpecV2" "pCountTNVEDcustom" "pFixSpecChr" "pGenNomNoSpec" "pMergeCelOp" "pConvertRegistr"
  dP.Add "SrcWS", 3
  dP.Add "SrcRowStart", 23
  '==== END EDIT THESE ROWS
  vLogCurrentFileName = fGetDT("DTF") & "-" & dP.item("Actions")
  With ThisWorkbook.Worksheets(CInt(dP.item("SrcWS")))
'====
  dP.Add "SrcRowEnd", .Cells.Find("*", [A1], , , xlByRows, xlPrevious).Row
  dP.Add "SrcColEnd", .Cells.Find("*", [A1], , , xlByColumns, xlPrevious).Column ' ОЧЕНЬ ВАЖНО ИСКАТЬ ПО xlByColumns ИНАЧЕ ПОТЕРЯ ДАННЫХ И КОТОСТРОФА!

  If CLng(dP.item("SrcRowEnd")) < CLng(dP.item("SrcRowStart")) Then
    Debug.Print Now & " MainIterationProc * Range err from R" & dP.item("SrcRowStart") & " to R" & dP.item("SrcRowEnd")
    flagExitSub = 1
  Else
    .Cells(dP.item("SrcRowEnd"), dP.item("SrcColEnd")).Select
    flagConfirmRows = MsgBox("Модули: " & Replace(dP.item("Actions"), ";", ", ") & Chr(13) & "Обработать строки с " & CStr(dP.item("SrcRowStart")) & " по " & CStr(dP.item("SrcRowEnd")) & "?", vbYesNo, "Главная процедура итерации (by BalRoG)") ' 6 ДА, 7 НЕТ 2 ОТМЕНА
    If flagConfirmRows <> 6 Then
      Debug.Print Now & " MainIterationProc * Cancel flagConfirmRows " & dP.item("SrcRowEnd")
      flagExitSub = 1
    End If
  End If
  vStartSel.Select
  If flagExitSub Then Exit Sub
'====
  End With

  'Call fLogToFile("this", Now & " * Start " & Application.VBE.ActiveCodePane.CodeModule, vLogCurrentFileName)

  For Each vProc In Split(dP.item("Actions"), ";")
    D1 = Now
    Debug.Print D1 & " * Launch proc " & vProc
    Call fLogToFile("this", " * Launch proc " & vProc, vLogCurrentFileName)
    Application.Run vProc, ThisWorkbook.Worksheets(CInt(dP.item("SrcWS"))), dP.item("SrcRowStart"), dP.item("SrcRowEnd"), dP.item("SrcColEnd"), vLogCurrentFileName
    Ddiff = fDateDiff(Now, D1)
    Debug.Print Now & " * Finish proc " & vProc & " in " & Ddiff
    Call fLogToFile("this", " * Finish proc " & vProc & " in " & Ddiff, vLogCurrentFileName)
  Next

  .EnableEvents = True: .ScreenUpdating = True
  End With
End Sub

'===================================================================================================
' Подсветка различий между строками с одинаковым реквизитом (напр. кодом, или наименованием
Function fAddHeaderCounter(ByVal txt As String, ByVal N As Long)
  If InStr(1, txt, " | ", vbTextCompare) Then
    fAddHeaderCounter = Split(txt, " | ")(0) & " | " & CStr(CLng(Split(txt, " | ")(1)) + N)
  Else
    fAddHeaderCounter = txt & " | " & CStr(N)
  End If
End Function

Sub pColorDiff(SrcWS, vFRow, vLRow, vLCol, vLogCurrentFileName)
  vColStart = 2
  vColStop = 38
  arrSkp = Array(4)
  vColSimilarity = 3
  vLastPair = ""

  Dim arrPair As Variant
      arrPair = Array()

  With SrcWS
    For k = vColStart To vColStop ' очищаем из заголовка счетчик и цвет
      .Cells(vFRow - 1, k).Interior.ColorIndex = xlNone
      .Cells(vFRow - 1, k).Value = Split(.Cells(vFRow - 1, k).Value, " | ")(0)
    Next
  
    For i = vFRow To vLRow
			
			'==========================================
			Select Case .Cells(i, vColSimilarity).Value = vLastPair
			'==========================================
      Case TRUE ' если образовалась пара (2ой элемент и дальше)
	      Call fPopArr(arrPair, i)
	      GoTo NextForI
			'==========================================
			Case FALSE ' если пошел новый массив
	      If fArrCnt(arrPair) > 1 Then ' если до этого была пара - сверить и подсветить различия
	        For j = 0 To UBound(arrPair) - 1
	          For k = vColStart To vColStop
	            .Cells(arrPair(j), k).Interior.ColorIndex = xlNone
	            .Cells(arrPair(j + 1), k).Interior.ColorIndex = xlNone
	
	            If IsInArray(j, arrSkp) Then GoTo NextForK
	            
	            If .Cells(arrPair(j), k).Value <> .Cells(arrPair(j + 1), k).Value Then
	              vTxt = vTxt & .Cells(vFRow - 1, k).Value & ", "
	              .Cells(arrPair(j), k).Interior.ColorIndex = 45
	              .Cells(arrPair(j + 1), k).Interior.ColorIndex = 45
	              .Cells(vFRow - 1, k).Interior.ColorIndex = 46
	              .Cells(vFRow - 1, k).Value = fAddHeaderCounter(.Cells(vFRow - 1, k).Value, 1)
	            End If
NextForK:
	          Next
	          .Cells(arrPair(j), vLCol + 2).Value = vTxt
	          vTxt = ""
	        Next
	      End If

	      If fArrCnt(arrPair) = 1 Then ' если до этого был одинокий элемент
	        .Cells(arrPair(0), vColSimilarity).Interior.ColorIndex = 15
	        Debug.Print CStr(i) & " * нет пары для " & .Cells(arrPair(0)).Value
	      End If
	
	      arrPair = Array() ' обнуление для новой итерацией
	      Call fPopArr(arrPair, i) ' первый элемент новой итерации
	      vLastPair = .Cells(i, vColSimilarity).Value ' установка селектора новой пары
			'==========================================
		Next
  End With
End Sub

'===================================================================================================
' Сортировка пузырьковая
Sub pBubbleSort(SrcWS, vFRow, vLRow, vLCol, vLogCurrentFileName)
  vColSrc = 2
  vColTrg = 3
  vColSkp = 0

  With SrcWS
    For i = vFRow To vLRow
      If vColSkp > 0 Then If Len(.Cells(i, vColSkp).Value) > 0 Then GoTo NextFor
    	.Cells(i, vColTrg).Value = BubbleSort(Split(.Cells(i, vColSrc).Value, ", "))
      'Call fLogToFile("this", Now & " * pGenNomNoSpec R" & CStr(i) & " = " & .Cells(i, vColSrc).Value & " >>> " & vTmpResult, vLogCurrentFileName)
NextFor:
    Next
  End With
End Sub

'===================================================================================================
' Конвертирование регистра букв
Sub pConvertRegistr(SrcWS, vFRow, vLRow, vLCol, vLogCurrentFileName)
  Dim arrV As Variant
  Dim arrVold As Variant
  Set dP = New Dictionary
      dP.Add "ColSrc", 0
      dP.Add "ColTrg", 0
  
  For j = 1 To 5
    dP.item("ColSrc") = j
    dP.item("ColTrg") = j
    
    With SrcWS
      arrV = Range(.Cells(vFRow, dP.item("ColSrc")), .Cells(vLRow, dP.item("ColSrc")))
      arrVold = arrV
      For i = 1 To UBound(arrV)
        arrV(i, 1) = ConvertRegistr(arrV(i, 1), 1)
        'Debug.Print CStr(i) & " = " & arrVold(i, 1) & " >>> " & arrV(i, 1)
      Next
      Range(.Cells(vFRow, dP.item("ColTrg")), .Cells(vLRow, dP.item("ColTrg"))) = arrV
    End With
  Next
  ' Не использовать! Долго! Только для дебага
  'For i = 1 To UBound(arrV)
  '  Call fLogToFile("this", " * R" & CStr(i) & " = " & arrVold(i, 1) & " >>> " & arrV(i, 1), vLogCurrentFileName)
  'Next
End Sub

'===================================================================================================
' Удалить спец символы в одном столбце для нахождения дублей
Sub pGenNomNoSpecV2(SrcWS, vFRow, vLRow, vLCol, vLogCurrentFileName)
  Dim arrV As Variant
  Dim arrVold As Variant
  Dim arrSkip As Variant
  Set dP = New Dictionary
  dP.Add "ColSrc", 21
  dP.Add "ColTrg", 22
  dP.Add "ColSkp", 0
  vRegExSpcChr = "([A-Za-z0-9|])+|шт|компл*"
  With SrcWS
    arrV = Range(.Cells(vFRow, dP.item("ColSrc")), .Cells(vLRow, dP.item("ColSrc")))
		If CInt(dP.item("ColSkp")) <> 0 Then
			arrSkip = Range(.Cells(vFRow, dP.item("ColSkp")), .Cells(vLRow, dP.item("ColSkp")))
		End If
    arrVold = arrV
    For i = 1 To UBound(arrV)
			If CInt(dP.item("ColSkp")) <> 0 Then
				If CInt(dP.item("ColSkp")) > 0 And Len(arrSkip(i, 1)) > 0 Then GoTo NextFor
				If CInt(dP.item("ColSkp")) < 0 And Len(arrSkip(i, 1)) = 0 Then GoTo NextFor
			End If
    	vTmpResult = ""
			vTmp = fSearchRegExArr(arrV(i, 1), vRegExSpcChr, , , , 2)
    	If Len(Join(vTmp, " ")) > 0 Then
        For m = 0 To UBound(vTmp)
          If Not fIsEven(m) Then vTmpResult = vTmpResult & vTmp(m)
        Next
      	arrV(i, 1) = vTmpResult
    	End If
NextFor:
      'Debug.Print CStr(i) & " = " & arrVold(i, 1) & " >>> " & arrV(i, 1)
    Next
    Range(.Cells(vFRow, dP.item("ColTrg")), .Cells(vLRow, dP.item("ColTrg"))) = arrV
  End With
  ' Не использовать! Долго! Только для дебага
  'For i = 1 To UBound(arrV)
  '  Call fLogToFile("this", " * R" & CStr(i) & " = " & arrVold(i, 1) & " >>> " & arrV(i, 1), vLogCurrentFileName)
  'Next
End Sub

'===================================================================================================
' Удалить спец символы в одном столбце для нахождения дублей
Sub pGenNomNoSpec(SrcWS, vFRow, vLRow, vLCol, vLogCurrentFileName)
  vColSrc = 3
  vColTrg = 6
  vColSkp = 0
  vRegExSpcChr = "([A-Za-z0-9|])+|шт|компл*"
  With SrcWS
    For i = vFRow To vLRow
      If vColSkp > 0 Then If Len(.Cells(i, vColSkp).Value) > 0 Then GoTo NextFor
      vTmpResult = ""
      'ByVal str As String, ByVal Ptrn As String, Optional ByVal pMatch As Integer, Optional ByVal pSubMatch As Integer, Optional ByVal pCase As Boolean, Optional ByVal pMod As Integer) As Variant
      vTmp = fSearchRegExArr(.Cells(i, vColSrc).Value, vRegExSpcChr, , , , 2)
      If Len(Join(vTmp, " ")) > 0 Then
        For m = 0 To UBound(vTmp)
          If Not fIsEven(m) Then vTmpResult = vTmpResult & vTmp(m)
        Next
        .Cells(i, vColTrg).Value = vTmpResult
        'Call fLogToFile("this", Now & " * pGenNomNoSpec R" & CStr(i) & " = " & .Cells(i, vColSrc).Value & " >>> " & vTmpResult, vLogCurrentFileName)
      End If
NextFor:
      '.Cells(i, vColTrg).Value = fSearchRegEx(.Cells(i, vColSrc).Value, vRegExSpcChr, , -1, True, False, 1) ' str, Ptrn, pMatch, pSubMatch, pGlobal, pCase, pMod
    Next
  End With
End Sub

'===================================================================================================
' Суммирование объединенных ячеек
Sub pMergeCelOp(SrcWS, vFRow, vLRow, vLCol, vLogCurrentFileName)
  Set dP = New Dictionary
  dP.Add "ColSrc", 13
  dP.Add "ColTrg", 15
  With SrcWS
    For i = vFRow To vLRow
      vLngMrg = .Cells(i, dP.item("ColTrg")).MergeArea.Cells.Count
      vTmpRng = Range(.Cells(i, dP.item("ColSrc")), .Cells((i + vLngMrg) - 1, dP.item("ColSrc")))
      Debug.Print CStr(i) & " = " & CStr(WorksheetFunction.Sum(vTmpRng))
      .Cells(i, dP.item("ColTrg")).Value = WorksheetFunction.Sum(vTmpRng)
      i = (i + vLngMrg) - 1
      '.Cells(i, dP.Item("ColTrg").Value
    Next
  End With
End Sub

'===================================================================================================
' Удалить спец символы
Sub pFixSpecChr(SrcWS, vFRow, vLRow, vLCol, vLogCurrentFileName) ' vbCrLf is the same as Chr(13) and Chr(10)
      'dColRegEx.Add "Index", 1       ' Порядковый номер в накладной
      'dColRegEx.Add "Price", 7       ' Цена
      'dColRegEx.Add "Sum", 8         ' Сумма
      'dColRegEx.Add "WghtUnit", 9    ' Вес одной единицы
      'dColRegEx.Add "WghtTotal", 10  ' Общий вес
  Dim LogRegExSPC As String

  Set dP = New Dictionary
      dP.Add "Src", 13 ' column

  Dim dColRegEx
  Set dColRegEx = CreateObject("Scripting.Dictionary")
      dColRegEx.Add "2", Array("Vendor", "^(?=^(?:(?!  +).)*$)(?!\s)[A-Z ]{1,50}(?:[A-Z])$")                                                  ' Брэнд
      dColRegEx.Add "3", Array("Code", "^(?=^(?:(?! +).)*$)[A-Z0-9][A-Z0-9()+-\/.]{1,50}$")                                                   ' Артикул
      dColRegEx.Add "4", Array("Units", "^(компл|шт)$")                                                                                       ' ЕД
      dColRegEx.Add "5", Array("Quantity", "^(?!0)[0-9]{1,3}$")                                                                              ' Комплектность
      dColRegEx.Add "6", Array("OEM", "^[A-Z0-9][A-Z0-9\/|-]{1,200}$")                                                                       ' Кросс на оригинал
      dColRegEx.Add "7", Array("VendProNam", "^(?=^(?:(?! +).)*$)[A-Z0-9][A-Z0-9()+-\/.]{1,50}$")                                                   ' Артикул
      dColRegEx.Add "8", Array("TypeEng", "^(?=^(?:(?!  +).)*$)(?!\s)[A-Za-z ]{1,50}(?:[A-Z])$")                                                  ' Брэнд
     'dColRegEx.Add "13", Array("OEMF",       "^[A-Z0-9][A-Z0-9-]{1,50}$")                                                                    ' Кросс первый ГЕНЕРИТСЯ ФОРМУЛОЙ НА ЛИСТЕ, НЕ НУЖНО ПРОВЕРЯТЬ!
      dColRegEx.Add "9", Array("Type", "^(?=^(?:(?!  +).)*$)([А-Я][А-Яа-я0-9\/| ]{0,200})$")                                                  ' Название (тип) детали
      dColRegEx.Add "10", Array("Descr1", "(?!\s)(?=^(?:(?!  +).)*$)^([A-Za-zА-Яа-я0-9ъ])?([A-Za-zА-Яа-я0-9()_ +\/:№,.-])*([A-Za-zА-Яа-я0-9\)]{1,50}){1}$") ' Описание в наим
      dColRegEx.Add "11", Array("Size", "^([0-9]{1,3})((\.[0-9]{0,6})?(\/|x)([0-9]{1,3})(\.[0-9]{0,6}|\/[0-9]{1,3})?)+$")                     ' Размерности
      dColRegEx.Add "12", Array("Engine", "^(?=.*[A-Z].*)(([A-Z])|((?!^\-)(([A-Z0-9-]){1,9}(\/[A-Z0-9-]{1,9})*)))$")                          ' Движки (?<!\-)$
      dColRegEx.Add "14", Array("CrossBrand", dColRegEx.item("8")(1))                                                                         ' КроссБрэнд
      dColRegEx.Add "15;16;17", Array("Group1-3", dColRegEx.item("9")(1))                                                                     ' Группы 1-3
      dColRegEx.Add "18", Array("TNVED", "(?!^[0-1])^[0-9]{10}$")                                                                             ' ТНВЭД
      dColRegEx.Add "19", Array("Origin", "(?=^(?:(?!  +).)*$)(?!^[ ,()])(?![ ,(]$)([А-Я ,]{3,50})([А-Я]{2,50}|\([А-Я]{2,50}\)){1}$")         ' Страна происхождения
      'dColRegEx.Add "18", Array("Comment", dColRegEx.Item("8")(1))                                                                           ' Комментарий
      dColRegEx.Add "21", Array("GTD", "^[0-9]{8}\/[0123][0-9][01][0-9][01][0-9]\/[0-9]{7}$")                                                 ' ГТД
  vRegExSpcChr = "[^A-Za-zА-Яа-я0-9()_ \/:№+,.-]"

  Set RegExpSpace = CreateObject("vbscript.regexp")
      RegExpSpace.Global = True
      RegExpSpace.Pattern = " +" '+ означает, что заменяем и два и три и четыре и более пробелов подряд

  Dim vTmp As Variant

  With SrcWS
    For i = vFRow To vLRow
      ' If i > 10 Then Exit For
      LogRegExSPC = "|"
      For Each vColArr In dColRegEx.Keys
        For Each vCol In Split(vColArr, ";")
          vCol = CInt(vCol) + dP.item("Src") - 1
          If Len(.Cells(i, vCol).Value) = 0 Or .Cells(i, vCol).HasFormula Then ' не убирать, иначе дальше побьются все формулы тримом пробелов!
            If .Cells(i, vCol).HasFormula Then
              Debug.Print Now & " * pFixSpecChr ERROR R" & CStr(i) & " C" & CStr(vCol) & " has formula"
              Call fLogToFile("this", " * pFixSpecChr ERROR R" & CStr(i) & " C" & CStr(vCol) & " has formula", vLogCurrentFileName)
            End If
            GoTo SkipFor3
          End If
'====
          .Cells(i, vCol).Value = Trim(RegExpSpace.Replace(.Cells(i, vCol).Value, " "))
'====
          vTmp = fSearchRegEx(.Cells(i, vCol).Value, vRegExSpcChr, , -1, True, False, 1) ' str, Ptrn, pMatch, pSubMatch, pGlobal, pCase, pMod
          If Len(vTmp) Then
            Debug.Print Now & " * pFixSpecChr R" & CStr(i) & " C" & CStr(vCol) & " has spec char " & vTmp
            Call fLogToFile("this", " * pFixSpecChr R" & CStr(i) & " C" & CStr(vCol) & " has spec char " & vTmp, vLogCurrentFileName)
            LogRegExSPC = Trim(LogRegExSPC & " SPC: C" & CStr(vCol) & " " & vTmp & " |")
          End If
'====
          vTmp = fSearchRegEx(.Cells(i, vCol).Value, dColRegEx.item(vColArr)(1), , -1, True, False, 1) ' str, Ptrn, pMatch, pSubMatch, pGlobal, pCase, pMod
          If Len(vTmp) = 0 Then
            Debug.Print Now & " * pFixSpecChr R" & CStr(i) & " C" & CStr(vCol) & " incorrect pattern " & .Cells(i, vCol).Value
            Call fLogToFile("this", " * pFixSpecChr R" & CStr(i) & " C" & CStr(vCol) & " incorrect pattern " & .Cells(i, vCol).Value, vLogCurrentFileName)
            LogRegExSPC = LogRegExSPC & " " & dColRegEx.item(vColArr)(0) & " C" & CStr(vCol) & " " & .Cells(i, vCol).Value & " |"
          End If
'====
SkipFor3:
        Next
        If Len(LogRegExSPC) > 1 Then .Cells(i, vLCol + 2).Value = Trim(LogRegExSPC)
      Next
SkipFor1:
    Next
  End With
End Sub

          'Dim arrRmvSpc() As Variant
          'arrRmvSpc = Array(0, 1, 2, 9, 10, 13)
          '
          '.Cells(i, j).Value = Trim(.Cells(i, j).Value)
          'vTxtLog = ""
          'vTmp = ""
          'vTmp = fSearchRegExArr(.Cells(i, j).Value, arrPat(k))
          'If Len(Join(vTmp, " ")) > 0 Then
          '  If Len(vTxtLog) = 0 Then
          '    vTxtLog = "C" & CStr(j) & " R" & CStr(i)
          '  End If
          '  vTxtLog = vTxtLog & " <" & Join(vTmp, "> <") & ">"
          '  For m = 0 To UBound(vTmp)
          '    If Not fIsEven(m) Then
          '      For N = 0 To UBound(arrRmvSpc)
          '        If vTmp(m) = Chr(arrRmvSpc(N)) Then
          '          Debug.Print CStr(i) & " " & CStr(j) & " " & Chr(arrRmvSpc(N))
          '          '.Cells(i, j).Value = Replace(.Cells(i, j).Value, Chr(arrRmvSpc(N)), "")
          '        End If
          '      Next
          '    End If
          '  Next
          'End If
          'If Len(vTxtLog) > 0 Then
          '  Debug.Print vTxtLog
          'End If

'===================================================================================================
' Великая ебля с количеством мест для таможни
' пример в папке D:\SharePoint\Документы\5-Поступление товара\
'   034 20190307 ТТН 1FAE 2019 00002 ORIGINAL PARTS Warsaw (в работе)\5 таможня
'   Информационный файл для таможни.xlsm
' Веса для таможни - смотри пример в тойже поставке в папке \2 оригиналы обрабатываемые
'   2019.03.20 УПАКОВОЧНЫЙ ВАРШАВА.xlsm

Sub pCountTNVEDcustom(SrcWS, vFRow, vLRow, vLCol, vLogCurrentFileName)
  Set dP = New Dictionary
  dP.Add "vFRow", 23 ' строка с первым тнвэдом
  dP.Add "vLRow", 24 ' строка с последним тнвэдом
  dP.Add "ColSrc", 3 ' места
  dP.Add "ColData", 1 ' тнвэды
  dP.Add "RowTrg", 24 ' первый оффсет (не трогать)
  dP.Add "OutputColUni", 6 ' куда выгружать уникальные коробки
  dP.Add "OutputColDup", 8 ' куда выгружать повторяющиеся коробки
  Set dDup = New Dictionary
  With SrcWS

  .Cells(dP.item("vFRow") - 1, dP.item("OutputColUni")).Value = "Коробки уникальные"
  .Cells(dP.item("vFRow") - 1, dP.item("OutputColDup")).Value = "Коробки повторяющиеся"
  .Cells(dP.item("vFRow"), dP.item("OutputColUni")).Value = .Cells(dP.item("vFRow"), dP.item("ColSrc")).Value
  .Cells(dP.item("vFRow"), dP.item("OutputColDup")).Value = ""

  For j = dP.item("RowTrg") To dP.item("vLRow")
    Set dDup = New Dictionary
    Set dP.item("DBRange") = Range(.Cells(2, dP.item("ColSrc")), .Cells(j - 1, dP.item("ColSrc")))
    txtUni = ""
    For Each arrCheckVals In Split(.Cells(j, dP.item("ColSrc")).Value, ", ")
      fSkipped = 0
      txtDup = ""
      For Each Cell In dP.item("DBRange")
        If IsInArray(CStr(arrCheckVals), Split(Cell, ", ")) Then
          txtDup = dDup.item(.Cells(Cell.Row, dP.item("ColData")).Value)
          If Len(txtDup) = 0 Then
            dDup.item(.Cells(Cell.Row, dP.item("ColData")).Value) = arrCheckVals
          Else
            dDup.item(.Cells(Cell.Row, dP.item("ColData")).Value) = txtDup & ", " & arrCheckVals
          End If
          fSkipped = 1
          Exit For
        End If
      Next
      ' Debug.Print Cell.Row & " = " & Len(Cell.Value)
      If fSkipped = 0 Then
        If Len(txtUni) = 0 Then txtUni = arrCheckVals Else txtUni = txtUni & ", " & arrCheckVals
      End If
    Next
    txtDupCel = ""
    If dDup.Count > 0 Then
      For Each vBox In dDup.Keys
        If Len(txtDupCel) = 0 Then
          txtDupCel = dDup.item(vBox) & " часть по коду " & vBox
        Else
          txtDupCel = txtDupCel & ", " & dDup.item(vBox) & " часть по коду " & vBox
        End If
      Next
    End If
    .Cells(j, dP.item("OutputColUni")).Value = txtUni
    .Cells(j, dP.item("OutputColDup")).Value = txtDupCel
  Next
  
  End With
End Sub

'===================================================================================================
' Для поршни и гильщы TEIKIN IZUMI выделить и удалить окончания паттернами -S/F -F/F -L -R -A -AG -G -025 -025 -075 -100 -STD LSF* и т.д.
Sub zWRITERegex()
  Debug.Print String(65535, vbCr)
  Set RegExp = CreateObject("vbscript.regexp")
 
  Dim matches
  GetStringInParens = ""
  vSheet = 1
  vColumn = 3
  vSkipColumn = 0
  vResultColumn = 6
  vCheckColumn = 4
  aInclude = "" ' массив паттернов через ";" | начинается с №1 | если не нужно пропускать паттерны - оставить переменную пустой
' "(-\d{3}|-STD)$" ' "^.+?(?= [А-Яа-я])" ' "((\s|^)|\d{1}SET=)(\d{1,2}PCS)(\s|$)" ' "[0-9]{5,9}/[0123][0-9][01][0-9][01][0-9]/[A-Za-zА-Яа-я]?[0-9]{5,8}"
  pSPatternArr = "^(L[A-Z]{2})(?=\d{5});" & _
                  "(-L|-R|-LR)*(-A|-AG|-G)*(-\d{3}|-STD)$;" & _
                  "(-\d{3}|-STD)$;" & _
                  "(-[FS]/F)$;" & _
                  "(-L|-R|-LR)*(-S|-G)$;" & _
                  "(-S|-G)+(-NOK)$;" & _
                  "(-TH|-MY)$;" & _
                  "(-IN/EX|-EX|-IN)$"

        pSCheck = ";" & _
                  "Поршень двигателя;" & _
                  "Кольца поршневые;" & _
                  ";" & _
                  ";" & _
                  "Ремкомплект двигателя;" & _
                  "GP;" & _
                  "Клапан двигателя / Клапан двигателя набор / Направляющая клапана / Прокладка коллектора"
                  ' паттерн №2 = "Поршень двигателя с пальцем / Кольца поршневые;"

  With Application
  .EnableEvents = False: .ScreenUpdating = False

  Set WS = ThisWorkbook.Sheets(vSheet)
  Set Rng = WS.Range(WS.Cells(2, vColumn), WS.Cells(Cells.Find("*", [A1], , , xlByRows, xlPrevious).Row, vColumn)) ' WS.Cells(1798, vColumn)) '
  With RegExp
    For Each Cell In Rng
      tmptmp = Cell.Row
      If tmptmp = 212 Then
        Fuckit = 1
      End If
      If vSkipColumn > 0 Then If Len(Cells(Cell.Row, vSkipColumn).Value) > 0 Then GoTo NextCell
      flagGotResult = 0
      i = 0
      j = 0

      For Each pSPattern In Split(pSPatternArr, ";")
        proceed = 0
        If Len(aInclude) > 0 Then If Not IsInArray(CStr(j + 1), Split(aInclude, ";")) Then GoTo NextPattern
        If Len(Split(pSCheck, ";")(j)) > 0 Then
          vTmpTxt = Split(pSCheck, ";")(j)
          If InStr(vTmpTxt, Cells(Cell.Row, vCheckColumn).Value) = 0 Then
            For Each vTmpTxt In Split(vTmpTxt, " / ")
              If InStr(Cells(Cell.Row, vCheckColumn).Value, vTmpTxt) > 0 Then
                proceed = 1
                Exit For
              End If
            Next
            If proceed = 0 Then GoTo NextPattern
          End If
        End If
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = pSPattern
        If Cell.Value <> "" Then
          If .Test(Cell.Value) And i = 0 Then
            i = i + 1
            Set matches = .Execute(Cell.Value)
            Debug.Print "R" & Cell.Row & " Ptrn" & CStr(j + 1) & " < " & matches(0) & " > in < " & Cell.Value & " >" ' & matches(0).submatches(2)
            If Len(matches(0)) Then
              Cells(Cell.Row, vResultColumn).Value = StrReverse(Replace(StrReverse(Cell.Value), StrReverse(matches(0)), "", , 1, vbTextCompare)) ' matches(0) '
              flagGotResult = 1
              Exit For
            End If
            'For Each subs In matches(0).SubMatches
            '  Debug.Print " - Sub = " & subs
            'Next
          Else
            ' Debug.Print Cell.Row & " = NULL " & Cell.Value
            ' Cells(Cell.Row, vResultColumn).Value = Cell.Value
          End If
        End If
NextPattern:
        j = j + 1
      Next
      If flagGotResult = 0 Then Cells(Cell.Row, vResultColumn).Value = ""
NextCell:
    Next
  End With
  'Debug.Print CStr(i)

  .EnableEvents = True: .ScreenUpdating = True
  ' DoEvents
  End With
End Sub

' ТЕСТЫ
Sub zTestFRegEx()
  Set cTxts = New Collection
  cTxts.Add ("18.01.2017 23:18:49 Выгрузка ТОРГ12 2605")
  cTxts.Add ("Загрузка 03.02.2017 03:27:18 ТТН 3929118 №1535")
  cTxts.Add ("Загрузка 05.06.2018 03:27:18 ТТН 3929118 05.06.2015 №1535")
  For Each txt In cTxts
    If A = "" Then
      A = CDate(fSearchRegEx(txt, "[0123][0-9].(0|1)[0-9].(19|20)[0-9]{2}", , 1, False, True, 2))
      Debug.Print CStr(A)
    ElseIf A < CDate(fSearchRegEx(txt, "[0123][0-9].(0|1)[0-9].(19|20)[0-9]{2}", , 1, False, True, 2)) Then
      A = CDate(fSearchRegEx(txt, "[0123][0-9].(0|1)[0-9].(19|20)[0-9]{2}", , 1, False, True, 2))
      Debug.Print "a > " & CStr(A)
    End If
  Next
End Sub
