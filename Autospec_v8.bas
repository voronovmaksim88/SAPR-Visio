Attribute VB_Name = "Autospec_v8"
' 2017-06-15 Воронов МВ
' переименовал старые переменные
' начал пилить обработчик К и КМ
' Шрифт везде теперь "10 pt" мм

' 2017-06-19 Воронов МВ
' подпилил КМ
' добавил в спеку кабель, но потом его надо будет убрать и сделать отдельной таблтцей

' 2017-06-20 Воронов МВ
' Добавляем привода - М

' 2017-06-21 серега
' autospec поправил

' 2017-06-21 Воронов МВ
' Добавляем определение шрифта и размера текста для заголовков столбцов

' 2017-06-26 Воронов МВ
' Мелкие доработки

' 2017-07-12 серега
' добавил sub ExportToFile()

' 2017-07-18 серега
' kabel -> CB
' LineNum -> LN
' Form with shapetype choice

' 2017-07-25 Воронов МВ
' добавил SF

' 2017-08-10 Воронов МВ
' добавил TV

' 2017-08-30 Воронов МВ
' добавил THI

' 2017-08-31 Отургашев СЕ
' теперь правильно выбирает шейпы, только те что есть на выбранных страницах

' 2017-09-04 Воронов МВ
' если нет свойства name то теперь макрос не падает
' добавил FC

' 2017-09-06 Воронов МВ
' поправил SB

' 2017-10-10 Воронов МВ
' прокачал SB

' 2017-10-12 Отургашев СЕ
' жирные заголовки, разные строчки для разных примечаний, высота строки

' 2017-10-23 Воронов МВ
' Добавил THE и QI

' 2017-12-08 Воронов МВ
' Добавил THE и QI

' 2018-01-10 Воронов МВ
' Добавил шины SHN - шины L1, L2, L3, N, PE

' 2018-01-10 Савинских АС
' Добавил тт-реле SSR

' 2018-02-05 Воронов МВ
' Добавил LS

' 2018-02-16 Воронов МВ
' Добавил KK

' 2018-03-09 Воронов МВ
' Добавил A это контроллеры и другие девайсы

' 2018-04-13 Воронов МВ
' Name для Box прописал

' 2018-04-20 Воронов МВ
' Box исправил на BOX

' 2018-04-28 Воронов МВ
' Для SB (кнопок) сделал что имя тепрь копируется из свойства Name
' KL - добавил имя

' 2018-06-06 Воронов МВ
' добавил  SS  это датчик дыма дискретный smoke

' 2018-06-30 serega
' добавил цвет в спецификацию

' 2019-02-16 Воронов МВ
' Для автоматов вывел ном.откл способность

' 2019-06-17 Воронов МВ
' Тепловое реле теперь выводится с током

' 2019-12-12 Воронов МВ
' имя для КТ

' 2020-04-03 Воронов МВ
' имя для PDI

' 2020-04-06 Воронов МВ
' увеличил ширину примечаний чтоб больше влазило

' 2023-08-23 Воронов МВ
' исключил из спецификации номера ссылок на провода между страницами
' исключил из спецификации номера проводов (хотя можно было бы их и ставить конечно),
' чтобы они указывали на необходимость проверить наличие маркировки,
' но пока так пусть

' 2025-03-20 Воронов МВ
' добавил комментариев
' напромтил механизм сжатия текста в первом столбце
' теперь вместо "КМ1, КМ2, КМ3, КМ4, КМ5" будет "КМ1-КМ5"


Public ShapeTypeExceptions As Collection
'Это пользовательский тип данных для хранения информации о строке спецификации.
Type TRowStruct
    rsPos As String        ' Позиционное обозначение
    rsDenom As String      ' Наименование
    rsManuf As String      ' Производитель
    rsModel As String      ' Модель
    rsNote As String       ' Примечание
    rsQty As Integer       ' Количество
    rsKey As Integer       ' Ключ для сортировки
    rsColor As Integer     ' Цвет для отображения в таблице
End Type
    
Function CreateR(Pos, Denom, Manuf, Model, Note, Num) As TRowStruct
'Создает новую структуру строки спецификации, заполняя её данными.
 CreateR.rsPos = Pos
 CreateR.rsDenom = Denom
 CreateR.rsManuf = Manuf
 CreateR.rsModel = Model
 CreateR.rsNote = Note
 CreateR.rsQty = 1
 CreateR.rsColor = 0
 CreateR.rsKey = Asc(Pos) * 100 + CInt(Num)
End Function

' Функция для модификации позиционного обозначения
Function OptimizeRsPos(originalRsPos As String) As String
    ' Проверка на пустую строку
    If Trim(originalRsPos) = "" Then
        OptimizeRsPos = ""
        Exit Function
    End If
    
    ' Разделяем входную строку на массив
    Dim items() As String
    items = Split(originalRsPos, ", ")
    
    ' Сортировка массива
    Call SortEquipmentArray(items)
    
    ' Поиск последовательностей и оптимизация
    Dim result As String
    Dim i As Integer
    Dim j As Integer
    Dim currentPrefix As String
    Dim currentNumber As Integer
    Dim sequenceLength As Integer
    Dim inSequence As Boolean
    Dim sequenceStart As Integer
    
    result = ""
    i = 0
    
    Do While i < UBound(items) + 1
        ' Получаем префикс и номер для текущего элемента
        currentPrefix = GetPrefix(items(i))
        currentNumber = GetNumberValue(items(i))
        
        ' Начинаем проверку последовательности
        sequenceLength = 1
        sequenceStart = i
        j = i + 1
        
        ' Проверяем, есть ли дальше последовательность
        Do While j < UBound(items) + 1
            Dim nextPrefix As String
            Dim nextNumber As Integer
            
            nextPrefix = GetPrefix(items(j))
            nextNumber = GetNumberValue(items(j))
            
            ' Если префикс тот же и номер следующий по порядку
            If currentPrefix = nextPrefix And nextNumber = currentNumber + 1 Then
                sequenceLength = sequenceLength + 1
                currentNumber = nextNumber
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        
        ' Добавляем элементы в результат
        If sequenceLength >= 4 Then
            ' Если последовательность достаточно длинная, используем сокращение
            If result <> "" Then result = result & ", "
            result = result & items(sequenceStart) & "-" & items(j - 1)
            i = j
        Else
            ' Иначе добавляем текущий элемент как есть
            If result <> "" Then result = result & ", "
            result = result & items(i)
            i = i + 1
        End If
    Loop
    
    OptimizeRsPos = result
End Function

' Функция для сортировки массива оборудования
Private Sub SortEquipmentArray(Arr() As String)
    Dim i As Integer, j As Integer
    Dim temp As String
    
    For i = LBound(Arr) To UBound(Arr) - 1
        For j = i + 1 To UBound(Arr)
            If CompareEquipment(Arr(i), Arr(j)) > 0 Then
                temp = Arr(i)
                Arr(i) = Arr(j)
                Arr(j) = temp
            End If
        Next j
    Next i
End Sub

' Функция сравнения двух элементов оборудования
Private Function CompareEquipment(item1 As String, item2 As String) As Integer
    Dim prefix1 As String, prefix2 As String
    Dim num1 As Integer, num2 As Integer
    
    prefix1 = GetPrefix(item1)
    prefix2 = GetPrefix(item2)
    
    ' Сначала сравниваем префиксы
    If prefix1 < prefix2 Then
        CompareEquipment = -1
        Exit Function
    ElseIf prefix1 > prefix2 Then
        CompareEquipment = 1
        Exit Function
    End If
    
    ' Если префиксы равны, сравниваем числовые значения
    num1 = GetNumberValue(item1)
    num2 = GetNumberValue(item2)
    
    If num1 < num2 Then
        CompareEquipment = -1
    ElseIf num1 > num2 Then
        CompareEquipment = 1
    Else
        CompareEquipment = 0
    End If
End Function

' Функция для извлечения префикса из строки (например, "KM" из "KM3")
Private Function GetPrefix(item As String) As String
    Dim i As Integer
    Dim result As String
    
    result = ""
    For i = 1 To Len(item)
        If IsNumeric(Mid(item, i, 1)) Then
            Exit For
        Else
            result = result & Mid(item, i, 1)
        End If
    Next i
    
    GetPrefix = result
End Function

' Функция для извлечения числового значения из строки (например, 3 из "KM3")
Private Function GetNumberValue(item As String) As Integer
    Dim i As Integer
    Dim NumStr As String
    
    NumStr = ""
    For i = 1 To Len(item)
        If IsNumeric(Mid(item, i, 1)) Then
            NumStr = NumStr & Mid(item, i, 1)
        End If
    Next i
    
    If NumStr = "" Then
        GetNumberValue = 0
    Else
        GetNumberValue = CInt(NumStr)
    End If
End Function

Function DenomStr(pShape As Visio.shape) As String
'Функция, которая формирует текст наименования компонента в зависимости от его типа.
'Для разных типов элементов (HL - лампы, QF - автоматы, UG - блоки питания и т.д.)
'формируется соответствующий текст с параметрами из свойств элемента.
Dim str(8) As String
Dim pType As String
pType = pShape.CellsU("User.ShapeType").ResultStr("")
 Select Case pType
  
Case "HL" ' лампочки
    str(0) = pShape.CellsU("Prop.Up").ResultStr("")
    str(1) = pShape.CellsU("User.ColorCaption").ResultStr("")
    DenomStr = "Световой индикатор (" + str(0) + " В) " + str(1)

Case "QF" ' автоматы
    str(0) = Round(pShape.CellsU("User.PolusNum").ResultStr(""))
    str(0) = str(0) + "П, "
    str(1) = "х-ка " + pShape.CellsU("Prop.Characteristic").ResultStr("")
    str(2) = ", Iн= " + pShape.CellsU("Prop.Current").ResultStr("") + "А"
    str(3) = ", ном. откл. спос. " + pShape.CellsU("Prop.Nom_Otkl_Spos").ResultStr("") + "кА"
    DenomStr = "Автоматический выключатель, " + str(0) + str(1) + str(2) + str(3)

Case "UG" ' блоки питания
    str(0) = pShape.CellsU("Prop.Power").ResultStr("")
    DenomStr = "Блок питания ( ~220\=24, " + str(0) + " Вт) "

Case "XT" ' клеммы
    str(0) = pShape.CellsU("Prop.Area").ResultStr("")
    DenomStr = "Клеммная группа, " + str(0) + " мм.кв."

Case "SA" ' переключатели
    str(0) = Round(pShape.CellsU("User.StateNum").ResultStr(""))
    DenomStr = "Преключатель на " + str(0) + " положения"
    
Case "TE", "TS", "PE", "PS", "PDE", "PDS", "HE", "HS", "M", "FC", "THE", "QI", "KL", "SS", "KT", "PDI", "QE"
  If pShape.CellExistsU("Prop.Name", visExistsAnywhere) Then
   DenomStr = pShape.CellsU("Prop.Name").ResultStr("")
  Else
   DenomStr = "?"
  End If

Case "K" 'реле
    str(0) = pShape.CellsU("Prop.PolusNum").ResultStr("")
    DenomStr = "Реле, " + str(0) + "-х пол."

Case "KM"
    str(0) = pShape.CellsU("Prop.Current").ResultStr("")
    DenomStr = "Контактор, ток до " + str(0) + "А по х-ке АС3"

Case "CB" 'кабели
    DenomStr = ""
    
Case "SF" 'АЗД автоматы защиты двигателя
    str(0) = pShape.CellsU("Prop.Current").ResultStr("")
    DenomStr = "Автомат защиты двигателя, ток " + str(0)
    
Case "TV" 'трансформаторы
    str(0) = pShape.CellsU("Prop.Uin").ResultStr("")
    str(1) = pShape.CellsU("Prop.Uout").ResultStr("")
    str(2) = pShape.CellsU("Prop.Power").ResultStr("")
    DenomStr = "Трансформатор (Uвх=" + str(0) + ", Uвых=" + str(1) + ", P=" + str(2) + ")"
    
Case "THI" 'интерфейсные датчики темпы и влажности
    DenomStr = pShape.CellsU("Prop.Name").ResultStr("")
    
Case "QS" 'Рубильники
    str(0) = pShape.CellsU("Prop.Current").ResultStr("")
    DenomStr = "Рубильник, ток " + str(0) + "А"
    
Case "SB" 'Кнопки
    DenomStr = pShape.CellsU("Prop.Name").ResultStr("")
    
Case "QFD" 'Диф автоматы
    DenomStr = "Дифференциальный автомат"
    
Case "SHN" 'Шины
    DenomStr = pShape.CellsU("Prop.Name").ResultStr("")

Case "SSR" 'ТТР твердотельные реле
    str(0) = pShape.CellsU("Prop.PolusNum").ResultStr("")
    DenomStr = "Твердотельное реле, " + str(0) + "-х пол."
    
Case "LS" 'не помню уже
    DenomStr = pShape.CellsU("Prop.Name").ResultStr("")

Case "A" 'тоже не помню
    DenomStr = pShape.CellsU("Prop.Name").ResultStr("")
    
Case "BOX" 'корпуса шкафов
    DenomStr = pShape.CellsU("Prop.Name").ResultStr("")
    
Case "KK" 'тепловые реле
    str(0) = pShape.CellsU("Prop.Current").ResultStr("")
    DenomStr = "Тепловое реле (" + str(0) + " A)"
 End Select
End Function

Function ModelStr(pShape As Visio.shape) As String
'Извлекает информацию о модели из свойств элемента.
 If pShape.CellExistsU("Prop.Model", visExistsAnywhere) Then
   ModelStr = pShape.CellsU("Prop.Model").ResultStr("")
  Else
   ModelStr = "?"
  End If
End Function

Function EqualR(a As TRowStruct, b As TRowStruct) As Boolean
'Сравнивает две структуры строк спецификации.
 EqualR = True
 If (a.rsDenom <> b.rsDenom) Or (a.rsManuf <> b.rsManuf) Or (a.rsModel <> b.rsModel) Or (a.rsNote <> b.rsNote) Then
     EqualR = False
 End If
End Function

Function CollectionContains(myCol As Collection, checkVal As Variant) As Boolean
'Проверяет, содержится ли значение в коллекции.
    On Error Resume Next
    CollectionContains = False
    Dim it As Variant
    For Each it In myCol
        If it = checkVal Then
            CollectionContains = True
            Exit Function
        End If
    Next
End Function
 
Sub Autospec()
    Const DEFAULT_H As Double = 5 / 25.4 ' default string height
    
    Dim COL_WIDTH(6) As Integer
    COL_WIDTH(0) = 30       ' Position
    COL_WIDTH(1) = 120      ' Name
    COL_WIDTH(2) = 40       ' Manufacturer
    COL_WIDTH(3) = 40       ' Model
    COL_WIDTH(4) = 90       ' Note
    COL_WIDTH(5) = 20       ' Quantity
    
    Dim COL_STR_LEN(6) As Integer
    COL_STR_LEN(0) = 15       ' Position (not used)
    COL_STR_LEN(1) = 60       ' Name
    COL_STR_LEN(2) = 20       ' Manufacturer
    COL_STR_LEN(3) = 20       ' Model
    COL_STR_LEN(4) = 20       ' Note
    COL_STR_LEN(5) = 10       ' Quantity
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    ' delete this cycle to set individual COL_STR_LEN
    For i = 0 To 5
        COL_STR_LEN(i) = Int(COL_WIDTH(i) / 2)
    Next i

 
    Dim Rows(1 To 64) As TRowStruct ' array of table rows info
    Dim RowsCount As Integer
    Dim FindRow As Integer
    RowsCount = 0: FindRow = 0
    Dim myRowStruct As TRowStruct
    Dim Symbol As String
       
    
    Dim vShape As Visio.shape
    Dim vShape1 As Visio.shape
    Dim Tabl(0 To 5) As Visio.shape
    Dim vShapes As Visio.Shapes
    Dim MainShape As Visio.shape
    Dim MainSelection As Visio.Selection
    Static cSort As String
    Dim SortA As Integer, SortB As Integer
    
    Dim Denomination As String ' расчётная переменная, по ней в частности счиатаем высоту строки
    Dim Model As String
    Dim Manuf As String
    Dim Note As String
    Dim tStr(0 To 5) As String
    
    Dim flag As Boolean
    Dim rsi As Integer
    Dim Y As Double 'координата по у
    
    Dim h As Double
    Dim h1 As Double
    Dim maxHeight As Double
    
    Dim X(7) As Double
    X(0) = 10 / 25.4 'начальный отступ по х
    For i = 0 To 5  ' set cols boundaries
        X(i + 1) = X(i) + COL_WIDTH(i) / 25.4
    Next i
    
    ' clear collection with specification exceptions
    Set ShapeTypeExceptions = New Collection
    
    h = 2 * DEFAULT_H
    Y = 200 / 25.4
    
    If cSort = "" Then cSort = "1-99"
    cSort = InputBox("Введите номер страницы или интервал (напр. 1-3)", "Спецификация", cSort)
    If InStr(cSort, "-") > 0 Then
        SortA = CInt(Left(cSort, InStr(cSort, "-") - 1))
        SortB = CInt(Right(cSort, Len(cSort) - InStr(cSort, "-")))
     Else
        SortA = Val(cSort)
        SortB = SortA
     End If
    If (SortA < 1) Or (SortB < 1) Then Exit Sub
    
    Call Form_Spec.ShowW(SortA, SortB)
    
    
    ActiveWindow.DeselectAll
    Set MainSelection = ActiveWindow.Selection
    
    For i = 0 To 5
      Set Tabl(i) = ActivePage.DrawRectangle(X(i), Y, X(i + 1), Y - h)
      If (COL_WIDTH(i) = 0) Then Tabl(i).Cells("HideText").FormulaU = "TRUE"
    Next i
        
        
' определяем шрифт и размер заголовков
For i = 0 To 5
    If Not (Tabl(i) Is Nothing) Then
    ' set table's font name and size
    If Tabl(i).CellExistsU("Char.Size", visExistsAnywhere) Then Tabl(i).Cells("Char.Size").FormulaU = "14 pt"
    If Tabl(i).CellExistsU("Char.Style", visExistsAnywhere) Then Tabl(i).Cells("Char.Style").FormulaU = "1"
    If Tabl(i).CellExistsU("Char.Font", visExistsAnywhere) Then Tabl(i).Cells("Char.Font").FormulaU = "FONTTOID(""Calibri"")"
    End If
Next i
    
    ' прописываем заголовки
    Tabl(0).Text = "Поз."
    Tabl(1).Text = "Наименование"
    Tabl(2).Text = "Изготовитель"
    Tabl(3).Text = "Марка"
    Tabl(4).Text = "Примечание"
    Tabl(5).Text = "Коли- чество"
    
        
    Y = Y - h
    
    For i = 0 To 5
        If Not (Tabl(i) Is Nothing) Then
        MainSelection.Select Tabl(i), visSelect
        End If
    Next i
        
    Dim vPage As Visio.Page
    
    For Each vPage In ActiveDocument.Pages ' Для каждой страницы в документе
    If (vPage.Index >= SortA) And (vPage.Index <= SortB) Then ' если её  номер принадлежит диапазону страниц по которым пользователь просит сделать пеку
    Set vShapes = vPage.Shapes ' vShapes присваиваем всешейпы на странице
     For Each vShape In vShapes ' для каждого шейпа на странице....
        If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then ' если у шейпа есть свойство ShapeType, то ....
            If vShape.CellsU("User.ShapeType").ResultStr("") <> "LineNum" And _
            vShape.CellsU("User.ShapeType").ResultStr("") <> "LN" And _
            vShape.CellsU("User.ShapeType").ResultStr("") <> "NumABC" Then  ' не отображаем в спеке номера проводов и ссылки между страницами
             If Not CollectionContains(ShapeTypeExceptions, vShape.CellsU("User.ShapeType").ResultStr("")) Then
              Symbol = vShape.CellsU("User.ShapeType").ResultStr("") + Format(vShape.CellsU("User.ShapeNum").ResultStr(""))
              Denomination = DenomStr(vShape)
              Model = ModelStr(vShape)
              Manuf = "?"
              If vShape.CellExistsU("prop.Manufacturer", visExistsAnywhere) Then Manuf = vShape.CellsU("prop.Manufacturer").ResultStr("")
                Note = ""
              If vShape.CellExistsU("prop.Note", visExistsAnywhere) Then Note = vShape.CellsU("prop.Note").ResultStr("")
                myRowStruct = CreateR(Symbol, Denomination, Manuf, Model, Note, Format(vShape.CellsU("User.ShapeNum").ResultStr("")))
              For i = 1 To RowsCount
               If EqualR(myRowStruct, Rows(i)) Then
                Rows(i).rsQty = Rows(i).rsQty + 1
                Rows(i).rsPos = Rows(i).rsPos + ", " + Symbol
                FindRow = 1
                Exit For
               End If
              Next i
              If (FindRow = 0) Then
               If (Manuf = "?") Or (Model = "?") Then myRowStruct.rsColor = 2 'red
               RowsCount = RowsCount + 1
               Rows(RowsCount) = myRowStruct
              End If
              FindRow = 0
             End If
            End If
           End If
     Next vShape
    End If
    Next vPage
    
    For i = 1 To RowsCount ' Sorting table
      flag = False
      For j = 1 To (RowsCount - i)
        If Rows(j).rsKey > Rows(j + 1).rsKey Then
         myRowStruct = Rows(j)
         Rows(j) = Rows(j + 1)
         Rows(j + 1) = myRowStruct
         flag = True
        End If
      Next j
      If Not flag Then
       Exit For
      End If
    Next i
    
    For rsi = 1 To RowsCount ' Table print cycle
                
        ' Сначала объявляем переменную для хранения максимальной высоты
        
        maxHeight = DEFAULT_H ' начальное значение
        
        tStr(0) = OptimizeRsPos(Rows(rsi).rsPos)
        
        For i = 0 To 5 ' Теперь включаем и первую ячейку в расчет
          If i = 0 Then
            ' Для первой ячейки (позиционные обозначения)
            Dim charsPerLine As Integer
            charsPerLine = 12 ' примерно 12 символов в строке для позиционных обозначений
            j = Len(tStr(0))
            h1 = DEFAULT_H * Int((j + charsPerLine - 1) / charsPerLine)
            ' Обновляем максимальную высоту
            If h1 > maxHeight Then maxHeight = h1
          Else
            Select Case i
            Case 1
              tStr(1) = Replace(Rows(rsi).rsDenom, "&", Chr(10))
            Case 2
              tStr(2) = Replace(Rows(rsi).rsManuf, "&", Chr(10))
            Case 3
              tStr(3) = Replace(Rows(rsi).rsModel, "&", Chr(10))
            Case 4
              tStr(4) = Replace(Rows(rsi).rsNote, "&", Chr(10))
            Case 5
              tStr(5) = Rows(rsi).rsQty
            End Select
            j = Len(tStr(i))
            If (InStr(tStr(i), Chr(10)) > 0) Then
              k = COL_STR_LEN(i) * (Len(tStr(i)) - Len(Replace(tStr(i), Chr(10), "")) + 1)
              If (k > j) Then j = k
            End If
            h = DEFAULT_H * Int((j + COL_STR_LEN(i) - 1) / COL_STR_LEN(i))
            ' Обновляем максимальную высоту
            If h > maxHeight Then maxHeight = h
          End If
        Next i
                
                
                For i = 0 To 5
                  Set Tabl(i) = ActivePage.DrawRectangle(X(i), Y, X(i + 1), Y - maxHeight)
                  If (COL_WIDTH(i) = 0) Then Tabl(i).Cells("HideText").FormulaU = "TRUE"
                  
                  If Not (Tabl(i) Is Nothing) Then
                  ' set table's font name, size and color
                     If Tabl(i).CellExistsU("Char.Size", visExistsAnywhere) Then Tabl(i).Cells("Char.Size").FormulaU = "10 pt"
                     If Tabl(i).CellExistsU("Char.Font", visExistsAnywhere) Then Tabl(i).Cells("Char.Font").FormulaU = "FONTTOID(""Calibri"")"
                     If (Rows(rsi).rsColor <> 0) And (Tabl(i).CellExistsU("Char.Color", visExistsAnywhere)) Then Tabl(i).Cells("Char.Color").FormulaU = Rows(rsi).rsColor
                  End If
                  Tabl(i).Text = tStr(i)
                Next i
                              
             
                Y = Y - maxHeight
                
       For i = 0 To 5
        If Not (Tabl(i) Is Nothing) Then
         MainSelection.Select Tabl(i), visSelect
        End If
       Next i
    
    Next rsi
    
    Set MainShape = MainSelection.group
    
End Sub

Sub ExportToFile()

Dim FileName As String


    Dim i As Integer
    Dim j As Integer
    
    Dim Rows(1 To 128) As TRowStruct ' array of table rows info
    Dim RowsCount As Integer
    RowsCount = 0
    Dim myRowStruct As TRowStruct
    Dim Symbol As String
    
    Dim vShape As Visio.shape
    Dim vShapes As Visio.Shapes
    Static cExp As String
    Dim SortA As Integer, SortB As Integer
    
    Dim Denomination As String ' расчётная переменная, по ней в частности считаем высоту строки
    Dim Model As String
    Dim Manuf As String
    Dim Note As String
        
    Dim flag As Boolean
    Dim rsi As Integer
    
FileName = "details_list"
FileName = InputBox("Введите имя файла для экспорта (без расширения)", "Экспорт", FileName)
FileName = FileName + ".csv"

    
    If cExp = "" Then cExp = "1-99"
    cExp = InputBox("Введите номер страницы или интервал (напр. 1-3)", "Экспорт", cExp)
    If InStr(cExp, "-") > 0 Then
        SortA = CInt(Left(cExp, InStr(cExp, "-") - 1))
        SortB = CInt(Right(cExp, Len(cExp) - InStr(cExp, "-")))
     Else
        SortA = Val(cExp)
        SortB = SortA
     End If
    If (SortA < 1) Or (SortB < 1) Then Exit Sub
              
    Dim vPage As Visio.Page
    
    For Each vPage In ActiveDocument.Pages ' Creating list cycle
    If (vPage.Index >= SortA) And (vPage.Index <= SortB) Then
    Set vShapes = vPage.Shapes
     For Each vShape In vShapes
        If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
          Symbol = vShape.CellsU("User.ShapeType").ResultStr("") + Format(vShape.CellsU("User.ShapeNum").ResultStr(""))
          Denomination = DenomStr(vShape)
          Model = ModelStr(vShape)
          Manuf = "?"
          If vShape.CellExistsU("prop.Manufacturer", visExistsAnywhere) Then Manuf = vShape.CellsU("prop.Manufacturer").ResultStr("")
          Note = ""
          If vShape.CellExistsU("prop.Note", visExistsAnywhere) Then Note = vShape.CellsU("prop.Note").ResultStr("")
          myRowStruct = CreateR(Symbol, Denomination, Manuf, Model, Note, Format(vShape.CellsU("User.ShapeNum").ResultStr("")))
          'If (myRowStruct.rsModel <> "") And (myRowStruct.rsModel <> "?") Then
            RowsCount = RowsCount + 1
            Rows(RowsCount) = myRowStruct
           
          'End If ' rsModel <> "?" or ""
        End If ' CellExistsU("User.ShapeType")
     Next vShape
    End If
    Next vPage
    
    For i = 1 To RowsCount ' Sorting table
      flag = False
      For j = 1 To (RowsCount - i)
        If Rows(j).rsKey > Rows(j + 1).rsKey Then
         myRowStruct = Rows(j)
         Rows(j) = Rows(j + 1)
         Rows(j + 1) = myRowStruct
         flag = True
        End If
      Next j
      If Not flag Then
       Exit For
      End If
    Next i
    
 Open FileName For Output As #1
    
    For rsi = 1 To RowsCount ' Table print cycle
            
         Print #1, Rows(rsi).rsManuf; ";"; Rows(rsi).rsModel; ";"; Rows(rsi).rsNote
                    
    Next rsi

Close #1


End Sub

