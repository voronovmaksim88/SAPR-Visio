Attribute VB_Name = "AutoNum_v2"
' 2017-06-15 Воронов МВ
' Добавил переменную cXT
' причесал объявление переменных
' добавил автонумеацию кабеля

' 2017-06-21 Воронов МВ
' Добавил автонумерацию блоков питания

' 2017-06-25 Воронов МВ
' Добавил автонумерацию частотников

' 2017-07-14 Воронов МВ
' Добавил автонумерацию LineNum

' 2017-07-18 серега
' kabel -> CB
' LineNum -> LN

' 2017-07-24 Воронов МВ
' добавил SF

' 2017-08-10 Воронов МВ
' добавил TV

' 2017-08-30 Воронов МВ
' добавил THI

' 2017-09-06 Воронов МВ
' добавил QS

' 2017-10-23 Воронов МВ
' добавил THE
' добавил QI

' 2017-12-07 Воронов МВ
' добавил QFD

' 2018-01-10 Савинских АС
' Добавил тт-реле SSR

' 2018-02-05 Воронов МВ
' добавил LS

' 2018-02-16 Воронов МВ
' добавил KK

' 2018-03-07 Воронов МВ
' добавил A это всякие девайсы в основном ПЛК

' 2018-03-07 Воронов МВ
' добавил KL это клеммы


' 2018-04-29 Воронов МВ
' увеличил массивы
'Dim Arr(255) As ShapeRec
'Dim Pars(255) As ParentRec


' 2018-06-05 Воронов МВ
' добавил "Т" это автотрансфрматоры

' 2018-06-25 Воронов МВ
' исправил "KK"
' добавил "F" предохранители

' 2018-08-06 Воронов МВ
' добавил "HA" звуковой излучатель

' 2018-10-28 Воронов МВ
' добавил "VD" диод

' 2019-07-18 Воронов МВ
' "NumABC" это 3 номера фаз

' 2019-11-22 Воронов МВ
' hashid теперь при выносе на лист суммируется к тому который есть,
' это надо чтоб на одном шэйпе держать несколько пар


' 2019-11-28 серега
' копирование свойств секции prop из родительского шейпа в дочерние


' 2020-01-09 Воронов МВ
' Arr и Pars увеличены до 1024

' 2020-01-14 Воронов МВ
' FC получили дочек

' 2020-01-27 Воронов МВ
' TE получили дочек

' 2020-02-03 Воронов МВ
' Num_ABC получили дочек

' 2020-02-22 Serega
' Ungroup in DropK + Links

' 2020-04-21 Воронов МВ
' PS теперь работает через HashID
' TS теперь работает через HashID
' PE теперь работает через HashID
' PDS теперь работает через HashID

' 2020-04-24 Serega
' Able to change text field in Links

' 2020-12-21 Воронов МВ
' добавил KV

' 2023-08-29 Воронов МВ
' поправил PDE




Type ShapeRec
 fID As Integer
 fX As Single
 fY As Single
End Type

Type ParentRec
 fNum As Integer
 fHash As Long
 fShape As Visio.Shape
End Type

Type LinkRec
 fHash As Long
 fFirstPage As Integer
 fSecondPage As Integer
 fText As String
End Type



Sub DropK(vsoShape As Visio.Shape, Optional ungroupAfter As Integer = 1)

Dim vShape As Visio.Shape
Dim vShapes As Visio.Shapes
Dim hash As Long

Randomize
hash = 10000000 + Int(Rnd() * 80000000)

Set vShapes = vsoShape.Shapes

For Each vShape In vShapes
 If (vShape.CellExistsU("User.HashID", visExistsAnywhere)) Then vShape.Cells("User.HashID").Formula = vShape.Cells("User.HashID").Formula + hash
Next vShape

If (ungroupAfter = 1) Then vsoShape.Ungroup
 
End Sub



Sub AutoNum()
Attribute AutoNum.VB_ProcData.VB_Invoke_Func = "n"
' Сочетание клавиш: Ctrl+w
'

Dim Arr(1023) As ShapeRec
Dim Pars(1023) As ParentRec
Dim Links(1023) As LinkRec
Dim OneArr As ShapeRec
Dim i As Integer, hash As Long, Count As Integer, parsCount As Integer
Dim SortType As Integer
Dim SortA As Integer, SortB As Integer
Dim cType As String
Static cSort As String
'здесь перечисляем все нужные типы'
'''''''''''''''''''''''''''''''''''
Dim cSA As Integer, cSB As Integer, cQF As Integer, cKM As Integer
Dim cHL As Integer, cK As Integer, cXT As Integer, cCB As Integer
Dim cTS As Integer, cPS As Integer, cPDS As Integer, cHS As Integer
Dim cPE As Integer, cPDE As Integer, cHE As Integer, cTE As Integer
Dim cM As Integer, cFC As Integer, cLN As Integer, сUG As Integer
Dim cSF As Integer, cTV As Integer, сTHI As Integer, сQS As Integer
Dim cTHE As Integer, cQI As Integer, cQFD As Integer, cSSR As Integer
Dim cLS As Integer, cKK As Integer, cKT As Integer, cA As Integer
Dim cKL As Integer, cT As Integer, cF As Integer, cHA As Integer
Dim cVD As Integer, сNumABC As Integer, сNum As Integer, cLink As Integer
Dim cKV As Integer
Count = 0: nPage = 0

Dim vPage As Visio.Page
'Set vPage = Application.ActivePage
Dim vShape As Visio.Shape
Dim vShapes As Visio.Shapes
Dim CurPage As Integer
Dim LinkCrt As Boolean

If cSort = "" Then cSort = "1"
cSort = InputBox("Введите номер страницы или интервал (напр. 1-3)", "Индексация", cSort)
If InStr(cSort, "-") > 0 Then
  SortA = CInt(Left(cSort, InStr(cSort, "-") - 1))
  SortB = CInt(Right(cSort, Len(cSort) - InStr(cSort, "-")))
 Else
  SortA = Val(cSort)
  SortB = SortA
End If
If (SortA < 1) Or (SortB < 1) Then Exit Sub
SortType = MsgBox("Упорядочивать сначала по вертикали?", vbOKCancel, "Индексация")

For Each vPage In ActiveDocument.Pages
If (vPage.Index >= SortA) And (vPage.Index <= SortB) Then

Set vShapes = vPage.Shapes
CurPage = -1

For Each vShape In vShapes
If CurPage < 0 Then
    If vShape.CellExistsU("User.PageNum", visExistsAnywhere) Then CurPage = vShape.Cells("User.PageNum").Result("")
End If
If vShape.CellExistsU("User.LinkNum", visExistsAnywhere) Then
    vShape.Cells("User.HostPage").Formula = CurPage
    LinkCrt = True
    hash = vShape.Cells("User.HashID").Result("")
    For i = 1 To cLink
        If Links(i).fHash = hash Then
            Links(i).fSecondPage = CurPage
            If vShape.CellExistsU("Prop.Text", visExistsAnywhere) Then
                Links(i).fText = vShape.Cells("Prop.Text").ResultStr("")
            End If
            vShape.Cells("User.LinkNum").Formula = i
            LinkCrt = False
            Exit For
        End If
    Next i
    If LinkCrt Then
        cLink = cLink + 1
        Links(cLink).fHash = hash
        Links(cLink).fFirstPage = CurPage
        If vShape.CellExistsU("Prop.Text", visExistsAnywhere) Then
            Links(cLink).fText = vShape.Cells("Prop.Text").ResultStr("")
        End If
        vShape.Cells("User.LinkNum").Formula = cLink
    End If
End If

If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
 Count = Count + 1
 Arr(Count).fID = vShape.ID
 Arr(Count).fX = vShape.Cells("PinX").Result(visDrawingUnits)
 Arr(Count).fY = vShape.Cells("PinY").Result(visDrawingUnits)
 i = Count
 
 If SortType = 1 Then
    Do While (i > (nPage + 1)) And ((Arr(i).fY > (Arr(i - 1).fY + 0.5)) Or ((Abs(Arr(i).fY - Arr(i - 1).fY) < 0.5) And (Arr(i).fX < Arr(i - 1).fX)))
     OneArr = Arr(i - 1)
     Arr(i - 1) = Arr(i)
     Arr(i) = OneArr
     i = i - 1
    Loop
 Else
    Do While (i > (nPage + 1)) And ((Arr(i).fX < (Arr(i - 1).fX - 0.5)) Or ((Abs(Arr(i).fX - Arr(i - 1).fX) < 0.5) And (Arr(i).fY > Arr(i - 1).fY)))
     OneArr = Arr(i - 1)
     Arr(i - 1) = Arr(i)
     Arr(i) = OneArr
     i = i - 1
    Loop
 End If

End If

Next vShape

For i = nPage + 1 To Count
Set vShape = vShapes.ItemFromID(Arr(i).fID)
 If (vShape.CellExistsU("User.ShapeType", visExistsAnywhere)) And (vShape.CellExistsU("User.ShapeNum", visExistsAnywhere)) Then
    cType = vShape.Cells("User.ShapeType").Formula
    'здесь тоже перечисляем все типы'
    '''''''''''''''''''''''''''''''''
    Select Case cType
     
    Case """SA"""
        cSA = cSA + 1
        vShape.Cells("User.ShapeNum").Formula = cSA
     
    Case """SB"""
        cSB = cSB + 1
        vShape.Cells("User.ShapeNum").Formula = cSB
     
    Case """QF"""
        cQF = cQF + 1
        vShape.Cells("User.ShapeNum").Formula = cQF
     
    Case """KM"""
        cKM = cKM + 1
        vShape.Cells("User.ShapeNum").Formula = cKM
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cKM
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
     
    Case """HL"""
        cHL = cHL + 1
        vShape.Cells("User.ShapeNum").Formula = cHL
            
    Case """XT"""
        cXT = cXT + 1
        vShape.Cells("User.ShapeNum").Formula = cXT
     
    Case """K"""
        cK = cK + 1
        vShape.Cells("User.ShapeNum").Formula = cK
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cK
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
                
    Case """CB"""
        cCB = cCB + 1
        vShape.Cells("User.ShapeNum").Formula = cCB
        
    Case """TE"""
        cTE = cTE + 1
        vShape.Cells("User.ShapeNum").Formula = cTE
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cTE
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
     
    Case """TS"""
        cTS = cTS + 1
        vShape.Cells("User.ShapeNum").Formula = cTS
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cTS
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
    
    Case """PE"""
        cPE = cPE + 1
        vShape.Cells("User.ShapeNum").Formula = cPE
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cPE
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
    
    
    Case """PS"""
        cPS = cPS + 1
        vShape.Cells("User.ShapeNum").Formula = cPS
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cPS
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
    
    Case """HE"""
        cHE = cHE + 1
        vShape.Cells("User.ShapeNum").Formula = cHE
    
    Case """HS"""
        cHS = cHS + 1
        vShape.Cells("User.ShapeNum").Formula = cHS
    
    Case """PDE"""
        cPDE = cPDE + 1
        vShape.Cells("User.ShapeNum").Formula = cPDE
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cPDE
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
    
    Case """PDS"""
        cPDS = cPDS + 1
        vShape.Cells("User.ShapeNum").Formula = cPDS
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cPDS
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
        
    Case """M"""
        cM = cM + 1
        vShape.Cells("User.ShapeNum").Formula = cM
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cM
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
        
    Case """UG"""
        сUG = сUG + 1
        vShape.Cells("User.ShapeNum").Formula = сUG
                
    Case """FC"""
        сFC = сFC + 1
        vShape.Cells("User.ShapeNum").Formula = сFC
        parsCount = parsCount + 1
        Pars(parsCount).fNum = сFC
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
        
    Case """LN"""
        cLN = cLN + 1
        vShape.Cells("User.ShapeNum").Formula = cLN
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cLN
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
        
    Case """SF"""
        сSF = сSF + 1
        vShape.Cells("User.ShapeNum").Formula = сSF
        
    Case """TV"""
        сTV = сTV + 1
        vShape.Cells("User.ShapeNum").Formula = сTV
        
    Case """THI"""
        сTHI = сTHI + 1
        vShape.Cells("User.ShapeNum").Formula = сTHI
                
    Case """QS"""
        сQS = сQS + 1
        vShape.Cells("User.ShapeNum").Formula = сQS
        
    Case """THE"""
        сTHE = сTHE + 1
        vShape.Cells("User.ShapeNum").Formula = сTHE
    
    Case """QI"""
        сQI = сQI + 1
        vShape.Cells("User.ShapeNum").Formula = сQI
        
    Case """QFD"""
        сQFD = сQFD + 1
        vShape.Cells("User.ShapeNum").Formula = сQFD
     
    Case """SSR"""
        cSSR = cSSR + 1
        vShape.Cells("User.ShapeNum").Formula = cSSR
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cSSR
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
    
    Case """LS"""
        cLS = cLS + 1
        vShape.Cells("User.ShapeNum").Formula = cLS
    
    Case """KK"""
        cKK = cKK + 1
        vShape.Cells("User.ShapeNum").Formula = cKK
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cKK
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
     
     
    Case """KT"""
        cKT = cKT + 1
        vShape.Cells("User.ShapeNum").Formula = cKT
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cKT
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
        
    Case """A"""
        cA = cA + 1
        vShape.Cells("User.ShapeNum").Formula = cA
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cA
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
        
            
    Case """KL"""
        cKL = cKL + 1
        vShape.Cells("User.ShapeNum").Formula = cKL
        
                    
    Case """T"""
        cT = cT + 1
        vShape.Cells("User.ShapeNum").Formula = cT
        
                            
    Case """F"""
        cF = cF + 1
        vShape.Cells("User.ShapeNum").Formula = cF
        
    Case """HA"""
        cHA = cHA + 1
        vShape.Cells("User.ShapeNum").Formula = cHA
        
    Case """VD"""
        cVD = cVD + 1
        vShape.Cells("User.ShapeNum").Formula = cVD
        
    Case """Num"""
        сNum = сNum + 1
        vShape.Cells("User.ShapeNum").Formula = сNum
        
    Case """NumABC"""
        сNumABC = сNumABC + 1
        vShape.Cells("User.ShapeNum").Formula = сNumABC
        parsCount = parsCount + 1
        Pars(parsCount).fNum = сNumABC
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
        
    Case """KV"""
        cKV = cKV + 1
        vShape.Cells("User.ShapeNum").Formula = cKV
        parsCount = parsCount + 1
        Pars(parsCount).fNum = cKV
        Pars(parsCount).fHash = vShape.Cells("User.HashID").Result("")
        Set Pars(parsCount).fShape = vShape
     
    
     
     End Select
 End If
Next i
nPage = Count

End If
Next vPage

''''' Searching for secondary shapes and numering them '''''
For Each vPage In ActiveDocument.Pages
If (vPage.Index >= SortA) And (vPage.Index <= SortB) Then

Set vShapes = vPage.Shapes
Dim iRF As Integer, iRS As Integer, inx As Integer
Dim vsoCellF As Visio.Cell, vsoCellS As Visio.Cell
Dim RowExists As Boolean

For Each vShape In vShapes
    If vShape.CellExistsU("User.LinkNum", visExistsAnywhere) Then
     inx = vShape.Cells("User.LinkNum").Result("")
     If (inx > 0) Then
     
      If (Links(inx).fSecondPage >= 0) Then
        vShape.Text = "Лист " + CStr(Links(inx).fSecondPage) + ": "
        Links(inx).fSecondPage = -1
      Else
        vShape.Text = "Лист " + CStr(Links(inx).fFirstPage) + ": "
      End If
      
      If vShape.CellExistsU("Prop.Text", visExistsAnywhere) Then
        Set vsoCharacters = vShape.Characters
        vsoCharacters.Begin = vsoCharacters.End
        vsoCharacters.AddCustomField "Prop.Text", visFmtStrNormal
      Else
        vShape.Text = vShape.Text + Links(inx).fText
      End If
     End If
    End If

    If vShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) And vShape.CellExistsU("User.HashID", visExistsAnywhere) Then
     For i = 1 To parsCount
      If Pars(i).fHash = vShape.Cells("User.HashID").Result("") Then
       vShape.Cells("User.SecondaryShapeNum").Formula = Pars(i).fNum
        For iRF = 0 To Pars(i).fShape.RowCount(visSectionProp) - 1 'Primary
         Set vsoCellF = Pars(i).fShape.CellsSRC(visSectionProp, iRF, 0)
         RowExists = False
          For iRS = 0 To vShape.RowCount(visSectionProp) - 1 'Secondary
            Set vsoCellS = vShape.CellsSRC(visSectionProp, iRS, 0)
                If vsoCellS.RowName = vsoCellF.RowName Then
                    vsoCellS.FormulaU = vsoCellF.FormulaU
                    RowExists = True
                    Exit For
                End If
          Next iRS 'Secodary
          ' Comment This Block To Prevent Creating New Rows
          ' BLOCK START '
         If RowExists = False Then
            inx = vShape.RowCount(visSectionProp) + 1
            vShape.AddRow visSectionProp, inx, visTagDefault
            vShape.CellsSRC(visSectionProp, inx - 1, 0).RowName = vsoCellF.RowName
            vShape.CellsSRC(visSectionProp, inx - 1, visCustPropsLabel).FormulaU = Pars(i).fShape.CellsSRC(visSectionProp, iRF, visCustPropsLabel).FormulaU
            Set vsoCellS = vShape.CellsSRC(visSectionProp, inx - 1, 0)
            vsoCellS.FormulaU = vsoCellF.FormulaU
         End If
          ' BLOCK END '
        Next iRF 'Primary
       Exit For
      End If
     Next i
    End If
Next vShape

End If
Next vPage
LonelyParen
LonelyChild
End Sub

Sub LonelyParen()
' по этому событию ищутся одинокие родители
    Dim visSheet As Visio.Page
    Dim visShape As Visio.Shape
    
    ' Перебор всех листов в текущем документе
    For Each visSheet In ThisDocument.Pages
        ' Перебор всех шейпов на текущем листе
        For Each visShape In visSheet.Shapes
            ' Ваш код для работы с текущим шейпом
            If visShape.CellExistsU("User.HashID", visExistsAnywhere) And Not visShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) Then
            ' Это условие выбирает шейпы родители
                Dim CyrrentHashID As Long
                CyrrentHashID = visShape.Cells("User.HashID").Result("")
                            
                If Not AvailabilityChild(CyrrentHashID) Then
                    MsgBox "найден родитель без ребёнка"
                    ' Получение координат "PinX" и "PinY" шейпа
                    Dim xCoord As Double
                    Dim yCoord As Double
                    Dim Shape_Oval As Visio.Shape
            
                    xCoord = visShape.Cells("PinX").ResultIU
                    yCoord = visShape.Cells("PinY").ResultIU
                    Radius = 0.4   ' Радиус круга
                    ' Создание нового круга на шейпе
                    Set Shape_Oval = visSheet.DrawOval(xCoord - Radius, yCoord - Radius, xCoord + Radius, yCoord + Radius)
                
                    ' Настройка внешнего вида круга, например, цвет заливки и границы
                    Shape_Oval.Cells("LineColor").FormulaU = "RGB(255, 0, 0)"  ' Синий цвет границы
                    Shape_Oval.Cells("LineWeight").FormulaU = "0.05" ' Толщина границы
                    Shape_Oval.CellsU("FillPattern").FormulaU = 0 ' Убрать заливку
                    'Shape_Oval.Text = CyrrentHashID
                    'Shape_Oval.Text = "Родитель"
                End If
            End If
        Next visShape
    Next visSheet
End Sub

Function AvailabilityChild(ParentHashID As Long) As Boolean
' эта функция проверяет наличие детей с заданным HashID
    AvailabilityChild = False
    Dim visSheet As Visio.Page
    Dim visShape As Visio.Shape
    
    ' Перебор всех листов в текущем документе
    For Each visSheet In ThisDocument.Pages
        ' Перебор всех шейпов на текущем листе
        For Each visShape In visSheet.Shapes
            ' Ваш код для работы с текущим шейпом
            If visShape.CellExistsU("User.HashID", visExistsAnywhere) And visShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) Then
            ' Это условие выбирает шейпы дети
                Dim CyrrentHashID As Double
                CyrrentHashID = visShape.Cells("User.HashID").Result("")
                If CyrrentHashID = ParentHashID Then
                    AvailabilityChild = True
                End If
            End If
        Next visShape
    Next visSheet
End Function
Sub LonelyChild()
' по этому событию ищутся одинокие дети
    Dim visSheet As Visio.Page
    Dim visShape As Visio.Shape
    
    ' Перебор всех листов в текущем документе
    For Each visSheet In ThisDocument.Pages
        ' Перебор всех шейпов на текущем листе
        For Each visShape In visSheet.Shapes
            ' Ваш код для работы с текущим шейпом
            If visShape.CellExistsU("User.HashID", visExistsAnywhere) And visShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) Then
            ' Это условие выбирает шейпы детей
                Dim CyrrentHashID As Long
                CyrrentHashID = visShape.Cells("User.HashID").Result("")
                            
                If Not AvailabilityParent(CyrrentHashID) Then
                    MsgBox "найден ребёнок без родителя"
                    ' Получение координат "PinX" и "PinY" шейпа
                    Dim xCoord As Double
                    Dim yCoord As Double
                    Dim Shape_Oval As Visio.Shape
            
                    xCoord = visShape.Cells("PinX").ResultIU
                    yCoord = visShape.Cells("PinY").ResultIU
                    Radius = 0.4   ' Радиус круга
                    ' Создание нового круга на шейпе
                    Set Shape_Oval = visSheet.DrawOval(xCoord - Radius, yCoord - Radius, xCoord + Radius, yCoord + Radius)
                
                    ' Настройка внешнего вида круга, например, цвет заливки и границы
                    Shape_Oval.Cells("LineColor").FormulaU = "RGB(255, 0, 0)"  ' Синий цвет границы
                    Shape_Oval.Cells("LineWeight").FormulaU = "0.05" ' Толщина границы
                    Shape_Oval.CellsU("FillPattern").FormulaU = 0 ' Убрать заливку
                    'Shape_Oval.Text = CyrrentHashID
                    'Shape_Oval.Text = "Родитель"
                End If
            End If
        Next visShape
    Next visSheet
End Sub

Function AvailabilityParent(ParentHashID As Long) As Boolean
' эта функция проверяет наличие родителей с заданным HashID
    AvailabilityParent = False
    Dim visSheet As Visio.Page
    Dim visShape As Visio.Shape
    
    ' Перебор всех листов в текущем документе
    For Each visSheet In ThisDocument.Pages
        ' Перебор всех шейпов на текущем листе
        For Each visShape In visSheet.Shapes
            ' Ваш код для работы с текущим шейпом
            If visShape.CellExistsU("User.HashID", visExistsAnywhere) And Not visShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) Then
            ' Это условие выбирает шейпы родителей
                Dim CyrrentHashID As Double
                CyrrentHashID = visShape.Cells("User.HashID").Result("")
                If CyrrentHashID = ParentHashID Then
                    AvailabilityParent = True ' родитель найден
                End If
            End If
        Next visShape
    Next visSheet
End Function
