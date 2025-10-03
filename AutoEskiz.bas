Attribute VB_Name = "AutoEskiz"
' 2017-06-26 Воронов М.В.
' поправил добавление контакторов
' поправил добавление реле

' 2017-08-03 Воронов М.В.
' поправил добавление кнопки

' 2017-10-24 Воронов М.В.
' поправил добавление кнопки

' 2018-05-19 serega
' всё удалил и написал заново

' 2018-07-02 serega
' убрали текс для клемм
' обнулили пробелы между элементами

' 2018-08-02 Воронов М.В.
' Между лампочками и между переключателями на дверцы пробел равный их ширине чтоб расставлять удобно было
' Добавил переключатель на 5 положений, сам шейп ещё подредактировать надо будет чтоб надписи выводил и номер

' 2019-10-30 Воронов М.В.
' Сегодня мне надоело писать эти слова Width & Height и я заменил их на заглавные бувы W и H



Sub AutoEskiz()
 
    Const VSS_NAME = "10_эскиз шкафа_1к4.vss"
    Dim vShape As Visio.Shape
    Dim vShapes As Visio.Shapes
    Dim vPage As Visio.Page
    Dim NewShape As Visio.Shape
    Dim sType As String
    Dim sNum As Integer
    Dim sStateNum As Integer
    Dim sColor As Integer
    Dim sPolusNum As Integer
    Dim sUserPolusNum As Integer
    Dim sName As String
    Dim sModel As String
    Dim sColorCaption As String
    Dim sCaption As String
    Dim sCaptionMain As String
    Dim sCaption1 As String
    Dim sCaption2 As String
    Dim sCaption3 As String
    
    Static cSort As String
    Dim SortA As Integer, SortB As Integer
        
    Dim i As Integer
    Dim NumStr As String
    
    Dim IndexVar As Variant
    Dim X As Variant
    Dim Y As Variant
    Dim Picture As String
    Dim ShapeWidth As Single
    Dim ShapeHeight As Single
    Dim sIndex As Integer
    
    ' Порядок shapetype в массивах: 0-other, 1-HL, 2-SA, 3-SB, 4-QF, 5-K, 6-KM
    ' Массив названий мастер-шейпов для каждой группы по умолчанию
    ' 2П, 1p и 2p заменяются на нужное количество полюсов АВТОМАТИЧЕСКИ
    PictArray = Array("Void", "HL_v1", "SA_2П_v1", "SB", "QF_1p", "K_2p", "KM_2p")
    MarginArray = Array(0, 7 / 25.4, 7 / 25.4, 7 / 25.4, 0, 0, 0) ' horizontal gaps between adjacent shapes ' пробелы между элементами
    ' Массивы положения первого элемента каждой группы
    X = Array(0, 0, 0, 0, 0, 0, 0)
    Y = Array(0, -1, -2, -3, -4, -5, -6)
               
    
    
If cSort = "" Then cSort = "1-99"
cSort = InputBox("Введите номер страницы или интервал (напр. 1-3)", "Делаем эскиз", cSort)
If InStr(cSort, "-") > 0 Then
  SortA = CInt(Left(cSort, InStr(cSort, "-") - 1))
  SortB = CInt(Right(cSort, Len(cSort) - InStr(cSort, "-")))
 Else
  SortA = Val(cSort)
  SortB = SortA
End If
If (SortA < 1) Or (SortB < 1) Then Exit Sub
    
 For Each vPage In ActiveDocument.Pages ' перебираем страницы и шейпы
  If (vPage.Index >= SortA) And (vPage.Index <= SortB) Then
   Set vShapes = vPage.Shapes
     For Each vShape In vShapes
       sType = ""
       If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
        sType = vShape.CellsU("User.ShapeType").ResultStr("")
        If vShape.CellExistsU("User.ShapeNum", 0) Then NumStr = CStr(CInt(vShape.CellsU("User.ShapeNum").ResultStr(""))) Else NumStr = ""
        IndexVar = Switch(sType = "HL", 1, sType = "SA", 2, sType = "SB", 3, sType = "QF", 4, sType = "K", 5, sType = "KM", 6)
        'для недвухполюсных контакторов рисуем прямоугольник
        If (sType = "KM") Then If (vShape.CellsU("Prop.PolusNum").ResultStr("") <> "2") Then IndexVar = Null
        
        If (Not IsNull(IndexVar)) Then
            sIndex = CInt(IndexVar)
            If vShape.CellExistsU("User.ShapeNum", 0) Then sNum = vShape.CellsU("User.ShapeNum").ResultStr("")
            If vShape.CellExistsU("User.StateNum", 0) Then sStateNum = vShape.CellsU("User.StateNum").ResultStr("")
            If vShape.CellExistsU("Prop.Color", 0) Then sColor = vShape.CellsU("Prop.Color").ResultStr("")
            If vShape.CellExistsU("User.ColorCaption", 0) Then sColorCaption = vShape.CellsU("User.ColorCaption").ResultStr("")
            If vShape.CellExistsU("Prop.Caption", 0) Then sCaption = vShape.CellsU("Prop.Caption").ResultStr("")
            If vShape.CellExistsU("Prop.CaptionMain", 0) Then sCaptionMain = vShape.CellsU("Prop.CaptionMain").ResultStr("")
            If vShape.CellExistsU("Prop.Caption1", 0) Then sCaption1 = vShape.CellsU("Prop.Caption1").ResultStr("")
            If vShape.CellExistsU("Prop.Caption2", 0) Then sCaption2 = vShape.CellsU("Prop.Caption2").ResultStr("")
            If vShape.CellExistsU("Prop.Caption3", 0) Then sCaption3 = vShape.CellsU("Prop.Caption3").ResultStr("")
            If vShape.CellExistsU("Prop.Caption4", 0) Then sCaption2 = vShape.CellsU("Prop.Caption4").ResultStr("")
            If vShape.CellExistsU("Prop.Caption5", 0) Then sCaption3 = vShape.CellsU("Prop.Caption5").ResultStr("")
            If vShape.CellExistsU("User.PolusNum", 0) Then sUserPolusNum = vShape.CellsU("User.PolusNum").ResultStr("")
            If vShape.CellExistsU("Prop.PolusNum", 0) Then sPolusNum = vShape.CellsU("Prop.PolusNum").ResultStr("")
            Picture = PictArray(sIndex)
            ' Здесь подставляются названия эскизных мастер-шейпов с соответствующим числом полюсов
            If (sType = "SA") And (sStateNum = 3) Then Picture = Replace(Picture, "2П", "3П")
            If (sType = "SA") And (sStateNum = 5) Then Picture = Replace(Picture, "2П", "5П")
            If (sType = "QF") Then Picture = Replace(Picture, "1p", CStr(sUserPolusNum) & "p")
            If (sType = "K") Or (sType = "KM") Then Picture = Replace(Picture, "2p", CStr(sPolusNum) & "p")
            ShapeWidth = Application.Documents.Item(VSS_NAME).Masters.ItemU(Picture).Shapes.ItemU(1).Cells("Width")
            ' метод Drop принимает координаты центра (pinX и pinY) шейпа, поэтому дропаем со смещением на половину ширины
            Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item(VSS_NAME).Masters.ItemU(Picture), X(sIndex) + ShapeWidth / 2, Y(sIndex))
            If NewShape.CellExistsU("User.ShapeNum", 0) Then NewShape.Cells("User.ShapeNum").Formula = sNum
            If NewShape.CellExistsU("Prop.Color", 0) Then NewShape.Cells("Prop.Color").Formula = sColor
            If NewShape.CellExistsU("Prop.Caption", 0) Then NewShape.Cells("Prop.Caption").Formula = Chr(34) + sCaption + Chr(34)
            If NewShape.CellExistsU("Prop.CaptionMain", 0) Then NewShape.Cells("Prop.CaptionMain").Formula = Chr(34) + sCaptionMain + Chr(34)
            If NewShape.CellExistsU("Prop.Caption1", 0) Then NewShape.Cells("Prop.Caption1").Formula = Chr(34) + sCaption1 + Chr(34)
            If NewShape.CellExistsU("Prop.Caption2", 0) Then NewShape.Cells("Prop.Caption2").Formula = Chr(34) + sCaption2 + Chr(34)
            If NewShape.CellExistsU("Prop.Caption3", 0) Then NewShape.Cells("Prop.Caption3").Formula = Chr(34) + sCaption3 + Chr(34)
            If NewShape.CellExistsU("Prop.ColorCaption", 0) Then NewShape.Cells("Prop.ColorCaption").Formula = Chr(34) + sColorCaption + Chr(34)
            X(sIndex) = X(sIndex) + MarginArray(sIndex) + NewShape.Cells("Width")
        ElseIf (vShape.CellExistsU("User.EskizShape", 0)) Then
            sIndex = 0
            Picture = vShape.CellsU("User.EskizShape").ResultStr("")
            'MsgBox (Application.Documents.Item(VSS_NAME).Masters.ItemU(Picture).Shapes.ItemU(1).Name)
            ShapeWidth = Application.Documents.Item(VSS_NAME).Masters.ItemU(Picture).Shapes.ItemU(1).Cells("Width")
            Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item(VSS_NAME).Masters.ItemU(Picture), X(sIndex) + ShapeWidth / 2, Y(sIndex))
            NewShape.Text = sType + NumStr
            X(sIndex) = X(sIndex) + MarginArray(sIndex) + NewShape.Cells("Width")
        ElseIf vShape.CellExistsU("Prop.H", 0) And vShape.CellExistsU("Prop.W", 0) Then
          ' Width & Height is set или говоря по-русски задана высота и ширина для этого шэйпа (объекта/блока/девайса/устройства).
          sIndex = 0
          ShapeWidth = Round(vShape.CellsU("Prop.W") / 4)
          ShapeHeight = Round(vShape.CellsU("Prop.H") / 4)
          If vShape.CellExistsU("Prop.Name", 0) Then sName = vShape.CellsU("Prop.Name").ResultStr("") Else sName = ""
          If vShape.CellExistsU("Prop.Model", 0) Then sModel = vShape.CellsU("Prop.Model").ResultStr("") Else sModel = ""
          Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item(VSS_NAME).Masters.ItemU(PictArray(sIndex)), X(sIndex) + ShapeWidth / 50.8, Y(sIndex))
          NewShape.Cells("Width").FormulaU = CStr(ShapeWidth) & " mm"
          NewShape.Cells("Height").FormulaU = CStr(ShapeHeight) & " mm"
          NewShape.Text = sType + NumStr + Chr(10) + sName + Chr(10) + sModel
          X(sIndex) = X(sIndex) + MarginArray(sIndex) + NewShape.Cells("Width")
        End If
       End If ' ShapeType exists
       Next vShape
     End If
    Next vPage
    
    
End Sub














