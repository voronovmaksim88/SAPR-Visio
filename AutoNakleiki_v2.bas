Attribute VB_Name = "AutoNakleiki_v2"
'2017-09-29 Воронов МВ
'добавил SF

'2017-10-17 Воронов МВ
'добавил SB

'2017-10-17 Воронов МВ
'добавил KT

'2018-05-23 Воронов МВ
'добавил SSR
'добавил KM
'добавил A

'2018-05-31 serega
'поменял всё после 'НАКЛЕЙКИ НА ЭЛЕМЕНТЫ

'2018-06-01 Воронов МВ
'переименовал в v2
' добавил "K" - реле

'2018-06-20 Воронов МВ
' добавил "KL" - клемма, они в виде исключений

'2018-07-05 Воронов МВ
' добавил "KK" - тепловухи
' добавил "TV" - транс
' добавил "F" - предохранитель

'2018-07-05 Воронов МВ
' добавил "M" - вентилятор шкафа

'2018-10-04 Воронов МВ
' добавил "QFD" - диавтомат
' добавил "QS" - рубильник

'2020-09-21 Воронов МВ
' закоментил наклейки на клеммы

'2020-09-26 Воронов МВ
' маркировка реле 14*7





Type TSticker
    Type As String
    Prefix As String
    Count As Integer
    FontSize As Integer
    width As Single
    Height As Single
End Type


Sub AutoNakleiki()
    Const GROUPS_COUNT = 17  ' количество различных групп наклеек на элементы
    Dim vShape As Visio.Shape
    Dim vShapes As Visio.Shapes
    Dim vPage As Visio.Page
    Dim pNumber As Integer
    Dim pCaption As String
    Dim pCaptionMain As String
    Dim pCaption1 As String
    Dim pCaption2 As String
    Dim pCaption3 As String
    Dim NewNaklieka As Shape
    Dim Stickers(GROUPS_COUNT) As TSticker
    Dim sNumbers(GROUPS_COUNT, 256) As Integer
    
    
    Dim X As Double
    Dim X0 As Double    ' сдвиг для надписей под держатель
    Dim Y As Double
    Dim i As Integer
    Dim k As Integer
    Dim group As Integer
    Dim w_HL As Double ' ширина наклейки
    Dim h_HL As Double ' высота наклейки
       
    Dim w_KL As Double ' ширина наклейки на клемму
    Dim h_KL As Double ' высота наклейки на клемму
          
    ' при добавлении новых групп — увеличивай GROUPS_COUNT
    GroupsArray = Array("HL", "SA", "SB", "QF", "KM", "KT", "A", "SF", "SSR", "K", "KK", "TV", "F", "M", "QFD", "QS", "UG")
    WidthArray = Array(20, 20, 20, 18, 18, 18, 18, 17, 18, 14, 16, 18, 7, 18, 18, 18, 18)
    HeightArray = Array(15, 15, 15, 10, 6, 10, 10, 7, 10, 7, 5, 10, 7, 10, 10, 10, 10)
    FontArray = Array(20, 20, 20, 16, 14, 16, 16, 16, 16, 16, 10, 16, 10, 16, 16, 16, 16)
    
    
    ' значения по умолчанию для наклеек
    For group = 0 To GROUPS_COUNT - 1
     Stickers(group).Count = 0                      ' счётчик элементов
     Stickers(group).FontSize = FontArray(group)    ' размер шрифта в pt
     Stickers(group).width = WidthArray(group)      ' ширина наклейки в мм
     Stickers(group).Height = HeightArray(group)    ' высота наклейки в мм
     Stickers(group).Type = GroupsArray(group)      ' shapetype
     Stickers(group).Prefix = GroupsArray(group)    ' символы перед shapenum на наклейке
    Next group
           
    ' исключения
    Stickers(6).Prefix = "#A"

  
  ' НАДПИСИ ПОД ДЕРЖАТЕЛЬ
  
    w_HL = 27 / 25.4
    h_HL = 19 / 25.4
    
    w_KL = 5.1 / 25.4
    h_KL = 10 / 25.4
    
    X0 = w_HL / 2 + 0.2
    
    X = X0
    Y = h_HL / 2 + 0.2
   
    For Each vPage In ActiveDocument.Pages ' перебираем лампочки
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "HL" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    pCaption = vShape.CellsU("Prop.Caption").ResultStr("")
                    Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("Надписи под держатель маркировки.vss").Masters.ItemU("HL"), X, Y)
                    NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                    NewNaklieka.Cells("Prop.Caption").Formula = Chr(34) + pCaption + Chr(34)
                    X = X + w_HL
                    If X > 190 / 25.4 Then
                        X = X0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    If X > X0 Then Y = Y + h_HL
    X = X0
    
 
    For Each vPage In ActiveDocument.Pages ' перебираем переключатели
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SA" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    pCaptionMain = vShape.CellsU("Prop.CaptionMain").ResultStr("")
                    'If vShape.CellsU("User.StateNum").ResultStr("") = "2" Then ' если двухполюсный переключатель
                    
                    If vShape.CellsU("User.StateNum") = 2 Then ' если двухполюсный переключатель
                        pCaption1 = vShape.CellsU("Prop.Caption1").ResultStr("")
                        pCaption2 = vShape.CellsU("Prop.Caption2").ResultStr("")
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("Надписи под держатель маркировки.vss").Masters.ItemU("SA2P"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                        NewNaklieka.Cells("Prop.CaptionMain").Formula = Chr(34) + pCaptionMain + Chr(34)
                        NewNaklieka.Cells("Prop.Caption1").Formula = Chr(34) + pCaption1 + Chr(34)
                        NewNaklieka.Cells("Prop.Caption2").Formula = Chr(34) + pCaption2 + Chr(34)
                    End If
                    
                    If vShape.CellsU("User.StateNum") = 3 Then ' если трёхполюсный переключатель
                        pCaption1 = vShape.CellsU("Prop.Caption1").ResultStr("")
                        pCaption2 = vShape.CellsU("Prop.Caption2").ResultStr("")
                        pCaption3 = vShape.CellsU("Prop.Caption3").ResultStr("")
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("Надписи под держатель маркировки.vss").Masters.ItemU("SA3P"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                        NewNaklieka.Cells("Prop.CaptionMain").Formula = Chr(34) + pCaptionMain + Chr(34)
                        NewNaklieka.Cells("Prop.Caption1").Formula = Chr(34) + pCaption1 + Chr(34)
                        NewNaklieka.Cells("Prop.Caption2").Formula = Chr(34) + pCaption2 + Chr(34)
                        NewNaklieka.Cells("Prop.Caption3").Formula = Chr(34) + pCaption3 + Chr(34)
                    End If
                    X = X + w_HL
                    If X > 190 / 25.4 Then
                        X = X0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    If X > X0 Then Y = Y + h_HL
    X = X0
    w_HL = 27 / 25.4
    h_HL = 19 / 25.4
    
    For Each vPage In ActiveDocument.Pages ' перебираем кнопки
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SB" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    pCaption = vShape.CellsU("Prop.Caption").ResultStr("")
                    Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("Надписи под держатель маркировки.vss").Masters.ItemU("SB"), X, Y)
                    NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                    NewNaklieka.Cells("Prop.Caption").Formula = Chr(34) + pCaption + Chr(34)
                    X = X + w_HL
                    If X > 190 / 25.4 Then
                        X = X0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    If X > X0 Then Y = Y + h_HL
    
    
   For Each vPage In ActiveDocument.Pages ' перебираем клеммы
       Set vShapes = vPage.Shapes
      For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "KL" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    pCaption = vShape.CellsU("Prop.Caption").ResultStr("")
                    Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("наклейки на элементы шкафа.vss").Masters.ItemU("KL"), X, Y)
                    NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                    NewNaklieka.Cells("Prop.Caption").Formula = Chr(34) + pCaption + Chr(34)
                    X = X + w_KL
                    If X > 190 / 25.4 Then
                        X = X0
                        Y = Y + h_KL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    If X > X0 Then Y = Y + h_KL


' НАКЛЕЙКИ НА ЭЛЕМЕНТЫ
    X = 0.2
    
     For Each vPage In ActiveDocument.Pages ' перебираем всё
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
               For group = 0 To GROUPS_COUNT - 1
                        
                If vShape.CellsU("User.ShapeType").ResultStr("") = Stickers(group).Type Then
                  Stickers(group).Count = Stickers(group).Count + 1
                  sNumbers(group, Stickers(group).Count - 1) = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                  For i = Stickers(group).Count - 2 To 0 Step -1
                   If (sNumbers(group, i) > sNumbers(group, i + 1)) Then
                    k = sNumbers(group, i)
                    sNumbers(group, i) = sNumbers(group, i + 1)
                    sNumbers(group, i + 1) = k
                   Else
                    Exit For
                   End If
                  Next i
                  
                  Exit For
                  
                End If
                                         
               Next group

            End If
        Next vShape
    Next vPage
    
    
    For group = 0 To GROUPS_COUNT - 1
     For i = 0 To Stickers(group).Count - 1
          Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("наклейки на элементы шкафа.vss").Masters.ItemU("Sticker"), X, Y)
          NewNaklieka.Cells("Char.Size").FormulaU = Stickers(group).FontSize & " pt"
          NewNaklieka.Text = Stickers(group).Prefix & sNumbers(group, i)
          NewNaklieka.Cells("Height").FormulaU = CStr(Stickers(group).Height) & " mm"
          NewNaklieka.Cells("Width").FormulaU = CStr(Stickers(group).width) & " mm"
          X = X + NewNaklieka.CellsU("Width")
          If (X + NewNaklieka.CellsU("Width")) > 8 Then ' 210/25.4 = 8.27
            X = 0.2
            Y = Y + Stickers(group).Height / 25.4
          End If
     Next i
     If (X > 0.21) Then Y = Y + Stickers(group).Height / 25.4
     X = 0.2
    Next group
          
   
End Sub







