Attribute VB_Name = "AutoEskiz"
' 2017-06-26 Воронов М.В.
' поправил добавление контакторов
' поправил добавление реле

' 2017-08-03 Воронов М.В.
' поправил добавление кнопки

' 2017-10-24 Воронов М.В.
' поправил добавление кнопки

' 2018-05-16 Воронов М.В.
' Мега инсайт !!! надо просто задавать размеры в самих шэйпах !!!


Sub AutoEskiz()
    Dim vShape As Visio.Shape
    Dim vShapes As Visio.Shapes
    Dim vPage As Visio.Page
    Dim NewShape As Visio.Shape
    Dim pNumber As Integer
    Dim pColor As Integer
    Dim pColorCaption As String
    Dim pCaption As String
    Dim pCaptionMain As String
    Dim pCaption1 As String
    Dim pCaption2 As String
    Dim pCaption3 As String
    Dim X As Double
    Dim Y As Double
    Dim w_HL As Double ' ширина лампочки
    Dim w_SA As Double ' ширина переключателя
    Dim w_K2p As Double ' ширина реле 2п
    Dim w_K4p As Double ' ширина реле 4п
    Dim w_KM2p As Double ' ширина контактора 2п
    Dim w_KM3p As Double ' ширина контактора 3п
    
    w_HL = 7 / 25.4
    w_SA = 7 / 25.4
    w_K2p = 5 / 25.4
    w_K4p = 7.5 / 25.4
    w_KM2p = 5 / 25.4
    w_KM3p = 12.5 / 25.4
    X = 0
    Y = 0
    For Each vPage In ActiveDocument.Pages ' перебераем лампочки
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "HL" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    pColor = Round(vShape.CellsU("Prop.Color").ResultStr(""))
                    pColorCaption = vShape.CellsU("User.ColorCaption").ResultStr("")
                    pCaption = vShape.CellsU("Prop.Caption").ResultStr("")
                    Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("HL_v1"), X, Y)
                    NewShape.Cells("User.ShapeNum").Formula = pNumber
                    NewShape.Cells("Prop.Color").Formula = pColor
                    NewShape.Cells("Prop.ColorCaption").Formula = Chr(34) + pColorCaption + Chr(34)
                    NewShape.Cells("Prop.Caption").Formula = Chr(34) + pCaption + Chr(34)
                    X = X + w_HL * 2
                End If
            End If
        Next vShape
    Next vPage
    
    X = 0
    Y = -2
    For Each vPage In ActiveDocument.Pages ' перебераем переключатели
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SA" Then
                    If Round(vShape.CellsU("User.StateNum").ResultStr("")) = 2 Then
                        pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                        pCaptionMain = vShape.CellsU("Prop.CaptionMain").ResultStr("")
                        pCaption1 = vShape.CellsU("Prop.Caption1").ResultStr("")
                        pCaption2 = vShape.CellsU("Prop.Caption2").ResultStr("")
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("SA_2П_v1"), X, Y)
                        NewShape.Cells("User.ShapeNum").Formula = pNumber
                        NewShape.Cells("Prop.CaptionMain").Formula = Chr(34) + pCaptionMain + Chr(34)
                        NewShape.Cells("Prop.Caption1").Formula = Chr(34) + pCaption1 + Chr(34)
                        NewShape.Cells("Prop.Caption2").Formula = Chr(34) + pCaption2 + Chr(34)
                        X = X + w_SA * 2
                    End If
                
                    If Round(vShape.CellsU("User.StateNum").ResultStr("")) = 3 Then
                        pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                        pCaptionMain = vShape.CellsU("Prop.CaptionMain").ResultStr("")
                        pCaption1 = vShape.CellsU("Prop.Caption1").ResultStr("")
                        pCaption2 = vShape.CellsU("Prop.Caption2").ResultStr("")
                        pCaption3 = vShape.CellsU("Prop.Caption3").ResultStr("")
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("SA_3П_v1"), X, Y)
                        NewShape.Cells("User.ShapeNum").Formula = pNumber
                        NewShape.Cells("Prop.CaptionMain").Formula = Chr(34) + pCaptionMain + Chr(34)
                        NewShape.Cells("Prop.Caption1").Formula = Chr(34) + pCaption1 + Chr(34)
                        NewShape.Cells("Prop.Caption2").Formula = Chr(34) + pCaption2 + Chr(34)
                        NewShape.Cells("Prop.Caption3").Formula = Chr(34) + pCaption3 + Chr(34)
                        X = X + w_SA * 2
                    End If
                End If
            End If
        Next vShape
    Next vPage
    
    
    For Each vPage In ActiveDocument.Pages ' перебераем кнопки
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SB" Then
                        pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                        pCaption = vShape.CellsU("Prop.Caption").ResultStr("")
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("SB"), X, Y)
                        NewShape.Cells("User.ShapeNum").Formula = pNumber
                        NewShape.Cells("Prop.Caption").Formula = Chr(34) + pCaption + Chr(34)
                        X = X + w_SA * 2
                End If
            End If
        Next vShape
    Next vPage
   
    
    X = 0
    Y = -3
        For Each vPage In ActiveDocument.Pages ' перебераем автоматы
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "QF" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "1" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_1p"), X, Y)
                        X = X + w_K2p * 1
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "2" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_2p"), X, Y)
                        X = X + w_K2p * 2
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "3" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_3p"), X, Y)
                        X = X + w_K2p * 3
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "4" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_4p"), X, Y)
                        X = X + w_K2p * 4
                    End If
                    NewShape.Cells("User.ShapeNum").Formula = pNumber
                    
                End If
            End If
        Next vShape
    Next vPage
    
    
    X = 0
    Y = -4
        For Each vPage In ActiveDocument.Pages ' перебераем рубильники НО КРИВО !!!
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "QS" Then
                 'MsgBox ("hello" + str(X))
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "1" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_1p"), X, Y)
                        X = X + w_K2p * 1
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "2" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_2p"), X, Y)
                        X = X + w_K2p * 2
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "3" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_3p"), X, Y)
                        X = X + w_K2p * 3
                        'MsgBox ("hello1111" + str(X))
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "4" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_4p"), X, Y)
                        X = X + w_K2p * 4
                    End If
                    NewShape.Cells("User.ShapeNum").Formula = pNumber
                    
                End If
            End If
        Next vShape
    Next vPage
    
    
    X = 0
    Y = -5
        For Each vPage In ActiveDocument.Pages ' перебераем реле
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "K" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    If vShape.CellsU("Prop.PolusNum").ResultStr("") = "2" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("K_2p"), X, Y)
                        X = X + w_K2p
                    End If
                    
                    If vShape.CellsU("Prop.PolusNum").ResultStr("") = "4" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("K_4p"), X, Y)
                        X = X + w_K4p
                    End If
                    NewShape.Cells("User.ShapeNum").Formula = pNumber
                    
                End If
            End If
        Next vShape
    Next vPage
    
    X = 0
    Y = -6
    For Each vPage In ActiveDocument.Pages ' перебераем контакторы
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "KM" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    If vShape.CellsU("Prop.PolusNum").ResultStr("") = "2" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("KM_2p"), X, Y)
                        X = X + w_KM2p
                    End If
                    
                    If vShape.CellsU("Prop.PolusNum").ResultStr("") = "3" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("KM_3p"), X, Y)
                        X = X + w_KM3p
                        'MsgBox ("hello" + Str(X))
                    End If
                    NewShape.Cells("User.SecondaryShapeNum").Formula = pNumber
                    
                End If
            End If
        Next vShape
    Next vPage
    
    
    
    
    
    
     For Each vPage In ActiveDocument.Pages ' перебераем ВСЁ !!!
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "QF" Or vShape.CellsU("User.ShapeType").ResultStr("") = "QS" Then
                 MsgBox ("hello" + str(X))
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "1" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_1p"), X, Y)
                        X = X + w_K2p * 1
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "2" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_2p"), X, Y)
                        X = X + w_K2p * 2
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "3" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_3p"), X, Y)
                        X = X + w_K2p * 3
                        'MsgBox ("hello1111" + str(X))
                    End If
                    
                    If vShape.CellsU("User.PolusNum").ResultStr("") = "4" Then
                        Set NewShape = Application.ActiveWindow.Page.Drop(Application.Documents.Item("10_эскиз шкафа_1к4.vss").Masters.ItemU("QF_4p"), X, Y)
                        X = X + w_K2p * 4
                    End If
                    NewShape.Cells("User.ShapeNum").Formula = pNumber
                    
                End If
            End If
        Next vShape
    Next vPage
    
    
End Sub














