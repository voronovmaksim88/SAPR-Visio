Attribute VB_Name = "AutoNakleiki"
'2017-09-29 ������� ��
'������� SF

'2017-10-17 ������� ��
'������� SB

'2017-10-17 ������� ��
'������� KT

'2018-05-23 ������� ��
'������� SSR
'������� KM
'������� A

'2018-05-27 ������� ��
'������� KL

Sub AutoNakleiki()
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
    Dim Est_SA As Boolean
    Dim Est_HL As Boolean
    Dim Est_QF As Boolean
    Dim Est_SF As Boolean
    Dim Est_SB As Boolean
    Dim Est_SSR As Boolean
    Dim Est_KM As Boolean
    Dim Est_A As Boolean
    Dim Est_KL As Boolean
    
    Dim X As Double
    Dim Y As Double
    Dim w_HL As Double ' ������ ��������
    Dim h_HL As Double ' ������ ��������
    
    w_HL = 27 / 25.4
    h_HL = 19 / 25.4

    X = 0
    Y = 0
    
    For Each vPage In ActiveDocument.Pages ' ���������� ��������
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "HL" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    pCaption = vShape.CellsU("Prop.Caption").ResultStr("")
                    Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("������� ��� ��������� ����������.vss").Masters.ItemU("HL"), X, Y)
                    NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                    NewNaklieka.Cells("Prop.Caption").Formula = Chr(34) + pCaption + Chr(34)
                    X = X + w_HL
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    
    

    
    For Each vPage In ActiveDocument.Pages ' ���������� �������������
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SA" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    pCaptionMain = vShape.CellsU("Prop.CaptionMain").ResultStr("")
                    'If vShape.CellsU("User.StateNum").ResultStr("") = "2" Then ' ���� ������������ �������������
                    
                    If vShape.CellsU("User.StateNum") = 2 Then ' ���� ������������ �������������
                        pCaption1 = vShape.CellsU("Prop.Caption1").ResultStr("")
                        pCaption2 = vShape.CellsU("Prop.Caption2").ResultStr("")
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("������� ��� ��������� ����������.vss").Masters.ItemU("SA2P"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                        NewNaklieka.Cells("Prop.CaptionMain").Formula = Chr(34) + pCaptionMain + Chr(34)
                        NewNaklieka.Cells("Prop.Caption1").Formula = Chr(34) + pCaption1 + Chr(34)
                        NewNaklieka.Cells("Prop.Caption2").Formula = Chr(34) + pCaption2 + Chr(34)
                    End If
                    
                    If vShape.CellsU("User.StateNum") = 3 Then ' ���� ����������� �������������
                        pCaption1 = vShape.CellsU("Prop.Caption1").ResultStr("")
                        pCaption2 = vShape.CellsU("Prop.Caption2").ResultStr("")
                        pCaption3 = vShape.CellsU("Prop.Caption3").ResultStr("")
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("������� ��� ��������� ����������.vss").Masters.ItemU("SA3P"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                        NewNaklieka.Cells("Prop.CaptionMain").Formula = Chr(34) + pCaptionMain + Chr(34)
                        NewNaklieka.Cells("Prop.Caption1").Formula = Chr(34) + pCaption1 + Chr(34)
                        NewNaklieka.Cells("Prop.Caption2").Formula = Chr(34) + pCaption2 + Chr(34)
                        NewNaklieka.Cells("Prop.Caption3").Formula = Chr(34) + pCaption3 + Chr(34)
                    End If
                    X = X + w_HL
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    
    
    X = 0
    w_HL = 27 / 25.4
    h_HL = 19 / 25.4
    
    For Each vPage In ActiveDocument.Pages ' ���������� ������
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SB" Then
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    pCaption = vShape.CellsU("Prop.Caption").ResultStr("")
                    Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("������� ��� ��������� ����������.vss").Masters.ItemU("SB"), X, Y)
                    NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
                    NewNaklieka.Cells("Prop.Caption").Formula = Chr(34) + pCaption + Chr(34)
                    X = X + w_HL
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
   
 
    
    
    '������ �� ���������� ����� ������� �������� �� ���� ��������
    
    w_HL = 20 / 25.4
    h_HL = 15 / 25.4
    X = 0
    
     For Each vPage In ActiveDocument.Pages ' ���������� �������������
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SA" Then
                    If Est_SA = False Then
                        Y = Y + h_HL
                        Est_SA = True
                    End If

                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("SA"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + w_HL
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage

    
     For Each vPage In ActiveDocument.Pages ' ���������� ��������
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "HL" Then
                    If Est_HL = False And Est_SA = False Then
                        Y = Y + h_HL
                        Est_HL = True
                    End If
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("HL"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + w_HL
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
           
      
    For Each vPage In ActiveDocument.Pages ' ���������� ������
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SB" Then
                    If Est_SB = False Then
                        Y = Y + h_HL
                        Est_SB = True
                    End If
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("SB"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + w_HL
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
          
    w_HL = 18 / 25.4
    h_HL = 10 / 25.4
    For Each vPage In ActiveDocument.Pages ' ���������� ��������
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "QF" Then
                    If Est_QF = False Then
                        Y = Y + w_HL
                        Est_QF = True
                    End If
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("QF"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + 18 / 25.4
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
    
    For Each vPage In ActiveDocument.Pages ' ���������� ���
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SF" Then
                    If Est_SF = False Then
                        Y = Y + h_HL
                        Est_SF = True
                    End If
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("SF"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + 18 / 25.4
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
    
    For Each vPage In ActiveDocument.Pages ' ���������� ��
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "KT" Then
                    If Est_SF = False Then
                        Y = Y + h_HL
                        Est_SF = True
                    End If
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("KT"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + 18 / 25.4
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
    
    
    w_HL = 18 / 25.4
    h_HL = 10 / 25.4
    For Each vPage In ActiveDocument.Pages ' ���������� SSR
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "SSR" Then
                    If Est_SSR = False Then
                        Y = Y + h_HL
                        Est_SSR = True
                    End If
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("SSR"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + 18 / 25.4
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
    
    
    w_HL = 18 / 25.4
    h_HL = 6 / 25.4
    For Each vPage In ActiveDocument.Pages ' ���������� KM
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "KM" Then
                    If Est_KM = False Then
                        Y = Y + h_HL
                        Est_KM = True
                    End If
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("KM"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + 18 / 25.4
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
  
    w_HL = 18 / 25.4
    h_HL = 10 / 25.4
    For Each vPage In ActiveDocument.Pages ' ���������� A
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "A" Then
                    If Est_A = False Then
                        Y = Y + h_HL
                        Est_A = True
                    End If
                    pNumber = Round(vShape.CellsU("User.ShapeNum").ResultStr(""))
                    
                        Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("A"), X, Y)
                        NewNaklieka.Cells("User.ShapeNum").Formula = pNumber
 
                    X = X + 18 / 25.4
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
  
    w_HL = 5.5
    h_HL = 5
    For Each vPage In ActiveDocument.Pages ' ���������� KL
        Set vShapes = vPage.Shapes
        For Each vShape In vShapes
            If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
                If vShape.CellsU("User.ShapeType").ResultStr("") = "KL" Then
                    If Est_KL = False Then
                        Y = Y + h_HL / 25.4
                        Est_KL = True
                    End If
                    pCaption = vShape.CellsU("Prop.Caption").ResultStr("")
                    Set NewNaklieka = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������� �� �������� �����.vss").Masters.ItemU("KL"), X, Y)
                    NewNaklieka.Cells("Prop.Caption").Formula = Chr(34) + pCaption + Chr(34)
 
                    X = X + w_HL / 25.4
                    If X > 180 / 25.4 Then
                        X = 0
                        Y = Y + h_HL / 25.4
                    End If
                End If
            End If
        Next vShape
    Next vPage
    X = 0
   
End Sub







