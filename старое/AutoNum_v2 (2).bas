Attribute VB_Name = "AutoNum_v2"
' 2017-06-15 ������� ��
' ������� ���������� cXT
' �������� ���������� ����������
' ������� ������������ ������

' 2017-06-21 ������� ��
' ������� ������������� ������ �������

' 2017-06-25 ������� ��
' ������� ������������� �����������

' 2017-07-14 ������� ��
' ������� ������������� LineNum

' 2017-07-18 ������
' kabel -> CB
' LineNum -> LN

' 2017-07-24 ������� ��
' ������� SF

' 2017-08-10 ������� ��
' ������� TV

' 2017-08-30 ������� ��
' ������� THI

' 2017-09-06 ������� ��
' ������� QS

' 2017-10-23 ������� ��
' ������� THE
' ������� QI

' 2017-12-07 ������� ��
' ������� QFD

' 2018-01-10 ��������� ��
' ������� ��-���� SSR

' 2018-02-05 ������� ��
' ������� LS

' 2018-02-16 ������� ��
' ������� KK

' 2018-03-07 ������� ��
' ������� A ��� ������ ������� � �������� ���

' 2018-03-07 ������� ��
' ������� KL ��� ������


' 2018-04-29 ������� ��
' �������� �������
'Dim Arr(255) As ShapeRec
'Dim Pars(255) As ParentRec


' 2018-06-05 ������� ��
' ������� "�" ��� �����������������

' 2018-06-25 ������� ��
' �������� "KK"
' ������� "F" ��������������

' 2018-08-06 ������� ��
' ������� "HA" �������� ����������

' 2018-10-28 ������� ��
' ������� "VD" ����

' 2019-07-18 ������� ��
' "NumABC" ��� 3 ������ ���

' 2019-11-22 ������� ��
' hashid ������ ��� ������ �� ���� ����������� � ���� ������� ����,
' ��� ���� ���� �� ����� ����� ������� ��������� ���


' 2019-11-28 ������
' ����������� ������� ������ prop �� ������������� ����� � ��������


' 2020-01-09 ������� ��
' Arr � Pars ��������� �� 1024

' 2020-01-14 ������� ��
' FC �������� �����

' 2020-01-27 ������� ��
' TE �������� �����

' 2020-02-03 ������� ��
' Num_ABC �������� �����

' 2020-02-22 Serega
' Ungroup in DropK + Links

' 2020-04-21 ������� ��
' PS ������ �������� ����� HashID
' TS ������ �������� ����� HashID
' PE ������ �������� ����� HashID
' PDS ������ �������� ����� HashID

' 2020-04-24 Serega
' Able to change text field in Links

' 2020-12-21 ������� ��
' ������� KV

' 2023-08-29 ������� ��
' �������� PDE




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
' ��������� ������: Ctrl+w
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
'����� ����������� ��� ������ ����'
'''''''''''''''''''''''''''''''''''
Dim cSA As Integer, cSB As Integer, cQF As Integer, cKM As Integer
Dim cHL As Integer, cK As Integer, cXT As Integer, cCB As Integer
Dim cTS As Integer, cPS As Integer, cPDS As Integer, cHS As Integer
Dim cPE As Integer, cPDE As Integer, cHE As Integer, cTE As Integer
Dim cM As Integer, cFC As Integer, cLN As Integer, �UG As Integer
Dim cSF As Integer, cTV As Integer, �THI As Integer, �QS As Integer
Dim cTHE As Integer, cQI As Integer, cQFD As Integer, cSSR As Integer
Dim cLS As Integer, cKK As Integer, cKT As Integer, cA As Integer
Dim cKL As Integer, cT As Integer, cF As Integer, cHA As Integer
Dim cVD As Integer, �NumABC As Integer, �Num As Integer, cLink As Integer
Dim cKV As Integer
Count = 0: nPage = 0

Dim vPage As Visio.Page
'Set vPage = Application.ActivePage
Dim vShape As Visio.Shape
Dim vShapes As Visio.Shapes
Dim CurPage As Integer
Dim LinkCrt As Boolean

If cSort = "" Then cSort = "1"
cSort = InputBox("������� ����� �������� ��� �������� (����. 1-3)", "����������", cSort)
If InStr(cSort, "-") > 0 Then
  SortA = CInt(Left(cSort, InStr(cSort, "-") - 1))
  SortB = CInt(Right(cSort, Len(cSort) - InStr(cSort, "-")))
 Else
  SortA = Val(cSort)
  SortB = SortA
End If
If (SortA < 1) Or (SortB < 1) Then Exit Sub
SortType = MsgBox("������������� ������� �� ���������?", vbOKCancel, "����������")

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
    '����� ���� ����������� ��� ����'
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
        �UG = �UG + 1
        vShape.Cells("User.ShapeNum").Formula = �UG
                
    Case """FC"""
        �FC = �FC + 1
        vShape.Cells("User.ShapeNum").Formula = �FC
        parsCount = parsCount + 1
        Pars(parsCount).fNum = �FC
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
        �SF = �SF + 1
        vShape.Cells("User.ShapeNum").Formula = �SF
        
    Case """TV"""
        �TV = �TV + 1
        vShape.Cells("User.ShapeNum").Formula = �TV
        
    Case """THI"""
        �THI = �THI + 1
        vShape.Cells("User.ShapeNum").Formula = �THI
                
    Case """QS"""
        �QS = �QS + 1
        vShape.Cells("User.ShapeNum").Formula = �QS
        
    Case """THE"""
        �THE = �THE + 1
        vShape.Cells("User.ShapeNum").Formula = �THE
    
    Case """QI"""
        �QI = �QI + 1
        vShape.Cells("User.ShapeNum").Formula = �QI
        
    Case """QFD"""
        �QFD = �QFD + 1
        vShape.Cells("User.ShapeNum").Formula = �QFD
     
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
        �Num = �Num + 1
        vShape.Cells("User.ShapeNum").Formula = �Num
        
    Case """NumABC"""
        �NumABC = �NumABC + 1
        vShape.Cells("User.ShapeNum").Formula = �NumABC
        parsCount = parsCount + 1
        Pars(parsCount).fNum = �NumABC
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
        vShape.Text = "���� " + CStr(Links(inx).fSecondPage) + ": "
        Links(inx).fSecondPage = -1
      Else
        vShape.Text = "���� " + CStr(Links(inx).fFirstPage) + ": "
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
' �� ����� ������� ������ �������� ��������
    Dim visSheet As Visio.Page
    Dim visShape As Visio.Shape
    
    ' ������� ���� ������ � ������� ���������
    For Each visSheet In ThisDocument.Pages
        ' ������� ���� ������ �� ������� �����
        For Each visShape In visSheet.Shapes
            ' ��� ��� ��� ������ � ������� ������
            If visShape.CellExistsU("User.HashID", visExistsAnywhere) And Not visShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) Then
            ' ��� ������� �������� ����� ��������
                Dim CyrrentHashID As Long
                CyrrentHashID = visShape.Cells("User.HashID").Result("")
                            
                If Not AvailabilityChild(CyrrentHashID) Then
                    MsgBox "������ �������� ��� ������"
                    ' ��������� ��������� "PinX" � "PinY" �����
                    Dim xCoord As Double
                    Dim yCoord As Double
                    Dim Shape_Oval As Visio.Shape
            
                    xCoord = visShape.Cells("PinX").ResultIU
                    yCoord = visShape.Cells("PinY").ResultIU
                    Radius = 0.4   ' ������ �����
                    ' �������� ������ ����� �� �����
                    Set Shape_Oval = visSheet.DrawOval(xCoord - Radius, yCoord - Radius, xCoord + Radius, yCoord + Radius)
                
                    ' ��������� �������� ���� �����, ��������, ���� ������� � �������
                    Shape_Oval.Cells("LineColor").FormulaU = "RGB(255, 0, 0)"  ' ����� ���� �������
                    Shape_Oval.Cells("LineWeight").FormulaU = "0.05" ' ������� �������
                    Shape_Oval.CellsU("FillPattern").FormulaU = 0 ' ������ �������
                    'Shape_Oval.Text = CyrrentHashID
                    'Shape_Oval.Text = "��������"
                End If
            End If
        Next visShape
    Next visSheet
End Sub

Function AvailabilityChild(ParentHashID As Long) As Boolean
' ��� ������� ��������� ������� ����� � �������� HashID
    AvailabilityChild = False
    Dim visSheet As Visio.Page
    Dim visShape As Visio.Shape
    
    ' ������� ���� ������ � ������� ���������
    For Each visSheet In ThisDocument.Pages
        ' ������� ���� ������ �� ������� �����
        For Each visShape In visSheet.Shapes
            ' ��� ��� ��� ������ � ������� ������
            If visShape.CellExistsU("User.HashID", visExistsAnywhere) And visShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) Then
            ' ��� ������� �������� ����� ����
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
' �� ����� ������� ������ �������� ����
    Dim visSheet As Visio.Page
    Dim visShape As Visio.Shape
    
    ' ������� ���� ������ � ������� ���������
    For Each visSheet In ThisDocument.Pages
        ' ������� ���� ������ �� ������� �����
        For Each visShape In visSheet.Shapes
            ' ��� ��� ��� ������ � ������� ������
            If visShape.CellExistsU("User.HashID", visExistsAnywhere) And visShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) Then
            ' ��� ������� �������� ����� �����
                Dim CyrrentHashID As Long
                CyrrentHashID = visShape.Cells("User.HashID").Result("")
                            
                If Not AvailabilityParent(CyrrentHashID) Then
                    MsgBox "������ ������ ��� ��������"
                    ' ��������� ��������� "PinX" � "PinY" �����
                    Dim xCoord As Double
                    Dim yCoord As Double
                    Dim Shape_Oval As Visio.Shape
            
                    xCoord = visShape.Cells("PinX").ResultIU
                    yCoord = visShape.Cells("PinY").ResultIU
                    Radius = 0.4   ' ������ �����
                    ' �������� ������ ����� �� �����
                    Set Shape_Oval = visSheet.DrawOval(xCoord - Radius, yCoord - Radius, xCoord + Radius, yCoord + Radius)
                
                    ' ��������� �������� ���� �����, ��������, ���� ������� � �������
                    Shape_Oval.Cells("LineColor").FormulaU = "RGB(255, 0, 0)"  ' ����� ���� �������
                    Shape_Oval.Cells("LineWeight").FormulaU = "0.05" ' ������� �������
                    Shape_Oval.CellsU("FillPattern").FormulaU = 0 ' ������ �������
                    'Shape_Oval.Text = CyrrentHashID
                    'Shape_Oval.Text = "��������"
                End If
            End If
        Next visShape
    Next visSheet
End Sub

Function AvailabilityParent(ParentHashID As Long) As Boolean
' ��� ������� ��������� ������� ��������� � �������� HashID
    AvailabilityParent = False
    Dim visSheet As Visio.Page
    Dim visShape As Visio.Shape
    
    ' ������� ���� ������ � ������� ���������
    For Each visSheet In ThisDocument.Pages
        ' ������� ���� ������ �� ������� �����
        For Each visShape In visSheet.Shapes
            ' ��� ��� ��� ������ � ������� ������
            If visShape.CellExistsU("User.HashID", visExistsAnywhere) And Not visShape.CellExistsU("User.ParentShapeType", visExistsAnywhere) Then
            ' ��� ������� �������� ����� ���������
                Dim CyrrentHashID As Double
                CyrrentHashID = visShape.Cells("User.HashID").Result("")
                If CyrrentHashID = ParentHashID Then
                    AvailabilityParent = True ' �������� ������
                End If
            End If
        Next visShape
    Next visSheet
End Function
