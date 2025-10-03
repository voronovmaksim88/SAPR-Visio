Attribute VB_Name = "ImportModules"
' 2017-10-16 ������� ��
' ������� AutoNakleiki.bas

' 2017-12-12 ������� ��
' ������� AutoPageNum

' 2018-04-20 ������� ��
' ������� Form_Box

' 2018-06-06 ������� ��
' Form_KL.frm
' Form_M.frm

' 2018-06-20 ������� ��
' ������� AutoNakleiki
' ���������� ���� �� ���������� ���� �� ������ ������ 3 �������

' 2018-07-31 ������� ��
'������� Form_SA , ��� ����� ��� ��������������, � ��� �������� ���� �����

' 2025-02-03 ������� ��
' ������ ����� ��� ��������. ������ �� ������ ��� ��� �����.

Sub ImportAllModules()

Const PATH As String = "D:\SynologyDrive\work_main\��� ����\VB+Visio\Sapr\"
Const FILES_NUM As Integer = 20

Dim File(1 To FILES_NUM) As String

File(1) = "Autospec_v8.bas"
File(2) = "FormShow.bas"
File(3) = "AutoNum_v2.bas"
File(4) = "Form_E.frm"
File(5) = "Form_HL_v3.frm"
File(6) = "Form_QF_v1.frm"
File(7) = "Form_Spec.frm"
File(8) = "AutoNakleiki_v2.bas"
File(9) = "AutoEskiz.bas"
File(10) = "EditAllPages.bas"
File(11) = "Form_EditAllPages.frm"
File(12) = "AutoPictureSize.bas"
File(13) = "Form_KM.frm"
File(14) = "Form_Area.frm"
File(15) = "AutoPageNum.bas"
File(16) = "Form_CB.frm"
File(17) = "Form_Box.frm"
File(18) = "Form_SA.frm"
File(19) = "Form_All_Macros.frm"
File(20) = "Form_M.frm"

Dim oDoc As Visio.Document
Dim VBProj As VBIDE.VBProject
Dim cmpComponents As VBIDE.VBComponents
Dim cmpComp As VBIDE.VBComponent
Dim i As Integer

Set oDoc = ActiveDocument
Set VBProj = ActiveDocument.VBProject
Set cmpComponents = VBProj.VBComponents

For i = 1 To FILES_NUM
    If Dir(PATH & File(i)) = "" Then
        MsgBox ("����������� ���� " & File(i))
        Exit Sub
    End If
Next i

For Each cmpComp In VBProj.VBComponents
  For i = 1 To FILES_NUM
    If (cmpComp.Name = Left(File(i), Len(File(i)) - 4)) Then
        'MsgBox ("������� " & cmpComp.Name)
        cmpComponents.Remove cmpComponents.Item(cmpComp.Name)
        Exit For
    End If
 Next i
Next

For i = 1 To FILES_NUM
 cmpComponents.Import (PATH & File(i))
Next i

End Sub

