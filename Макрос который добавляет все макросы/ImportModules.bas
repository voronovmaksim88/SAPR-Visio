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

Sub ImportAllModules()

Const PATH As String = "D:\MyProgGit\SAPR-Visio\"
Const FILES_NUM As Integer = 30

Dim File(1 To FILES_NUM) As String
Dim �������������� As Boolean
Dim ������������������������� As Integer

�������������� = True
������������������������� = 0

File(1) = "AutoEskiz.bas"
File(2) = "AutoNakleiki_v2.bas"
File(3) = "AutoNum_v2.bas"
File(4) = "AutoPageNum.bas"
File(5) = "AutoPictureSize.bas"
File(6) = "Autospec_v8.bas"
File(7) = "Clear_Prop.bas"
File(8) = "FormShow.bas"
File(9) = "Form_E.frm"
File(10) = "Form_HL_v3.frm"
File(11) = "Form_QF_v1.frm"
File(12) = "Form_Spec.frm"
File(13) = "EditAllPages.bas"
File(14) = "Form_EditAllPages.frm"
File(15) = "Form_KM.frm"
File(16) = "Form_Area.frm"
File(17) = "Form_CB.frm"
File(18) = "Form_Box_Postgre_v2r0.frm"
File(19) = "Form_M.frm"
File(20) = "Form_SA.frm"
File(21) = "Form_All_Macros.frm"
File(22) = "PostgeSQL_0_test_connection.bas"
File(23) = "PostgeSQL_ControlCabinets.bas"
File(24) = "PostgeSQL_HWD.bas"
File(25) = "PostgeSQL_IP.bas"
File(26) = "PostgeSQL_Manufacturers.bas"
File(27) = "PostgeSQL_Material.bas"
File(28) = "PostgeSQL_GlobalTypes.bas"
File(29) = "Form_Sensors_PostgreSQL.frm"
File(30) = "PostgeSQL_Sensors.bas"


Dim oDoc As Visio.Document
Dim VBProj As VBIDE.VBProject
Dim cmpComponents As VBIDE.VBComponents
Dim cmpComp As VBIDE.VBComponent
Dim i As Integer

Set oDoc = ActiveDocument
Set VBProj = ActiveDocument.VBProject
Set cmpComponents = VBProj.VBComponents

' ��������� ������� ���� ������ ����� ������� �������
For i = 1 To FILES_NUM
    If Dir(PATH & File(i)) = "" Then
        MsgBox "����������� ���� " & File(i), vbExclamation, "������ �������"
        �������������� = False
        Exit Sub
    End If
Next i

' ������� ������������ ����������, ������� ��������� � ������ ������� �������
On Error Resume Next
For Each cmpComp In VBProj.VBComponents
    For i = 1 To FILES_NUM
        If (cmpComp.Name = Left(File(i), Len(File(i)) - 4)) Then
            'MsgBox ("������� " & cmpComp.Name)
            cmpComponents.Remove cmpComponents.item(cmpComp.Name)
            Exit For
        End If
    Next i
Next
On Error GoTo 0

' ����������� ��� ����������
On Error Resume Next
For i = 1 To FILES_NUM
    cmpComponents.Import (PATH & File(i))
    If Err.Number = 0 Then
        ������������������������� = ������������������������� + 1
    Else
        �������������� = False
        MsgBox "������ ��� ������� ����� " & File(i) & vbCrLf & _
               "������: " & Err.Description, vbExclamation, "������ �������"
        Err.Clear
    End If
Next i
On Error GoTo 0

' ���������� ��������� �� �������� �������, ���� ��� ������ ��� ������
If �������������� And ������������������������� = FILES_NUM Then
    MsgBox "��� ������ ������� �������������!" & vbCrLf & _
           "������������� ������: " & ������������������������� & " �� " & FILES_NUM, _
           vbInformation, "������ ��������"
ElseIf ������������������������� > 0 Then
    MsgBox "������ �������� � ����������������." & vbCrLf & _
           "������������� ������: " & ������������������������� & " �� " & FILES_NUM, _
           vbExclamation, "������ ��������"
End If

End Sub
