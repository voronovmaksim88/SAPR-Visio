VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_CB 
   Caption         =   "UserForm1"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   OleObjectBlob   =   "Form_CB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_CB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub ComboBox_Model_Change()

End Sub

Private Sub CommandButton_Cancel_Click()
    Form_CB.Hide
End Sub

Private Sub CommandButton_OK_Click()
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Model").FormulaU = Chr(34) + ComboBox_Model.Text + Chr(34)
    Form_CB.Hide
End Sub


Private Sub Label2_Click()

End Sub

Private Sub UserForm_Activate()
    ComboBox_Model.List = Array("�����(�)-LS 5�1.5-0.660", "�����(�)-LS 5�2.5-0.660", "�����(�)-LS 5�4-0.660", "�����(�)-LS 5�6-0.660", "�����(�)-LS 5�10-0.660", "�����(�)-LS 4�1.5-0.660", "�����(�)-LS 4�2.5-0.660", "�����(�)-LS 3�1.5-0.660", "�����(�)-LS 3�2.5-0.660", "��� 2�0,75", "��� 2�1,5", "��� 2�2,5", "��� 3�0,75", "��� 3�1,5", "��� 3�2,5", "��� 4�0,75", "��� 4�1,5", "��� 5�0,75", "���� 2x0,75", "���� 3x0,75", "���� 4x0,75", "FTP4-ST (01-0145), ����� ����, 4 ���� Cat5e, 24AWG ������������ ��������������")
    ComboBox_Model.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Model").ResultStr("")
End Sub
