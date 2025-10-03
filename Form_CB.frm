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
    ComboBox_Model.List = Array("¬¬√нг(ј)-LS 5х1.5-0.660", "¬¬√нг(ј)-LS 5х2.5-0.660", "¬¬√нг(ј)-LS 5х4-0.660", "¬¬√нг(ј)-LS 5х6-0.660", "¬¬√нг(ј)-LS 5х10-0.660", "¬¬√нг(ј)-LS 4х1.5-0.660", "¬¬√нг(ј)-LS 4х2.5-0.660", "¬¬√нг(ј)-LS 3х1.5-0.660", "¬¬√нг(ј)-LS 3х2.5-0.660", "ѕ¬— 2х0,75", "ѕ¬— 2х1,5", "ѕ¬— 2х2,5", "ѕ¬— 3х0,75", "ѕ¬— 3х1,5", "ѕ¬— 3х2,5", "ѕ¬— 4х0,75", "ѕ¬— 4х1,5", "ѕ¬— 5х0,75", "ћ ЁЎ 2x0,75", "ћ ЁЎ 3x0,75", "ћ ЁЎ 4x0,75", "FTP4-ST (01-0145), ¬ита€ пара, 4 пары Cat5e, 24AWG многожильные экранированные")
    ComboBox_Model.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Model").ResultStr("")
End Sub
