VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_KL 
   Caption         =   "Êכוללא"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6510
   OleObjectBlob   =   "Form_KL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_KL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption").FormulaU = Chr(34) + ComboBox5.Text + Chr(34)
    Form_KL.Hide
End Sub

Private Sub CommandButton2_Click()
    Form_KL.Hide
End Sub

Private Sub UserForm_Activate()
    ComboBox5.List = Array("-", "+", "L1", "L2", "L3", "A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3")
    ComboBox5.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption").ResultStr("")
End Sub


