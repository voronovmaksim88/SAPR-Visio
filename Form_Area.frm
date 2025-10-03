VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Area 
   Caption         =   "Форма сечения проводника"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3390
   OleObjectBlob   =   "Form_Area.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Area"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ComboBox_Area_Change()

End Sub

Private Sub CommandButton1_Click()
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Area").FormulaU = Chr(34) + ComboBox_Area.Text + Chr(34)
    Form_Area.Hide
End Sub

Private Sub CommandButton2_Click()
    Form_Area.Hide
End Sub

Private Sub UserForm_Activate() ' выполняется при активации формы
    ComboBox_Area.List = Array("1.5", "2.5", "4", "6", "10", "16", "25", "35", "70", "120")
    ComboBox_Area.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Area").ResultStr("")
        If Not Application.ActiveWindow.Selection.Count = 0 Then
        Label4.Caption = Application.ActiveWindow.Selection(1).Name
    End If
End Sub
