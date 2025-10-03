VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_SA 
   Caption         =   "UserForm1"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5700
   OleObjectBlob   =   "Form_SA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_SA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
   
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + ComboBox_Manufacturer.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Model").FormulaU = Chr(34) + ComboBox_Model.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.CaptionMain").FormulaU = Chr(34) + ComboBox_CaptionMain.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption1").FormulaU = Chr(34) + ComboBox_Caption1.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption2").FormulaU = Chr(34) + ComboBox_Caption2.Text + Chr(34)
    Form_SA.Hide
End Sub

Private Sub CommandButton2_Click()
    Form_SA.Hide
End Sub




Private Sub Label8_Click()

End Sub

Private Sub UserForm_Activate()

    'ComboBox5.List = Array("Питание 24В", "Питание 220", "Сеть", "Работа", "Авария", "Управление")
    ComboBox_ShapeNum.Text = Round(ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeNum").ResultStr(""), 0)
    ComboBox_Manufacturer.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").ResultStr("")
    ComboBox_Model.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Model").ResultStr("")
    ComboBox_CaptionMain = ActiveWindow.Selection.PrimaryItem.Cells("prop.CaptionMain").ResultStr("")
    ComboBox_Caption1 = ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption1").ResultStr("")
    ComboBox_Caption2 = ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption2").ResultStr("")
    ComboBox_PolusNum = Round(ActiveWindow.Selection.PrimaryItem.Cells("User.PolusNum").ResultStr(""), 0)
End Sub



