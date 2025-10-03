VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_KM 
   Caption         =   "UserForm1"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   OleObjectBlob   =   "Form_KM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_KM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' 2017-06-21 Воронов МВ
' перелопатил пол макросса


Private Sub Model()
    If ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = (Chr(34) + "Chint" + Chr(34)) Then
        TextBox_Model.Text = "NXC-" + ComboBox_Current.Text
    End If
    
    If ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = (Chr(34) + "Chint" + Chr(34)) And _
    ActiveWindow.Selection.PrimaryItem.Cells("prop.PolusNum").FormulaU = (Chr(34) + "2" + Chr(34)) Then
        TextBox_Model.Text = "NCH8-20"
    End If
    
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").FormulaU = Chr(34) + TextBox_Model.Text + Chr(34)
End Sub


Private Sub ComboBox_Current_Change()
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Current").FormulaU = Chr(34) + ComboBox_Current.Text + Chr(34)
    Model ' вызываем функцию расчёта имени пмодели
End Sub

Private Sub ComboBox_PolusNum_Change()
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.PolusNum").FormulaU = Chr(34) + ComboBox_PolusNum.Text + Chr(34)
    Model
End Sub

Private Sub CommandButton1_Click()
   
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + ComboBox1.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Current").FormulaU = Chr(34) + ComboBox_Current.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.PolusNum").FormulaU = Chr(34) + ComboBox_PolusNum.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").FormulaU = Chr(34) + TextBox_Model.Text + Chr(34)
    
    Form_KM.Hide
    
    
End Sub


Private Sub CommandButton2_Click()
    Form_KM.Hide
End Sub


Private Sub UserForm_Activate() ' выполняется при активации формы
    ComboBox1.List = Array("Chint", "Iek", "Schneider Electric", "LS", "Dekraft")
    ComboBox_Current.List = Array("6", "9", "12", "18", "25", "32", "40", "50", "65")
    ComboBox_PolusNum.List = Array("2", "3")
    
    ComboBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").ResultStr("")
    ComboBox2.Text = Round(ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeNum").ResultStr(""))
    ComboBox_Current.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Current").ResultStr("")
    ComboBox_PolusNum.Text = Round(ActiveWindow.Selection.PrimaryItem.Cells("Prop.PolusNum").ResultStr(""))
    TextBox_Model.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").ResultStr("")
End Sub

