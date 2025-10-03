VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_QF_v1 
   Caption         =   "Form_QF"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   OleObjectBlob   =   "Form_QF_v1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_QF_v1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox3_Change()

End Sub

Private Sub CommandButton_OK_Click()
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + ComboBox1.Text + Chr(34)
    'ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeNum").FormulaU = Chr(34) + ComboBox2.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Current").FormulaU = Chr(34) + ComboBox3.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Characteristic").FormulaU = Chr(34) + ComboBox4.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("User.PolusNum").FormulaU = Chr(34) + ComboBox5.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Model").FormulaU = Chr(34) + TextBox1.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Nom_Otkl_Spos").FormulaU = Chr(34) + TextBox2.Text + Chr(34)
    
    Dim vShape As Visio.shape
    Dim vSelection As Visio.Selection
    Dim shapeCount As Integer
    Dim i As Integer
    Set vSelection = Visio.ActiveWindow.Selection
    
    For Each vShape In vSelection
      shapeCount = vShape.Shapes.Count
      For i = 1 To shapeCount
       If InStr(1, vShape.Shapes(i).Text, "QF", vbTextCompare) > 0 Then
        vShape.Shapes(i).Text = "QF" + ComboBox2.Text
       End If
       
       If InStr(1, vShape.Shapes(i).Text, "B", vbTextCompare) > 0 Or InStr(1, vShape.Shapes(i).Text, "C", vbTextCompare) > 0 Or InStr(1, vShape.Shapes(i).Text, "D", vbTextCompare) > 0 Then
        vShape.Shapes(i).Text = ComboBox4.Text + ComboBox3.Text
       End If
       
      Next i
    Next vShape
    
    Form_QF_v1.Hide
    
End Sub

Private Sub CommandButton2_Click()
    Form_QF_v1.Hide
End Sub



Private Sub UserForm_Activate() ' выполняется при активации формы
    ComboBox1.List = Array("Chint", "Iek", "Schneider Electric", "LS", "Dekraft")
    ComboBox2.List = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20")
    ComboBox3.List = Array("1", "2", "4", "6", "10", "16", "20", "25", "32", "40", "50", "63", "80", "100", "125")
    ComboBox4.List = Array("B", "C", "D")
    ComboBox5.List = Array("1", "2", "3", "4")
   
    ComboBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").ResultStr("")
    ComboBox2.Text = Round(ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeNum").ResultStr(""), 0)
    ComboBox3.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Current").ResultStr("")
    ComboBox4.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Characteristic").ResultStr("")
    ComboBox5.Text = Round(ActiveWindow.Selection.PrimaryItem.Cells("User.PolusNum").ResultStr(""), 0)
    TextBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").ResultStr("")
    TextBox2.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Nom_Otkl_Spos").ResultStr("")
    
    If Not Application.ActiveWindow.Selection.Count = 0 Then
        TextBox3.Text = Application.ActiveWindow.Selection(1).Name
    End If
End Sub
