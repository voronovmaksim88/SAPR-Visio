VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_HL_v3 
   Caption         =   "Form_HL"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5895
   OleObjectBlob   =   "Form_HL_v3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_HL_v3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function CalculateModel(pShape As Visio.Shape) As String

CalculateModel = "Something of " + pShape.Cells("prop.Manufacturer").ResultStr("")
    If pShape.Cells("prop.Manufacturer").ResultStr("") = "Chint" Then
        
        If ActiveWindow.Selection.PrimaryItem.Cells("prop.Up") = "24" Then
            CalculateModel = "ND16-22DS/2 "
        End If
        
        If ActiveWindow.Selection.PrimaryItem.Cells("prop.Up") = "220" Then
            CalculateModel = "ND16-22D/2 "
        End If

        Select Case ComboBox2.Text
            Case "Белолунный"
                ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 1
                CalculateModel = CalculateModel + "(W)"
            Case "Красный"
                ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 2
                CalculateModel = CalculateModel + "(R)"
            Case "Зелёный"
                ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 3
                CalculateModel = CalculateModel + "(G)"
            Case "Синий"
                ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 4
                CalculateModel = CalculateModel + "(B)"
            Case "Жёлтый"
                ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 5
                CalculateModel = CalculateModel + "(Y)"
            Case Else
                ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 0
                CalculateModel = CalculateModel + "(?)"
        End Select
        
        If ActiveWindow.Selection.PrimaryItem.Cells("prop.Up") = "24" Then
            CalculateModel = CalculateModel + " AC\DC24В"
        End If
        
        If ActiveWindow.Selection.PrimaryItem.Cells("prop.Up") = "220" Then
            CalculateModel = CalculateModel + " AC230В"
        End If
        
    End If
End Function
Private Sub CommandButton1_Click()
   
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + ComboBox1.Text + Chr(34)
  
    Select Case ComboBox2.Text
        Case "Белолунный"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 1
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Бел.)" + Chr(34)
        Case "Красный"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 2
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Красн.)" + Chr(34)
        Case "Зелёный"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 3
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Зел.)" + Chr(34)
        Case "Синий"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 4
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Син.)" + Chr(34)
        Case "Жёлтый"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 5
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Жёл.)" + Chr(34)
        Case Else
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 0
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "" + Chr(34)
    End Select
    
    
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Up").FormulaU = Chr(34) + ComboBox4.Text + Chr(34)
    
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption").FormulaU = Chr(34) + ComboBox5.Text + Chr(34)
    
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Model").FormulaU = Chr(34) + TextBox1.Text + Chr(34)
    
    Form_HL_v3.Hide
End Sub

Private Sub CommandButton2_Click()
    Form_HL_v3.Hide
End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + ComboBox1.Text + Chr(34)
  
    Select Case ComboBox2.Text
        Case "Белолунный"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 1
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Бел.)" + Chr(34)
        Case "Красный"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 2
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Красн.)" + Chr(34)
        Case "Зелёный"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 3
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Зел.)" + Chr(34)
        Case "Синий"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 4
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Син.)" + Chr(34)
        Case "Жёлтый"
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 5
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Жёл.)" + Chr(34)
        Case Else
            ActiveWindow.Selection.PrimaryItem.Cells("prop.Color") = 0
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "" + Chr(34)
    End Select
    
    
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Up").FormulaU = Chr(34) + ComboBox4.Text + Chr(34)
    
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption").FormulaU = Chr(34) + ComboBox5.Text + Chr(34)
    
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Model").FormulaU = Chr(34) + TextBox1.Text + Chr(34)
    TextBox1.Text = CalculateModel(ActiveWindow.Selection.PrimaryItem)
End Sub

Private Sub UserForm_Activate()
    
    ComboBox1.List = Array("Chint", "Iek", "Schneider Electric", "Tecfor")
    ComboBox2.List = Array("Белолунный", "Красный", "Зелёный", "Синий", "Жёлтый")
    ComboBox3.List = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
    ComboBox4.List = Array("12", "24", "220")
    ComboBox5.List = Array("Питание 24В", "Питание 220", "Сеть", "Работа", "Авария", "Управление")
    
    ComboBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").ResultStr("")
    
    Select Case ActiveWindow.Selection.PrimaryItem.Cells("prop.Color")
        Case 1
           ComboBox2.Text = "Белолунный"
           ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Бел.)" + Chr(34)
        Case 2
            ComboBox2.Text = "Красный"
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Красн.)" + Chr(34)
        Case 3
            ComboBox2.Text = "Зелёный"
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Зел.)" + Chr(34)
        Case 4
            ComboBox2.Text = "Синий"
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Син.)" + Chr(34)
        Case 5
            ComboBox2.Text = "Жёлтый"
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "(Жёл.)" + Chr(34)
        Case Else
            ComboBox2.Text = "ХЗ"
            ActiveWindow.Selection.PrimaryItem.Cells("User.ColorCaption").FormulaU = Chr(34) + "" + Chr(34)
    End Select
    
    ComboBox3.Text = Round(ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeNum").ResultStr(""), 0)
    ComboBox4.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Up").ResultStr("")
    ComboBox5.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Caption").ResultStr("")
    If Len(ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").Formula) > 0 Then
       TextBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").ResultStr("")
    Else
       TextBox1.Text = CalculateModel(ActiveWindow.Selection.PrimaryItem)
    End If

End Sub


