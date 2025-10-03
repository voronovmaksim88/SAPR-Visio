VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_M 
   Caption         =   "Привод"
   ClientHeight    =   5028
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6132
   OleObjectBlob   =   "Form_M.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Name").FormulaU = Chr(34) + TextBox_Name.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Note").FormulaU = Chr(34) + ComboBox6.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + TextBox_Manufacturer.Text + Chr(34)
    Form_M.Hide
End Sub

Private Sub CommandButton2_Click()
    Form_M.Hide
End Sub

Private Sub Label9_Click()

End Sub

Private Sub UserForm_Activate()
    TextBox_Name.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Name").ResultStr("")
    ComboBox6.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Note").ResultStr("")
    TextBox_Manufacturer.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").ResultStr("")
    
    ' Прописываем имя шэйпа.
    If Not Application.ActiveWindow.Selection.Count = 0 Then
        TextBox3.Text = Application.ActiveWindow.Selection(1).Name
    End If
End Sub
