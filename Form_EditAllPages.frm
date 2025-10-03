VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_EditAllPages 
   Caption         =   "Параметры проекта"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   6420
   OleObjectBlob   =   "Form_EditAllPages.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_EditAllPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SortA As Integer, SortB As Integer

' форма для группового изменения данных странциц

Private Sub CommandButton1_Click()
Dim vPage As Visio.Page
Dim vShape As Visio.Shape
Dim vShapes As Visio.Shapes

For Each vPage In ActiveDocument.Pages
If (vPage.Index >= SortA) And (vPage.Index <= SortB) Then
Set vShapes = vPage.Shapes
For Each vShape In vShapes
If vShape.CellExistsU("Prop.Zayavka", visExistsAnywhere) Then
 If TextBox1.Text <> "" Then vShape.Cells("Prop.Zayavka").FormulaU = Chr(34) + TextBox1.Text + Chr(34)
End If
If vShape.CellExistsU("Prop.Zakazchik", visExistsAnywhere) Then
If TextBox2.Text <> "" Then vShape.Cells("Prop.Zakazchik").FormulaU = Chr(34) + TextBox2.Text + Chr(34)
End If
If vShape.CellExistsU("Prop.Razrabotchik", visExistsAnywhere) Then
If TextBox3.Text <> "" Then vShape.Cells("Prop.Razrabotchik").FormulaU = Chr(34) + TextBox3.Text + Chr(34)
End If
If vShape.CellExistsU("Prop.Nazvanie", visExistsAnywhere) Then
If TextBox4.Text <> "" Then vShape.Cells("Prop.Nazvanie").FormulaU = Chr(34) + TextBox4.Text + Chr(34)
End If
If vShape.CellExistsU("Prop.Prilozhenie", visExistsAnywhere) Then
If TextBox5.Text <> "" Then vShape.Cells("Prop.Prilozhenie").FormulaU = Chr(34) + TextBox5.Text + Chr(34)
End If

Next vShape
End If

Next vPage


    Form_EditAllPages.Hide
 
End Sub

Private Sub CommandButton2_Click()
    Form_EditAllPages.Hide
End Sub

Private Sub UserForm_Activate()

Dim vPage As Visio.Page
Dim vShape As Visio.Shape
Dim vShapes As Visio.Shapes
Static cSort As String

If cSort = "" Then cSort = "1-99"
    cSort = InputBox("Введите номер страницы или интервал (напр. 1-3)", "Заполнение полей", cSort)
    If InStr(cSort, "-") > 0 Then
        SortA = CInt(Left(cSort, InStr(cSort, "-") - 1))
        SortB = CInt(Right(cSort, Len(cSort) - InStr(cSort, "-")))
     Else
        SortA = Val(cSort)
        SortB = SortA
     End If
    If (SortA < 1) Or (SortB < 1) Then Exit Sub

For Each vPage In ActiveDocument.Pages
If (vPage.Index >= SortA) And (vPage.Index <= SortB) Then
Set vShapes = vPage.Shapes
For Each vShape In vShapes
If vShape.CellExistsU("Prop.Zayavka", visExistsAnywhere) Then TextBox1.Text = vShape.Cells("Prop.Zayavka").ResultStr("")
If vShape.CellExistsU("Prop.Zakazchik", visExistsAnywhere) Then TextBox2.Text = vShape.Cells("Prop.Zakazchik").ResultStr("")
If vShape.CellExistsU("Prop.Razrabotchik", visExistsAnywhere) Then TextBox3.Text = vShape.Cells("Prop.Razrabotchik").ResultStr("")
If vShape.CellExistsU("Prop.Nazvanie", visExistsAnywhere) Then TextBox4.Text = vShape.Cells("Prop.Nazvanie").ResultStr("")
If vShape.CellExistsU("Prop.Prilozhenie", visExistsAnywhere) Then TextBox5.Text = vShape.Cells("Prop.Prilozhenie").ResultStr("")
Next vShape
End If
Next vPage
 
End Sub

