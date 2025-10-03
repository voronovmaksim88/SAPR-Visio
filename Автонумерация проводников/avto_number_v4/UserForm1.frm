VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "помощник нумерации"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14040
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Double
Dim x As Double
Dim y As Double
Dim myMouseListener As MouseListener
Sub Reg()
    Set VisApp = Visio.Application
End Sub

Private Sub VisApp_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
Set s2 = Documents("01_электрика.vss")
Set mastObj = s2.Masters("number_v1")
Set shpObj1 = ActivePage.Drop(mastObj, n, 10)
shpObj1.Text = CStr(s)
s = s + 1
n = n + 0.4
End Sub

Private Sub CommandButton2_Click()
Set s2 = Documents("01_электрика.vss")
Set mastObj = s2.Masters("number_v1")
Set shpObj1 = ActivePage.Drop(mastObj, n, 10)
shpObj1.Text = CStr(s)
s = s + 1
n = n + 0.4
End Sub

Private Sub CommandButton5_Click()
enable_num = Not (enable_num)
Set myMouseListener = New MouseListener

If enable_num = True Then
CommandButton5.Caption = "остановить нумерацию"
End If

If enable_num = False Then
CommandButton5.Caption = "начать нумерацию"
End If

Label4.Caption = CStr(s)
End Sub

Private Sub CommandButton6_Click() ' продолжить с последнего
Dim text1 As String
Dim i1 As Integer
text1 = "number_v1"
s = 0
For i = 1 To (ActivePage.Shapes.Count)
If InStr(ActivePage.Shapes(i).Name, text1) > 0 Then 'прорверяем если в имени шэйпа "number_v1"

' проверяем что в текущем шэйпе "number_v1" не содержаться символы букв которыми мы обозначаем фазы и нейтраль
' это не очень универсально но нам пойдёт
If _
(InStr(ActivePage.Shapes(i).Text, "A") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "B") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "C") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "А") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "В") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "С") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "N") = 0) _
Then
    
    If CInt(ActivePage.Shapes(i).Text) > s Then
    s = CInt(ActivePage.Shapes(i).Text)
    End If

End If

End If

Next
s = s + 1
Label4.Caption = CStr(s)
End Sub

Private Sub Label1_Click()
s = Val(Label1.Caption)
End Sub

Private Sub CommandButton3_Click()
s = Val(TextBox1.Text)
Label4.Caption = CStr(s)
End Sub

Private Sub CommandButton4_Click()
Dim text1 As String
Dim i1 As Integer
text1 = "number_v1"

For i = 1 To ActivePage.Shapes.Count
If InStr(ActivePage.Shapes(i).Name, text1) > 0 Then 'прорверяем если в имени шэйпа "number_v1"

' проверяем что в текущем шэйпе "number_v1" не содержаться символы букв которыми мы обозначаем фазы и нейтраль
' это не очень универсально но нам пойдёт
If (InStr(ActivePage.Shapes(i).Text, "A") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "B") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "C") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "А") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "В") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "С") = 0) And _
(InStr(ActivePage.Shapes(i).Text, "N") = 0) Then

    If CInt(ActivePage.Shapes(i).Text) >= CInt(TextBox2.Text) And CInt(ActivePage.Shapes(i).Text) <= CInt(TextBox3.Text) Then
    ActivePage.Shapes(i).Text = ActivePage.Shapes(i).Text + CInt(TextBox4.Text)
    End If

End If

End If

Next

End Sub

Private Sub Label2_Click()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Click()
Dim text1 As String
Dim int1 As Integer
text1 = abcd
int1 = InStr(1, "d", text1, vbTextCompare)
UserForm1.Caption = int1
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X1 As Single, ByVal Y1 As Single)

x = X1
y = Y1

End Sub


