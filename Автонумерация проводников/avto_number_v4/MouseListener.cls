VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim WithEvents vsoWindow As Window
Attribute vsoWindow.VB_VarHelpID = -1
Dim iX As Double
Dim iY As Double

Dim n As Double
Dim x As Double
Dim y As Double


Private Sub Class_Initialize()
   Set vsoWindow = ActiveWindow
End Sub
Private Sub Class_Terminate()
   Set vsoWindow = Nothing
End Sub
Private Sub vsoWindow_MouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
   iX = x
   iY = y
End Sub
Private Sub vsoWindow_MouseUp(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
   If Button = 1 Then
       Debug.Print "Left mouse button released"
   ElseIf Button = 2 Then
       Debug.Print "Right mouse button released"
   ElseIf Button = 16 Then
       Debug.Print "Center mouse button released"
   End If

'MsgBox (Str(iX) + ", " + Str(iY))
If enable_num = True Then
    Set s2 = Documents("01_электрика.vss")
    Set mastObj = s2.Masters("number_v1")
    Set shpObj1 = ActivePage.Drop(mastObj, iX, iY)
    shpObj1.Text = CStr(s)
    s = s + 1
    
    End If

End Sub
