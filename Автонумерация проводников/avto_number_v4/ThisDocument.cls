VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Dim myMouseListener As MouseListener'
Private Sub Document_DocumentSaved(ByVal doc As IVDocument)
'Set myMouseListener = New MouseListener'
End Sub
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
   'Set myMouseListener = Nothing'
End Sub
Sub num()
Attribute num.VB_ProcData.VB_Invoke_Func = "n"
UserForm1.Show
End Sub
