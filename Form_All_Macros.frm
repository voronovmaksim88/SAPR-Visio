VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_All_Macros 
   Caption         =   "All_Macros"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1950
   OleObjectBlob   =   "Form_All_Macros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_All_Macros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Button_AutoEskiz_Click()
    Call AutoEskiz.AutoEskiz
    Unload Me
End Sub

Private Sub Button_autonum_Click()
    Call AutoNum
    Unload Me
End Sub

Private Sub Button_AutoSpec_Click()
    Call Autospec
    Unload Me
End Sub

Private Sub Button_EditAllPages_Click()
    Call EditAllPages.EditAllPages
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    Call AutoPageNum.AutoPageNum
    Unload Me
End Sub

Private Sub CommandButton_AutoNakleiki_Click()
    Call AutoNakleiki_v2.AutoNakleiki
    Unload Me
End Sub
