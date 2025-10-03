VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Spec 
   Caption         =   "Спецификация"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   6375
   OleObjectBlob   =   "Form_Spec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Spec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const SHAPE_TYPES_MAX As Integer = 31

Dim chkbox(0 To SHAPE_TYPES_MAX) As MSForms.CheckBox
Dim ShpTypeStr(0 To SHAPE_TYPES_MAX) As String
Dim ShapeTypesActual As Integer
Dim SortA As Integer
Dim SortB As Integer
Dim CancelSpecGeneration As Boolean


Private Sub Button_All_Click()
 ' "Cancel" button - closes the form without generating specification
 CancelSpecGeneration = True
 Set ShapeTypeExceptions = New Collection
 Unload Me
End Sub

Private Sub Button_AllBut_Click()
 ' Generates specification with exceptions for selected types
 CancelSpecGeneration = False
 Set ShapeTypeExceptions = New Collection
 Dim i As Integer
 For i = 0 To ShapeTypesActual - 1
  If (chkbox(i).value = -1) Then ShapeTypeExceptions.Add chkbox(i).Caption
 Next i
 Unload Me
End Sub

Private Sub Button_Only_Click()
 ' Generates specification only for selected types
 CancelSpecGeneration = False
 ShapeTypeExceptions.Clear
 Dim i As Integer
 For i = 0 To ShapeTypesActual - 1
  If (chkbox(i).value = 0) Then ShapeTypeExceptions.Add chkbox(i).Caption
 Next i
 Unload Me
End Sub

Function ShpTypeInList(ShpType As String) As Boolean
 Dim j As Integer
 ShpTypeInList = False
 For j = 0 To ShapeTypesActual - 1
  If ShpTypeStr(j) = ShpType Then
   ShpTypeInList = True
   Exit For
  End If
 Next j
End Function


Private Sub UserForm_Activate()
 CancelSpecGeneration = False
 Dim i As Integer
 Dim vPage As Visio.Page
 Dim vShape As Visio.Shape
 Dim vShapes As Visio.Shapes
 Dim sType As String
 
 For i = 0 To SHAPE_TYPES_MAX
  ShpTypeStr(i) = ""
 Next i
   
 ShapeTypesActual = 0
 
  For Each vPage In ActiveDocument.Pages ' Creating list cycle
   If (vPage.Index >= SortA) And (vPage.Index <= SortB) Then
    Set vShapes = vPage.Shapes
     For Each vShape In vShapes
        If vShape.CellExistsU("User.ShapeType", visExistsAnywhere) Then
         sType = vShape.CellsU("User.ShapeType").ResultStr("")
         If Not ShpTypeInList(sType) Then
          ShpTypeStr(ShapeTypesActual) = sType
          ' sorting
          For i = ShapeTypesActual - 1 To 0 Step -1
           If ShpTypeStr(i) > sType Then
            ShpTypeStr(i + 1) = ShpTypeStr(i)
            If (i = 0) Then
             ShpTypeStr(i) = sType
            End If
            Else
            ShpTypeStr(i + 1) = sType
            Exit For
           End If
          Next i
          ShapeTypesActual = ShapeTypesActual + 1
         End If
        End If
     Next vShape
   End If
  Next vPage
 
 
 For i = 0 To ShapeTypesActual - 1
  Set chkbox(i) = Controls.Add("Forms.CheckBox.1", "CheckBox" & i)
  chkbox(i).Caption = ShpTypeStr(i)
  chkbox(i).Left = 35 * (i Mod 8) + 15
  chkbox(i).Top = (i \ 8) * 25
  
  ' Set checkbox for types "CB" and "LineNum"
  If chkbox(i).Caption = "CB" Or chkbox(i).Caption = "LineNum" Then
   chkbox(i).value = True
  End If
  
 Next i
End Sub

Friend Sub ShowW(sub_sortA As Integer, sub_sortB As Integer)
 SortA = sub_sortA
 SortB = sub_sortB
 Me.Show
End Sub

Public Function IsSpecGenerationCancelled() As Boolean
 IsSpecGenerationCancelled = CancelSpecGeneration
End Function
