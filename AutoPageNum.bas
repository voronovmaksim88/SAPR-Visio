Attribute VB_Name = "AutoPageNum"
Sub AutoPageNum()
Dim vPage As Visio.Page
Dim vShape As Visio.Shape
Dim vShapes As Visio.Shapes
Dim GroupLen(8) As Integer
Dim Count As Integer
Dim lGroup As Integer
Dim tGroup As Integer
Dim GroupName As String
Count = 0

For Each vPage In ActiveDocument.Pages
Set vShapes = vPage.Shapes
For Each vShape In vShapes
If vShape.CellExistsU("User.PageNum", visExistsAnywhere) Then
 Count = Count + 1
 If vShape.Cells("Prop.Prilozhenie").ResultStr("") <> GroupName Then
  tGroup = tGroup + 1
  If (tGroup > 1) Then GroupLen(tGroup - 1) = Count - lGroup
  lGroup = Count
  GroupName = vShape.Cells("Prop.Prilozhenie").ResultStr("")
 End If
 vShape.Cells("User.PageNum").Formula = Count - lGroup + 1
End If
Next vShape
Next vPage
GroupLen(tGroup) = Count - lGroup + 1

lGroup = 1: Count = 0
For Each vPage In ActiveDocument.Pages
Set vShapes = vPage.Shapes
For Each vShape In vShapes
If vShape.CellExistsU("User.PageNum", visExistsAnywhere) Then
 Count = Count + 1
 If Count > GroupLen(lGroup) Then
  Count = 1
  lGroup = lGroup + 1
 End If
 vShape.Cells("User.PageTotal").Formula = GroupLen(lGroup)
End If
Next vShape
Next vPage


End Sub
