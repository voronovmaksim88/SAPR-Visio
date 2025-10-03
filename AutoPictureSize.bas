Attribute VB_Name = "AutoPictureSize"
Sub AutoPictureSize()
' для того чтоб подогнать все картинки к одному размеру надо чтоб на листе больше ничего не было кроме этих картинок
Dim width As String
Dim Height As String
Dim i As Integer

width = CStr(50) + " mm"
Height = CStr(20) + " mm"
 
For i = 1 To ActivePage.Shapes.Count
    Application.ActiveWindow.Page.Shapes.ItemFromID(i).CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = width
    Application.ActiveWindow.Page.Shapes.ItemFromID(i).CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = Height
Next

End Sub

