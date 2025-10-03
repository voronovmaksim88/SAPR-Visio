Attribute VB_Name = "EditAllImageSize"
' макрос для автоматического выравнивания по размеру всех картинок на листе
' удобен при оздании структуры меню
Sub EditAllImageSize()
Dim Width As String
Dim Height As String
Dim i As Integer

Width = CStr(70) + " mm"
Height = CStr(30) + " mm"
 
For i = 1 To ActivePage.Shapes.Count
    Application.ActiveWindow.Page.Shapes.ItemFromID(i).CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = Width
    Application.ActiveWindow.Page.Shapes.ItemFromID(i).CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = Height
Next

End Sub

