Attribute VB_Name = "Clear_Prop"
Sub Clear_Prop()
' Этот макрос предназначен для очистки Prop
' Имеет смысл запускать его каждый раз перед началом разработки схемы.
' Запуск необходимо делать только со страницы функциональной схемы иначе можно похерить весь проект.

    Dim vPage As Visio.Page ' Это страница
    Dim vShape As Visio.Shape ' Это фигура
    Dim vShapes As Visio.Shapes ' Это фигуры
    'Строка `Dim vShapes As Visio.Shapes` написана на языке программирования VBA (Visual Basic for Applications)
    'и используется в макросе среды программирования Microsoft Visio для объявления переменной `vShapes` типа `Visio.Shapes`.
    '`Visio.Shapes` является типом данных в Microsoft Visio, представляющим коллекцию фигур (shapes).
    'Коллекция `Visio.Shapes` содержит все фигуры, которые находятся на текущей странице в документе Visio.
    'С помощью строки `Dim vShapes As Visio.Shapes` мы объявляем переменную `vShapes`,
    'которая будет хранить ссылку на объект коллекции `Visio.Shapes`.
    'Это позволит нам обращаться к фигурам на текущей странице и выполнять с ними различные операции в макросе.
    
    Set vShapes = ActivePage.Shapes
    
    For Each vShape In vShapes
        If vShape.CellExistsU("Prop.Manufacturer", visExistsAnywhere) Then
            vShape.Cells("Prop.Manufacturer").FormulaU = """?"""
            ' Chr(34) - это аналог текстовой константы """"
        End If
        
        If vShape.CellExistsU("Prop.Model", visExistsAnywhere) Then
            vShape.Cells("Prop.Model").FormulaU = """?"""
        End If
        
        If vShape.CellExistsU("Prop.Note", visExistsAnywhere) Then
            vShape.Cells("Prop.Note").FormulaU = """?"""
        End If
    Next
    
    MsgBox ("Deleted Prop: Manufacturer, Model, Note")

End Sub
