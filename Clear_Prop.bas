Attribute VB_Name = "Clear_Prop"
Sub Clear_Prop()
' ���� ������ ������������ ��� ������� Prop
' ����� ����� ��������� ��� ������ ��� ����� ������� ���������� �����.
' ������ ���������� ������ ������ �� �������� �������������� ����� ����� ����� �������� ���� ������.

    Dim vPage As Visio.Page ' ��� ��������
    Dim vShape As Visio.Shape ' ��� ������
    Dim vShapes As Visio.Shapes ' ��� ������
    '������ `Dim vShapes As Visio.Shapes` �������� �� ����� ���������������� VBA (Visual Basic for Applications)
    '� ������������ � ������� ����� ���������������� Microsoft Visio ��� ���������� ���������� `vShapes` ���� `Visio.Shapes`.
    '`Visio.Shapes` �������� ����� ������ � Microsoft Visio, �������������� ��������� ����� (shapes).
    '��������� `Visio.Shapes` �������� ��� ������, ������� ��������� �� ������� �������� � ��������� Visio.
    '� ������� ������ `Dim vShapes As Visio.Shapes` �� ��������� ���������� `vShapes`,
    '������� ����� ������� ������ �� ������ ��������� `Visio.Shapes`.
    '��� �������� ��� ���������� � ������� �� ������� �������� � ��������� � ���� ��������� �������� � �������.
    
    Set vShapes = ActivePage.Shapes
    
    For Each vShape In vShapes
        If vShape.CellExistsU("Prop.Manufacturer", visExistsAnywhere) Then
            vShape.Cells("Prop.Manufacturer").FormulaU = """?"""
            ' Chr(34) - ��� ������ ��������� ��������� """"
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
