VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_E 
   Caption         =   "Form_E"
   ClientHeight    =   7140
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   5892
   OleObjectBlob   =   "Form_E.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_E"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'version 1.1.1 06-10-2017
' 1.1.1 Подтягивает Note из шейпа

'2018-01-16 Воронов МВ
'обновил путь к базе

'2018-06-05 Воронов МВ
'обновил путь к базе

Const ST_TABLENAME As String = "Sensors"
'Const ST_DBFILENAME As String = "db\Data_Base_Sibplc_v1.mdb"
'Const ST_DBFILENAME As String = "D:\YandexDisk\db\Data_Base_Sibplc_v13.mdb"
Const ST_DBFILENAME As String = "D:\SynologyDrive\work_main\db\Data_Base_Sibplc_v13.mdb"
Dim lObj_Dbs As DAO.Database
Dim LastChange As Integer

Dim ActiveFilters As New Collection


Function CollectionContains(myCol As Collection, checkVal As Variant) As Boolean
    On Error Resume Next
    CollectionContains = False
    Dim it As Variant
    For Each it In myCol
        If it = checkVal Then
            CollectionContains = True
            Exit Function
        End If
    Next
End Function

Sub ApplyFilters(BoxNum As Integer)

Dim lObj_Rs1 As DAO.Recordset
Dim sqlString As String

sqlString = " FROM " & ST_TABLENAME & " WHERE ID > 0 "
If (BoxNum = 1) Or (CollectionContains(ActiveFilters, 1)) Then sqlString = sqlString & " AND manufacturer = '" & ComboBox1.Text & "'"
If (BoxNum = 3) Or (CollectionContains(ActiveFilters, 3)) Then sqlString = sqlString & " AND type = '" & ComboBox3.Text & "'"
If (BoxNum = 4) Or (CollectionContains(ActiveFilters, 4)) Then sqlString = sqlString & " AND measured_value = '" & ComboBox4.Text & "'"
If (BoxNum = 5) Or (CollectionContains(ActiveFilters, 5)) Then sqlString = sqlString & " AND name = '" & ComboBox5.Text & "'"

LastChange = 99

If BoxNum <> 1 Then
ComboBox1.Clear
ComboBox1.Text = ""
Set lObj_Rs1 = lObj_Dbs.OpenRecordset("SELECT DISTINCT manufacturer " & sqlString)
 With lObj_Rs1
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then
         ComboBox1.AddItem (.Fields(0))
         If ComboBox1.Text = "" Then ComboBox1.Text = .Fields(0)
         End If
       .MoveNext
      Loop
      .Close
End With
End If

If BoxNum <> 3 Then
ComboBox3.Clear
ComboBox3.Text = ""
Set lObj_Rs1 = lObj_Dbs.OpenRecordset("SELECT DISTINCT type " & sqlString)
 With lObj_Rs1
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then
         ComboBox3.AddItem (.Fields(0))
         If ComboBox3.Text = "" Then ComboBox3.Text = .Fields(0)
         End If
       .MoveNext
      Loop
      .Close
End With
End If

If BoxNum <> 4 Then
ComboBox4.Clear
ComboBox4.Text = ""
Set lObj_Rs1 = lObj_Dbs.OpenRecordset("SELECT DISTINCT measured_value " & sqlString)
 With lObj_Rs1
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then
         ComboBox4.AddItem (.Fields(0))
         If ComboBox4.Text = "" Then ComboBox4.Text = .Fields(0)
         End If
       .MoveNext
      Loop
      .Close
End With
End If

If BoxNum <> 5 Then
ComboBox5.Clear
ComboBox5.Text = ""
Set lObj_Rs1 = lObj_Dbs.OpenRecordset("SELECT DISTINCT name " & sqlString)
 With lObj_Rs1
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then
         ComboBox5.AddItem (.Fields(0))
         If ComboBox5.Text = "" Then ComboBox5.Text = .Fields(0)
         End If
       .MoveNext
      Loop
      .Close
End With
End If

If BoxNum <> 6 Then
ComboBox6.Clear
ComboBox6.Text = ""
Set lObj_Rs1 = lObj_Dbs.OpenRecordset("SELECT DISTINCT model " & sqlString)
 With lObj_Rs1
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then
         ComboBox6.AddItem (.Fields(0))
         If ComboBox6.Text = "" Then ComboBox6.Text = .Fields(0)
         End If
       .MoveNext
      Loop
      .Close
End With
End If

LastChange = 0
End Sub


Private Sub ComboBox1_Change()

If LastChange > 0 Then Exit Sub
ApplyFilters (1)
  
End Sub


Private Sub ComboBox1_AfterUpdate()

ActiveFilters.Add (1)

End Sub

Private Sub ComboBox3_Change()

If LastChange > 0 Then Exit Sub
ApplyFilters (3)

End Sub

Private Sub ComboBox3_AfterUpdate()

ActiveFilters.Add (3)

End Sub

Private Sub ComboBox4_Change()

If LastChange > 0 Then Exit Sub
ApplyFilters (4)

End Sub

Private Sub ComboBox4_AfterUpdate()

ActiveFilters.Add (4)

End Sub


Private Sub ComboBox5_Change()

If LastChange > 0 Then Exit Sub
ApplyFilters (5)
  
End Sub


Private Sub ComboBox5_AfterUpdate()

ActiveFilters.Add (5)

End Sub

Private Sub ComboBox6_Change()

LastChange = 6
Dim lObj_Rs2 As DAO.Recordset

Set lObj_Rs2 = lObj_Dbs.OpenRecordset("SELECT TOP 1 note FROM " & ST_TABLENAME & " WHERE model='" & ComboBox6.Text & "'")

 With lObj_Rs2
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then TextBox2.Text = .Fields(0) Else TextBox2.Text = ""
      .MoveNext
      Loop
       .Close
   End With
   
Set lObj_Rs2 = lObj_Dbs.OpenRecordset("SELECT TOP 1 manufacturer FROM " & ST_TABLENAME & " WHERE model='" & ComboBox6.Text & "'")

 With lObj_Rs2
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox1.Text = .Fields(0) Else ComboBox1.Text = ""
      .MoveNext
      Loop
       .Close
   End With
   
Set lObj_Rs2 = lObj_Dbs.OpenRecordset("SELECT TOP 1 type FROM " & ST_TABLENAME & " WHERE model='" & ComboBox6.Text & "'")

 With lObj_Rs2
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox3.Text = .Fields(0) Else ComboBox3.Text = ""
      .MoveNext
      Loop
       .Close
   End With
   
Set lObj_Rs2 = lObj_Dbs.OpenRecordset("SELECT TOP 1 measured_value FROM " & ST_TABLENAME & " WHERE model='" & ComboBox6.Text & "'")

 With lObj_Rs2
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox4.Text = .Fields(0) Else ComboBox4.Text = ""
      .MoveNext
      Loop
       .Close
   End With
   
   Set lObj_Rs2 = lObj_Dbs.OpenRecordset("SELECT TOP 1 name FROM " & ST_TABLENAME & " WHERE model='" & ComboBox6.Text & "'")

 With lObj_Rs2
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox5.Text = .Fields(0) Else ComboBox5.Text = ""
      .MoveNext
      Loop
       .Close
   End With
   
 Dim suf As String
 If (ComboBox3.Text = "Дискретный") Or (ComboBox3.Text = "Д") Then
  suf = "S"
 ElseIf (ComboBox3.Text = "Интерфейсный") Or (ComboBox3.Text = "И") Then
  suf = "I"
 Else
  suf = "E"
 End If
  Select Case ComboBox4.Text
     Case "Температура"
       TextBox3.Text = "T"
     Case "Давление"
       TextBox3.Text = "P"
     Case "Перепад давления"
       TextBox3.Text = "PD"
     Case "Влажность"
       TextBox3.Text = "H"
     Case "Температура, Влажность"
       TextBox3.Text = "TH"
     Case "Скорость воздушного потока"
       TextBox3.Text = "Q"
     Case Else
       TextBox3.Text = ""
  End Select
  TextBox3.Text = TextBox3.Text + suf
   
   LastChange = 0

End Sub

Private Sub CommandButton1_Click()
 
 If (ComboBox6.Text <> "") Then

 ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + ComboBox1.Text + Chr(34)
 ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeNum").FormulaU = Chr(34) + ComboBox2.Text + Chr(34)

 If ActiveWindow.Selection.PrimaryItem.CellExists("prop.SensorType", 0) Then
  'to avoid errors
    If (ComboBox3.Text = "Дискретный") Or (ComboBox3.Text = "Д") Then
      ActiveWindow.Selection.PrimaryItem.Cells("Prop.SensorType").FormulaU = "1"
     ElseIf (ComboBox3.Text = "Интерфейсный") Or (ComboBox3.Text = "И") Then
      ActiveWindow.Selection.PrimaryItem.Cells("Prop.SensorType").FormulaU = "2"
     Else
      ActiveWindow.Selection.PrimaryItem.Cells("Prop.SensorType").FormulaU = "0"
    End If
 End If
 
 
' Dim mval As String
' Dim suf As String
' If (ComboBox3.Text = "Дискретный") Or (ComboBox3.Text = "Д") Then
'  suf = "S"
' ElseIf (ComboBox3.Text = "Интерфейсный") Or (ComboBox3.Text = "И") Then
'  suf = "I"
' Else
'  suf = "E"
' End If
'  Select Case ComboBox4.Text
'     Case "Температура"
'       mval = "T"
'     Case "Давление"
'       mval = "P"
'     Case "Перепад давления"
'       mval = "PD"
'     Case "Влажность"
'       mval = "H"
'     Case "Температура, Влажность"
'       mval = "TH"
'  End Select
'  TextBox3.Text = mval + suf
 
  ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeType").FormulaU = Chr(34) + TextBox3.Text + Chr(34)
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.MeasuredParameter").FormulaU = Chr(34) + Left(TextBox3.Text, Len(TextBox3.Text) - 1) + Chr(34)
    
  
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.Note").FormulaU = Chr(34) + TextBox2.Text + Chr(34)
  
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.Name").FormulaU = Chr(34) + ComboBox5.Text + Chr(34)
  

  
    'ActiveWindow.Selection.PrimaryItem.Cells("prop.Current").FormulaU = Chr(34) + ComboBox3.Text + Chr(34)
    'ActiveWindow.Selection.PrimaryItem.Cells("prop.Characteristic").FormulaU = Chr(34) + ComboBox4.Text + Chr(34)
    'ActiveWindow.Selection.PrimaryItem.Cells("User.Polus_number").FormulaU = Chr(34) + ComboBox5.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").FormulaU = Chr(34) + ComboBox6.Text + Chr(34)
    
    Dim vShape As Visio.shape
    Dim vSelection As Visio.Selection
    Dim shapeCount As Integer
    Dim i As Integer
    Set vSelection = Visio.ActiveWindow.Selection
    
    For Each vShape In vSelection
      shapeCount = vShape.Shapes.Count
      For i = 1 To shapeCount
       If InStr(1, vShape.Shapes(i).Text, "QF", vbTextCompare) > 0 Then
        vShape.Shapes(i).Text = "QF" + ComboBox2.Text
       End If
       
       If InStr(1, vShape.Shapes(i).Text, "B", vbTextCompare) > 0 Or InStr(1, vShape.Shapes(i).Text, "C", vbTextCompare) > 0 Or InStr(1, vShape.Shapes(i).Text, "D", vbTextCompare) > 0 Then
        vShape.Shapes(i).Text = ComboBox4.Text + ComboBox3.Text
       End If
       
      Next i
    Next vShape
    
    End If 'Model ComboBox is Empty
    
    lObj_Dbs.Close
    Form_E.Hide
    
    End Sub


Private Sub CommandButton2_Click()
lObj_Dbs.Close
Form_E.Hide
End Sub


Private Sub CommandButton3_Click()

 lObj_Dbs.Close
 Call UserForm_Activate
 
End Sub

Private Sub UserForm_Activate() ' выполняется при активации формы
Dim st_type As String
Dim st_mv As String
Dim st_man As String
Dim st_mod As String
Dim st As String
    
Dim i As Integer


Set lObj_Dbs = DAO.OpenDatabase(ST_DBFILENAME)
Dim lObj_Rs As DAO.Recordset

ComboBox1.Clear
Set lObj_Rs = lObj_Dbs.OpenRecordset("SELECT DISTINCT manufacturer FROM " & ST_TABLENAME)
'lObj_Dbs.Execute "SELECT * FROM Area"
 With lObj_Rs
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox1.AddItem (.Fields(0))
         .MoveNext
      Loop
       .Close
   End With

ComboBox3.Clear
Set lObj_Rs = lObj_Dbs.OpenRecordset("SELECT DISTINCT type FROM " & ST_TABLENAME)
 With lObj_Rs
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox3.AddItem (.Fields(0))
         .MoveNext
      Loop
       .Close
   End With

ComboBox4.Clear
Set lObj_Rs = lObj_Dbs.OpenRecordset("SELECT DISTINCT measured_value FROM " & ST_TABLENAME)
 With lObj_Rs
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox4.AddItem (.Fields(0))
         .MoveNext
      Loop
       .Close
   End With
   
ComboBox5.Clear
Set lObj_Rs = lObj_Dbs.OpenRecordset("SELECT DISTINCT name FROM " & ST_TABLENAME)
 With lObj_Rs
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox5.AddItem (.Fields(0))
         .MoveNext
      Loop
       .Close
   End With
   
ComboBox6.Clear
Set lObj_Rs = lObj_Dbs.OpenRecordset("SELECT DISTINCT model FROM " & ST_TABLENAME)
 With lObj_Rs
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox6.AddItem (.Fields(0))
         .MoveNext
      Loop
       .Close
  End With


 ComboBox2.Text = Format(ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeNum").ResultStr(""))
      
    'ComboBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").ResultStr("")
    
    'If ActiveWindow.Selection.PrimaryItem.CellExists("prop.SensorType", 0) Then
    ' to avoid errors
    'If ActiveWindow.Selection.PrimaryItem.Cells("prop.SensorType").ResultStr("") = "1" Then
    ' ComboBox3.Text = "Дискретный"
    'ElseIf ActiveWindow.Selection.PrimaryItem.Cells("prop.SensorType").ResultStr("") = "2" Then
    ' ComboBox3.Text = "Интерфейсный"
    'Else
    ' ComboBox3.Text = "Аналоговый"
    'End If
    'End If
    
    'TextBox3.Text = ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeType").ResultStr("")
    
    'ComboBox3.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Current").ResultStr("")
    'ComboBox4.Text = ActiveWindow.Selection.PrimaryItem.Cells("prop.Characteristic").ResultStr("")
    'ComboBox5.Text = Format(ActiveWindow.Selection.PrimaryItem.Cells("User.Polus_number").ResultStr(""))
    'TextBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").ResultStr("")
     
    ComboBox6.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").ResultStr("")
    Call ComboBox6_Change
    TextBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").ResultStr("")
    TextBox2.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Note").ResultStr("")
    
    For i = 1 To ActiveFilters.Count
    ActiveFilters.Remove (1)
    Next i
    
End Sub

Private Sub UserForm_Deactivate()

End Sub

