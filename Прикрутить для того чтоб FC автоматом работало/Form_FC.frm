VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_FC 
   Caption         =   "Form_FC"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   5895
   OleObjectBlob   =   "Form_FC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_FC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Const ST_TABLENAME As String = "FC"
'Const ST_DBFILENAME As String = "db\Data_Base_Sibplc_v13.mdb"
Const ST_DBFILENAME As String = "D:\YandexDisk\db\Data_Base_Sibplc_v12.mdb"

Dim lObj_Dbs As DAO.Database
Dim LastChange As Integer
Dim pin As Integer
Dim uin As Integer
Dim pout As Integer
Dim uout As Integer

Dim ActiveFilters As New Collection

Function DotDouble(num As String) As String
 DotDouble = Replace$(num, ",", ".")
End Function

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
If (BoxNum = 3) Or (CollectionContains(ActiveFilters, 3)) Then sqlString = sqlString & " AND power = " & DotDouble(ComboBox3.Text)
If (BoxNum = 4) Or (CollectionContains(ActiveFilters, 4)) Then sqlString = sqlString & " AND iin = " & DotDouble(ComboBox4.Text)
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
Set lObj_Rs1 = lObj_Dbs.OpenRecordset("SELECT DISTINCT power " & sqlString)
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
Set lObj_Rs1 = lObj_Dbs.OpenRecordset("SELECT DISTINCT iin " & sqlString)
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
   
Set lObj_Rs2 = lObj_Dbs.OpenRecordset("SELECT TOP 1 power FROM " & ST_TABLENAME & " WHERE model='" & ComboBox6.Text & "'")

 With lObj_Rs2
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox3.Text = .Fields(0) Else ComboBox3.Text = ""
      .MoveNext
      Loop
       .Close
   End With
   
Set lObj_Rs2 = lObj_Dbs.OpenRecordset("SELECT TOP 1 iin FROM " & ST_TABLENAME & " WHERE model='" & ComboBox6.Text & "'")

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
   
   Set lObj_Rs2 = lObj_Dbs.OpenRecordset("SELECT TOP 1 phasein,uin,phaseout,uout FROM " & ST_TABLENAME & " WHERE model='" & ComboBox6.Text & "'")
 With lObj_Rs2
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then pin = .Fields(0) Else pin = 0
         If Not IsNull(.Fields(1)) Then uin = .Fields(1) Else uin = 0
         If Not IsNull(.Fields(2)) Then pout = .Fields(2) Else pout = 0
         If Not IsNull(.Fields(2)) Then uout = .Fields(3) Else uout = 0
      .MoveNext
      Loop
       .Close
   End With

   
   
   TextBox3.Text = pin & "*" & uin & " / " & pout & "*" & uout
   
   LastChange = 0

End Sub

Private Sub CommandButton1_Click()
 
 If (ComboBox6.Text <> "") Then

 ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + ComboBox1.Text + Chr(34)
 ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeNum").FormulaU = Chr(34) + ComboBox2.Text + Chr(34)

  ActiveWindow.Selection.PrimaryItem.Cells("User.ShapeType").FormulaU = Chr(34) + "FC" + Chr(34)
   
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.Note").FormulaU = Chr(34) + TextBox2.Text + Chr(34)
  
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.power").FormulaU = Chr(34) + ComboBox3.Text + Chr(34)
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.iin").FormulaU = Chr(34) + ComboBox4.Text + Chr(34)
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.Name").FormulaU = Chr(34) + ComboBox5.Text + Chr(34)
  
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.PhaseIn").FormulaU = pin
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.UIn").FormulaU = uin
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.PhaseOut").FormulaU = pout
  ActiveWindow.Selection.PrimaryItem.Cells("Prop.UOut").FormulaU = uout
  
  

  
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").FormulaU = Chr(34) + ComboBox6.Text + Chr(34)
    
    
      
    
    End If 'Model ComboBox is Empty
    
    lObj_Dbs.Close
    Form_FC.Hide
    
    End Sub


Private Sub CommandButton2_Click()
lObj_Dbs.Close
Form_FC.Hide
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


LastChange = 101


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
Set lObj_Rs = lObj_Dbs.OpenRecordset("SELECT DISTINCT power FROM " & ST_TABLENAME)
 With lObj_Rs
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then ComboBox3.AddItem (.Fields(0))
         .MoveNext
      Loop
       .Close
   End With

ComboBox4.Clear
Set lObj_Rs = lObj_Dbs.OpenRecordset("SELECT DISTINCT iin FROM " & ST_TABLENAME)
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
      
     
    ComboBox6.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").ResultStr("")
    Call ComboBox6_Change
    TextBox1.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").ResultStr("")
    TextBox2.Text = ActiveWindow.Selection.PrimaryItem.Cells("Prop.Note").ResultStr("")
    
    For i = 1 To ActiveFilters.Count
     ActiveFilters.Remove (1)
    Next i
    LastChange = 0
    
End Sub

Private Sub UserForm_Deactivate()

End Sub

