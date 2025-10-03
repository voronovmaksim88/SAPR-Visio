Attribute VB_Name = "PostgeSQL_Material"
Option Explicit
' Define structure for storing one record from control_cabinet_materials table

' Connection string constant
Const CONNECTION_STRING As String = "DSN=PostgreSQL_Vizio_x32;Uid=kis3admin;Pwd=kis3admin1313#;"



' Global array to store all records from control_cabinet_materials
Public Materials() As MaterialRecord


Sub ReadMaterials()
    ' Declare connection object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Declare recordset object to store query results
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler
    
    ' Open connection
    conn.Open CONNECTION_STRING
    
    ' SQL query to get id and name from control_cabinet_materials
    Dim sql As String
    sql = "SELECT id, name FROM control_cabinet_materials"
    
    ' Open recordset with static cursor to allow moving
    rs.Open sql, conn, 3, 1 ' 3 = adOpenStatic, 1 = adLockReadOnly
    
    ' Check if query returned any data
    If rs.EOF Then
        MsgBox "No materials found in the database!", vbExclamation, "Error"
        GoTo Cleanup
    End If
    
    ' Count records to dimension the array
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst
    
    ' Resize the Materials array
    ReDim Materials(0 To recordCount - 1)
    
    ' Initialize result string for output
    Dim result As String
    result = "List of Materials:" & vbCrLf & vbCrLf
    
    ' Loop through records and populate array
    Dim i As Long
    i = 0
    Do Until rs.EOF
        With Materials(i)
            .ID = rs.Fields("id").value
            .Name = Nz(rs.Fields("name").value, "")
        End With
        
        ' Append to result string for display
        result = result & "ID: " & Materials(i).ID & "   Name: " & Materials(i).Name & vbCrLf
        
        i = i + 1
        rs.MoveNext
    Loop
    
    ' Display results
    If MyDebug Then
        MsgBox result, vbInformation, "Materials List"
    End If
    
Cleanup:
    ' Clean up objects
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    GoTo Cleanup
End Sub

' Helper function to handle Null values
Private Function Nz(value As Variant, default As Variant) As Variant
    If IsNull(value) Then
        Nz = default
    Else
        Nz = value
    End If
End Function
