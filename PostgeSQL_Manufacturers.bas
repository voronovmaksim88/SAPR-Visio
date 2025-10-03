Attribute VB_Name = "PostgeSQL_Manufacturers"
Option Explicit

' Connection string constant
Const CONNECTION_STRING As String = "DSN=PostgreSQL_Vizio_x32;Uid=kis3admin;Pwd=kis3admin1313#;"


' Global array for Keeping all records table Manufacturers
Public Manufacturers() As ManufacturerRecord

Sub ReadManufacturers()
    ' Declare connection object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Declare recordset object to store query results
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler
    
    ' Open connection
    conn.Open CONNECTION_STRING
    
    ' SQL query to get id and name from manufacturers
    Dim sql As String
    sql = "SELECT id, name FROM manufacturers"
    
    ' Open recordset with static cursor to allow moving
    rs.Open sql, conn, 3, 1 ' 3 = adOpenStatic, 1 = adLockReadOnly
    
    ' Check if query returned any data
    If rs.EOF Then
        MsgBox "No manufacturers found in the database!", vbExclamation, "Error"
        GoTo Cleanup
    End If
    
    ' Count records to dimension the array
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst
    
    ' Resize the Manufacturers array
    ReDim Manufacturers(0 To recordCount - 1)
    
    ' Initialize result string for output
    Dim result As String
    result = "List of Manufacturers:" & vbCrLf & vbCrLf
    
    ' Loop through records and populate array
    Dim i As Long
    i = 0
    Do Until rs.EOF
        With Manufacturers(i)
            .ID = rs.Fields("id").value
            .Name = Nz(rs.Fields("name").value, "")
        End With
        
        ' Append to result string for display
        result = result & "ID: " & Manufacturers(i).ID & "   Name: " & Manufacturers(i).Name & vbCrLf
        
        i = i + 1
        rs.MoveNext
    Loop
    
    If MyDebug Then
    ' Display results
        MsgBox result, vbInformation, "Manufacturers List"
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

