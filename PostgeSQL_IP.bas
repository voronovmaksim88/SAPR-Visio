Attribute VB_Name = "PostgeSQL_IP"
' ===================================
' Module: PostgreSQL_IP
' Purpose: Loading IP protection degrees from 'ips' table
' ===================================

Option Explicit

' Database connection string
Const CONNECTION_STRING As String = "DSN=PostgreSQL_Vizio_x32;Uid=kis3admin;Pwd=kis3admin1313#;"

' Global array to store all records from ips table
Public IPs() As IpRecord

Sub ReadIPs()
    ' Database connection objects
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler

    ' Open database connection
    conn.Open CONNECTION_STRING
    
    ' SQL query: get id and name from ips table, sorted by name
    Dim sql As String
    sql = "SELECT id, name FROM ips ORDER BY name"
    
    ' Execute query using static cursor and read-only mode
    rs.Open sql, conn, 3, 1  ' 3 = adOpenStatic, 1 = adLockReadOnly

    ' Check: if table has no data
    If rs.EOF Then
        MsgBox "Table 'ips' is empty or not found!", vbExclamation, "Warning"
        GoTo Cleanup
    End If

    ' Count records
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst

    ' Resize IPs array to match number of found records
    ReDim IPs(0 To recordCount - 1)

    ' String to display result in message
    Dim result As String
    result = "List of IP protection degrees:" & vbCrLf & vbCrLf

    ' Fill array with data from result set
    Dim i As Long
    For i = 0 To recordCount - 1
        With IPs(i)
            .ID = rs.Fields("id").value
            .Name = Nz(rs.Fields("name").value, "")
        End With
        
        ' Add each record to result string
        result = result & "ID: " & IPs(i).ID & " | Name: " & IPs(i).Name & vbCrLf
        
        rs.MoveNext
    Next i

    If MyDebug Then
        ' Show result (optional - can be removed in production)
        MsgBox result, vbInformation, "IP Protection Degrees"
    End If

    GoTo Cleanup

ErrorHandler:
    MsgBox "Error loading IP protection degrees: " & Err.Description, vbCritical, "Error"
    
Cleanup:
    ' Free resources: close objects and clear memory
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Sub

' Helper function to handle NULL values from database
' Analog of Nz function in Access
Private Function Nz(value As Variant, default As Variant) As Variant
    If IsNull(value) Then
        Nz = default
    Else
        Nz = value
    End If
End Function
