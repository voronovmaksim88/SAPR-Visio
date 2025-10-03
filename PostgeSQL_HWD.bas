Attribute VB_Name = "PostgeSQL_HWD"
' ===================================
' Module: PostgeSQL_HWD
' Purpose: Loading dimensions from 'heights', 'widths', 'depths' tables
' ===================================

Option Explicit

' Database connection string
Const CONNECTION_STRING As String = "DSN=PostgreSQL_Vizio_x32;Uid=kis3admin;Pwd=kis3admin1313#;"

' Global arrays to store all records from tables
Public Heights() As HeightRecord
Public Widths() As WidthRecord
Public Depths() As DepthRecord

Sub ReadHeights()
    ' Database connection objects
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler

    ' Open database connection
    conn.Open CONNECTION_STRING
    
    ' SQL query: get id and value from heights table, sorted by value
    Dim sql As String
    sql = "SELECT id, value FROM heights ORDER BY value"
    
    ' Execute query using static cursor and read-only mode
    rs.Open sql, conn, 3, 1  ' 3 = adOpenStatic, 1 = adLockReadOnly

    ' Check: if table has no data
    If rs.EOF Then
        MsgBox "Table 'heights' is empty or not found!", vbExclamation, "Warning"
        GoTo Cleanup
    End If

    ' Count records
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst

    ' Resize Heights array to match number of found records
    ReDim Heights(0 To recordCount - 1)

    ' String to display result in message
    Dim result As String
    result = "List of Heights:" & vbCrLf & vbCrLf

    ' Fill array with data from result set
    Dim i As Long
    For i = 0 To recordCount - 1
        With Heights(i)
            .ID = rs.Fields("id").value
            .value = Nz(rs.Fields("value").value, "")
        End With
        
        ' Add each record to result string
        result = result & "ID: " & Heights(i).ID & " | Value: " & Heights(i).value & vbCrLf
        
        rs.MoveNext
    Next i

    If MyDebug Then
        ' Show result (optional - can be removed in production)
        MsgBox result, vbInformation, "Heights List"
    End If
    
    GoTo Cleanup

ErrorHandler:
    MsgBox "Error loading heights: " & Err.Description, vbCritical, "Error"
    
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

Sub ReadWidths()
    ' Database connection objects
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler

    ' Open database connection
    conn.Open CONNECTION_STRING
    
    ' SQL query: get id and value from widths table, sorted by value
    Dim sql As String
    sql = "SELECT id, value FROM widths ORDER BY value"
    
    ' Execute query using static cursor and read-only mode
    rs.Open sql, conn, 3, 1  ' 3 = adOpenStatic, 1 = adLockReadOnly

    ' Check: if table has no data
    If rs.EOF Then
        MsgBox "Table 'widths' is empty or not found!", vbExclamation, "Warning"
        GoTo Cleanup
    End If

    ' Count records
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst

    ' Resize Widths array to match number of found records
    ReDim Widths(0 To recordCount - 1)

    ' String to display result in message
    Dim result As String
    result = "List of Widths:" & vbCrLf & vbCrLf

    ' Fill array with data from result set
    Dim i As Long
    For i = 0 To recordCount - 1
        With Widths(i)
            .ID = rs.Fields("id").value
            .value = Nz(rs.Fields("value").value, "")
        End With
        
        ' Add each record to result string
        result = result & "ID: " & Widths(i).ID & " | Value: " & Widths(i).value & vbCrLf
        
        rs.MoveNext
    Next i

    If MyDebug Then
        ' Show result (optional - can be removed in production)
        MsgBox result, vbInformation, "Widths List"
    End If
    
    GoTo Cleanup

ErrorHandler:
    MsgBox "Error loading widths: " & Err.Description, vbCritical, "Error"
    
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

Sub ReadDepths()
    ' Database connection objects
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler

    ' Open database connection
    conn.Open CONNECTION_STRING
    
    ' SQL query: get id and value from depths table, sorted by value
    Dim sql As String
    sql = "SELECT id, value FROM depths ORDER BY value"
    
    ' Execute query using static cursor and read-only mode
    rs.Open sql, conn, 3, 1  ' 3 = adOpenStatic, 1 = adLockReadOnly

    ' Check: if table has no data
    If rs.EOF Then
        MsgBox "Table 'depths' is empty or not found!", vbExclamation, "Warning"
        GoTo Cleanup
    End If

    ' Count records
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst

    ' Resize Depths array to match number of found records
    ReDim Depths(0 To recordCount - 1)

    ' String to display result in message
    Dim result As String
    result = "List of Depths:" & vbCrLf & vbCrLf

    ' Fill array with data from result set
    Dim i As Long
    For i = 0 To recordCount - 1
        With Depths(i)
            .ID = rs.Fields("id").value
            .value = Nz(rs.Fields("value").value, "")
        End With
        
        ' Add each record to result string
        result = result & "ID: " & Depths(i).ID & " | Value: " & Depths(i).value & vbCrLf
        
        rs.MoveNext
    Next i

    If MyDebug Then
        ' Show result (optional - can be removed in production)
        MsgBox result, vbInformation, "Depths List"
    End If

    GoTo Cleanup

ErrorHandler:
    MsgBox "Error loading depths: " & Err.Description, vbCritical, "Error"
    
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

