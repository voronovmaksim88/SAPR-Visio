Attribute VB_Name = "PostgeSQL_Sensors"
Option Explicit

' Connection string (use the same as in other modules)
Const CONNECTION_STRING As String = "DSN=PostgreSQL_Vizio_x32;Uid=kis3admin;Pwd=kis3admin1313#;"

' Global array for storing sensor manufacturers
Public SensorsManufacturers() As SensorsManufacturerRecord

' Global array for storing sensor types
Public SensorsTypes() As SensorsType


' Global array for storing sensor measured values
Public SensorMeasuredValues() As SensorMeasuredValue

' Global array to store all Sensors table records
Public Sensors() As SensorRecord

' Filtered sensors
Public FilteredSensors() As SensorRecord

Sub ReadSensors()
    ' Declare connection object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Declare recordset object to store query results
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler
    
    ' Open connection
    conn.Open CONNECTION_STRING
    
    ' SQL query to get sensors from equipment table
    Dim sql As String
    sql = "SELECT e.id, e.name, e.model, e.vendor_code, e.description, " & _
          "e.manufacturer_id, e.price, e.currency_id, e.relevance, e.price_date, e.discriminator " & _
          "FROM equipment e " & _
          "INNER JOIN sensors s ON e.id = s.id " & _
          "WHERE e.discriminator = 'sensor'"
    
    ' Open recordset with static cursor to allow moving
    rs.Open sql, conn, 3, 1 ' 3 = adOpenStatic, 1 = adLockReadOnly
    
    ' Check if query returned any data
    If rs.EOF Then
        MsgBox "No sensors found in the database!", vbExclamation, "Error"
        GoTo Cleanup
    End If
    
    ' Count records to dimension the array
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst
    
    ' Resize the Sensors array
    ReDim Sensors(0 To recordCount - 1)
    
    ' Initialize result string for output
    Dim result As String
    result = "List of Sensors:" & vbCrLf & vbCrLf
    
    ' Loop through records and populate array
    Dim i As Long
    i = 0
    Do Until rs.EOF
        With Sensors(i)
            .ID = rs.Fields("id").value
            .Name = Nz(rs.Fields("name").value, "")
            .Model = Nz(rs.Fields("model").value, "")
            .VendorCode = Nz(rs.Fields("vendor_code").value, "")
            .Description = Nz(rs.Fields("description").value, "")
            .manufacturerID = Nz(rs.Fields("manufacturer_id").value, 0)
            .Price = Nz(rs.Fields("price").value, 0)
            .CurrencyID = Nz(rs.Fields("currency_id").value, 0)
            .Relevance = Nz(rs.Fields("relevance").value, True)
            .PriceDate = Nz(rs.Fields("price_date").value, #1/1/1900#)
            ' --- New: fill SensorTypes array for this sensor (array of Long IDs) ---
            Dim rsTypes As Object
            Set rsTypes = CreateObject("ADODB.Recordset")
            Dim sqlTypes As String
            sqlTypes = "SELECT sensor_type_id FROM sensor_types_association WHERE sensor_id = " & .ID
            rsTypes.Open sqlTypes, conn, 3, 1
            If Not rsTypes.EOF Then
                rsTypes.MoveLast
                Dim typeCount As Long
                typeCount = rsTypes.recordCount
                rsTypes.MoveFirst
                Dim arrTypeIDs() As Long
                ReDim arrTypeIDs(0 To typeCount - 1)
                Dim j As Long
                j = 0
                Do Until rsTypes.EOF
                    arrTypeIDs(j) = Nz(rsTypes.Fields("sensor_type_id").value, 0)
                    j = j + 1
                    rsTypes.MoveNext
                Loop
                .SensorTypes = arrTypeIDs
            Else
                ReDim .SensorTypes(-1 To -1) ' Empty array
            End If
            rsTypes.Close
            Set rsTypes = Nothing
            ' --- End new ---
        End With
        ' Append to result string for display
        Dim typeNames As String
        typeNames = ""
        If UBound(Sensors(i).SensorTypes) >= 0 Then
            Dim k As Long
            For k = LBound(Sensors(i).SensorTypes) To UBound(Sensors(i).SensorTypes)
                Dim typeID As Long
                typeID = Sensors(i).SensorTypes(k)
                Dim t As Long
                For t = LBound(SensorsTypes) To UBound(SensorsTypes)
                    If SensorsTypes(t).ID = typeID Then
                        If typeNames <> "" Then typeNames = typeNames & ", "
                        typeNames = typeNames & SensorsTypes(t).Name
                        Exit For
                    End If
                Next t
            Next k
        End If
        result = result & "ID: " & Sensors(i).ID & vbCrLf & _
                 "Name: " & Sensors(i).Name & vbCrLf & _
                 "Model: " & Sensors(i).Model & vbCrLf & _
                 "Vendor Code: " & Sensors(i).VendorCode & vbCrLf & _
                 "Description: " & Left(Sensors(i).Description, 50) & IIf(Len(Sensors(i).Description) > 50, "...", "") & vbCrLf & _
                 "Manufacturer ID: " & Sensors(i).manufacturerID & vbCrLf & _
                 "Price: " & Sensors(i).Price & vbCrLf & _
                 "Currency ID: " & Sensors(i).CurrencyID & vbCrLf & _
                 "Relevance: " & Sensors(i).Relevance & vbCrLf & _
                 "Price Date: " & Sensors(i).PriceDate & vbCrLf & _
                 "Sensor Types: " & typeNames & vbCrLf & vbCrLf
        i = i + 1
        rs.MoveNext
    Loop
    
    If MyDebug Then
        ' Display results
        MsgBox result, vbInformation, "Sensors List"
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



Sub ReadSensorsType()
    ' Declare connection object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Declare recordset object to store query results
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler
    
    ' Open connection
    conn.Open CONNECTION_STRING
    
    ' SQL query to get sensor types
    Dim sql As String
    sql = "SELECT id, name FROM sensor_types"
    
    ' Open recordset with static cursor to allow moving
    rs.Open sql, conn, 3, 1 ' 3 = adOpenStatic, 1 = adLockReadOnly
    
    ' Check if query returned any data
    If rs.EOF Then
        MsgBox "No sensor types found in the database!", vbExclamation, "Error"
        GoTo Cleanup
    End If
    
    ' Count records to dimension the array
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst
    
    ' Resize the SensorsTypes array
    ReDim SensorsTypes(0 To recordCount - 1)
    
    ' Loop through records and populate array
    Dim i As Long
    i = 0
    Do Until rs.EOF
        With SensorsTypes(i)
            .ID = Nz(rs.Fields("id").value, 0)
            .Name = Nz(rs.Fields("name").value, "")
        End With
        i = i + 1
        rs.MoveNext
    Loop

    ' Build a string to display the list of sensor types
    Dim result As String
    result = "List of Sensor Types:" & vbCrLf & vbCrLf
    For i = 0 To UBound(SensorsTypes)
        result = result & "ID: " & SensorsTypes(i).ID & ", Name: " & SensorsTypes(i).Name & vbCrLf
    Next i
    MsgBox result, vbInformation, "Sensor Types List"

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
    MsgBox "An error occurred: " & Err.Description, vbCritical, "ErrorHandler"
    GoTo Cleanup
End Sub




Sub ReadSensorMeasuredValue()
    ' Declare connection object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Declare recordset object to store query results
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler
    
    ' Open connection
    conn.Open CONNECTION_STRING
    
    ' SQL query to get measured values
    Dim sql As String
    sql = "SELECT id, name FROM sensor_measured_values"
    
    ' Open recordset with static cursor to allow moving
    rs.Open sql, conn, 3, 1 ' 3 = adOpenStatic, 1 = adLockReadOnly
    
    ' Check if query returned any data
    If rs.EOF Then
        MsgBox "No measured values found in the database!", vbExclamation, "Error"
        GoTo Cleanup
    End If
    
    ' Count records to dimension the array
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst
    
    ' Resize the SensorMeasuredValues array
    ReDim SensorMeasuredValues(0 To recordCount - 1)
    
    ' Loop through records and populate array
    Dim i As Long
    i = 0
    Do Until rs.EOF
        With SensorMeasuredValues(i)
            .ID = Nz(rs.Fields("id").value, 0)
            .Name = Nz(rs.Fields("name").value, "")
        End With
        i = i + 1
        rs.MoveNext
    Loop
    
    ' Build a string to display the list of measured values
    Dim result As String
    result = "List of Sensor Measured Values:" & vbCrLf & vbCrLf
    For i = 0 To UBound(SensorMeasuredValues)
        result = result & "ID: " & SensorMeasuredValues(i).ID & ", Name: " & SensorMeasuredValues(i).Name & vbCrLf
    Next i
    MsgBox result, vbInformation, "Sensor Measured Values List"

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
    MsgBox "An error occurred: " & Err.Description, vbCritical, "ErrorHandler"
    GoTo Cleanup
End Sub






Sub FilterSensors()
    ' Check if Sensors array is not empty
    If UBound(Sensors) < LBound(Sensors) Then
        ' If array is empty, exit
        Exit Sub
    End If
    
    ' Initialize FilteredSensors with the same size as Sensors
    ReDim FilteredSensors(LBound(Sensors) To UBound(Sensors))
    
    ' Copy all records from Sensors to FilteredSensors
    Dim i As Long
    For i = LBound(Sensors) To UBound(Sensors)
        FilteredSensors(i) = Sensors(i)
    Next i
End Sub


' Helper function to handle Null values
Private Function Nz(value As Variant, default As Variant) As Variant
    If IsNull(value) Then
        Nz = default
    Else
        Nz = value
    End If
End Function




