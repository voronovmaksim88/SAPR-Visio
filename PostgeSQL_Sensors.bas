Attribute VB_Name = "PostgeSQL_Sensors"
Option Explicit

' Connection string (use the same as in other modules)
Const CONNECTION_STRING As String = "DSN=PostgreSQL_Vizio_x32;Uid=kis3admin;Pwd=kis3admin1313#;"

' Global array for storing sensor types
Public SensorsTypes() As SensorsType


' Global array for storing sensor measured values
Public SensorMeasuredValues() As SensorMeasuredValue

' Global array to store all Sensors table records
Public Sensors() As SensorRecord

' Global array for storing sensor shape types
Public SensorShapeTypes() As SensorShapeType

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
          "e.manufacturer_id, e.price, e.currency_id, e.relevance, e.price_date, e.discriminator, " & _
          "s.sensor_type_id, s.sensors_shape_type_id, s.measured_value_id " & _
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
            ' Get sensor type ID directly from sensors table
            .SensorTypeID = Nz(rs.Fields("sensor_type_id").value, 0)
            ' Get shape type ID directly from sensors table
            .ShapeTypeID = Nz(rs.Fields("sensors_shape_type_id").value, 0)
            ' Get measured value ID directly from sensors table
            .MeasuredValueID = Nz(rs.Fields("measured_value_id").value, 0)
        End With
        ' Append to result string for display
        Dim typeID As Long
        typeID = Sensors(i).SensorTypeID
        Dim typeName As String
        typeName = ""
        Dim t As Long
        For t = LBound(SensorsTypes) To UBound(SensorsTypes)
            If SensorsTypes(t).ID = typeID Then
                typeName = SensorsTypes(t).Name
                Exit For
            End If
        Next t
        
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
                 "Sensor Type: " & typeName & vbCrLf & _
                 "Shape Type ID: " & Sensors(i).ShapeTypeID & vbCrLf & vbCrLf
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




Sub ReadSensorsShapeType()
    ' Declare connection object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Declare recordset object to store query results
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler
    
    ' Open connection
    conn.Open CONNECTION_STRING
    
    ' SQL query to get sensor shape types
    Dim sql As String
    sql = "SELECT id, name, shape_code FROM sensors_shape_type"
    
    ' Open recordset with static cursor to allow moving
    rs.Open sql, conn, 3, 1 ' 3 = adOpenStatic, 1 = adLockReadOnly
    
    ' Check if query returned any data
    If rs.EOF Then
        MsgBox "No sensor shape types found in the database!", vbExclamation, "Error"
        GoTo Cleanup
    End If
    
    ' Count records to dimension the array
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst
    
    ' Resize the SensorShapeTypes array
    ReDim SensorShapeTypes(0 To recordCount - 1)
    
    ' Loop through records and populate array
    Dim i As Long
    i = 0
    Do Until rs.EOF
        With SensorShapeTypes(i)
            .ID = Nz(rs.Fields("id").value, 0)
            .Name = Nz(rs.Fields("name").value, "")
            .ShapeCode = Nz(rs.Fields("shape_code").value, "")
        End With
        i = i + 1
        rs.MoveNext
    Loop

    ' Build a string to display the list of sensor shape types
    Dim result As String
    result = "List of Sensor Shape Types:" & vbCrLf & vbCrLf
    For i = 0 To UBound(SensorShapeTypes)
        result = result & "ID: " & SensorShapeTypes(i).ID & ", Name: " & SensorShapeTypes(i).Name & ", ShapeCode: " & SensorShapeTypes(i).ShapeCode & vbCrLf
    Next i
    MsgBox result, vbInformation, "Sensor Shape Types List"

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
    
    ' Get selected manufacturer from ComboBox
    Dim selectedManufacturer As String
    selectedManufacturer = Form_Sensors_PostgreSQL.ComboBox_Manufacturer.Text
    ApplyManufacturerFilter selectedManufacturer
    
    ' Get selected sensor type from ComboBox
    Dim selectedSensorType As String
    selectedSensorType = Form_Sensors_PostgreSQL.ComboBox_SensorType.Text
    ApplySensorTypeFilter selectedSensorType
    
    ' Get selected measured value from ComboBox and apply filter if needed
    Dim selectedMeasuredValue As String
    selectedMeasuredValue = Form_Sensors_PostgreSQL.ComboBox_SensorMeasuredValue.Text
    ApplyMeasuredValueFilter selectedMeasuredValue
    
    ' Get selected model from ComboBox and apply filter if needed
    Dim selectedModel As String
    selectedModel = Form_Sensors_PostgreSQL.ComboBox_Model.Text
    ApplyModelFilter selectedModel
    
    ' Get selected name from ComboBox and apply filter if needed
    Dim selectedName As String
    selectedName = Form_Sensors_PostgreSQL.ComboBox_Name.Text
    ApplyNameFilter selectedName
    
    ' Display final record count in LabelNum
    Dim recordCount As Long
    On Error Resume Next
    recordCount = UBound(FilteredSensors) - LBound(FilteredSensors) + 1
    If Err.Number <> 0 Then
        recordCount = 0
        Err.Clear
    End If
    On Error GoTo 0
    Form_Sensors_PostgreSQL.Label_NumOfRecord.Caption = "Number of records: " & recordCount
End Sub


' Helper function to handle Null values
Private Function Nz(value As Variant, default As Variant) As Variant
    If IsNull(value) Then
        Nz = default
    Else
        Nz = value
    End If
End Function

Private Sub ApplyManufacturerFilter(selectedManufacturer As String)
    ' If "all" is selected or empty, skip filtering
    If selectedManufacturer = "all" Or selectedManufacturer = "" Then
        Exit Sub
    End If
    
    ' Find ID of selected manufacturer
    Dim targetManufacturerID As Long
    targetManufacturerID = 0
    Dim j As Long
    
    ' Check if Manufacturers array is empty
    On Error Resume Next
    If UBound(Manufacturers) < LBound(Manufacturers) Then
        ' Array is empty, exit
        Exit Sub
    End If
    On Error GoTo 0
    
    For j = LBound(Manufacturers) To UBound(Manufacturers)
        If Manufacturers(j).Name = selectedManufacturer Then
            targetManufacturerID = Manufacturers(j).ID
            Exit For
        End If
    Next j
    
    ' If manufacturer not found, exit
    If targetManufacturerID = 0 Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where manufacturerID matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As SensorRecord
    
    ' Create temporary array for filtered records
    ReDim tempArray(LBound(FilteredSensors) To UBound(FilteredSensors))
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredSensors) To UBound(FilteredSensors)
        If FilteredSensors(i).manufacturerID = targetManufacturerID Then
            tempArray(filteredCount) = FilteredSensors(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredSensors to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredSensors(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredSensors(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array
        ReDim FilteredSensors(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredSensors
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub

Private Sub ApplySensorTypeFilter(selectedSensorType As String)
    ' If "all" is selected or empty, skip filtering
    If selectedSensorType = "all" Or selectedSensorType = "" Then
        Exit Sub
    End If
    
    ' Find ID of selected sensor type
    Dim targetSensorTypeID As Long
    targetSensorTypeID = 0
    Dim j As Long
    
    ' Check if SensorsTypes array is empty
    On Error Resume Next
    If UBound(SensorsTypes) < LBound(SensorsTypes) Then
        ' Array is empty, exit
        Exit Sub
    End If
    On Error GoTo 0
    
    For j = LBound(SensorsTypes) To UBound(SensorsTypes)
        If SensorsTypes(j).Name = selectedSensorType Then
            targetSensorTypeID = SensorsTypes(j).ID
            Exit For
        End If
    Next j
    
    ' If sensor type not found, exit
    If targetSensorTypeID = 0 Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where SensorTypeID matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As SensorRecord
    
    ' Create temporary array for filtered records
    ReDim tempArray(LBound(FilteredSensors) To UBound(FilteredSensors))
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredSensors) To UBound(FilteredSensors)
        If FilteredSensors(i).SensorTypeID = targetSensorTypeID Then
            tempArray(filteredCount) = FilteredSensors(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredSensors to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredSensors(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredSensors(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array
        ReDim FilteredSensors(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredSensors
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub




Private Sub ApplyMeasuredValueFilter(selectedMeasuredValue As String)
    ' If "all" is selected or empty, skip filtering
    If selectedMeasuredValue = "all" Or selectedMeasuredValue = "" Then
        Exit Sub
    End If
    
    ' Find ID of selected measured value
    Dim targetMeasuredValueID As Long
    targetMeasuredValueID = 0
    Dim j As Long
    
    ' Check if SensorMeasuredValues array is empty
    On Error Resume Next
    If UBound(SensorMeasuredValues) < LBound(SensorMeasuredValues) Then
        ' Array is empty, exit
        Exit Sub
    End If
    On Error GoTo 0
    
    For j = LBound(SensorMeasuredValues) To UBound(SensorMeasuredValues)
        If SensorMeasuredValues(j).Name = selectedMeasuredValue Then
            targetMeasuredValueID = SensorMeasuredValues(j).ID
            Exit For
        End If
    Next j
    
    ' If measured value not found, exit
    If targetMeasuredValueID = 0 Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where MeasuredValueID matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As SensorRecord
    
    ' Create temporary array for filtered records
    On Error Resume Next
    ReDim tempArray(LBound(FilteredSensors) To UBound(FilteredSensors))
    If Err.Number <> 0 Then
        ' Array is empty, nothing to filter
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredSensors) To UBound(FilteredSensors)
        If FilteredSensors(i).MeasuredValueID = targetMeasuredValueID Then
            tempArray(filteredCount) = FilteredSensors(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredSensors to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredSensors(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredSensors(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array
        ReDim FilteredSensors(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredSensors
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub



Private Sub ApplyModelFilter(selectedModel As String)
    ' If "all" is selected or empty, skip filtering
    If selectedModel = "all" Or selectedModel = "" Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where Model matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As SensorRecord
    
    ' Create temporary array for filtered records
    On Error Resume Next
    ReDim tempArray(LBound(FilteredSensors) To UBound(FilteredSensors))
    If Err.Number <> 0 Then
        ' Array is empty, nothing to filter
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredSensors) To UBound(FilteredSensors)
        If FilteredSensors(i).Model = selectedModel Then
            tempArray(filteredCount) = FilteredSensors(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredSensors to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredSensors(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredSensors(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array
        ReDim FilteredSensors(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredSensors
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub


Private Sub ApplyNameFilter(selectedName As String)
    ' If "all" is selected or empty, skip filtering
    If selectedName = "all" Or selectedName = "" Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where Name matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As SensorRecord
    
    ' Create temporary array for filtered records
    On Error Resume Next
    ReDim tempArray(LBound(FilteredSensors) To UBound(FilteredSensors))
    If Err.Number <> 0 Then
        ' Array is empty, nothing to filter
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredSensors) To UBound(FilteredSensors)
        If FilteredSensors(i).Name = selectedName Then
            tempArray(filteredCount) = FilteredSensors(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredSensors to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredSensors(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredSensors(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array
        ReDim FilteredSensors(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredSensors
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub



