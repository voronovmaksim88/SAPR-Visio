Attribute VB_Name = "PostgeSQL_ControlCabinets"
Option Explicit

' Connection string constant
Const CONNECTION_STRING As String = "DSN=PostgreSQL_Vizio_x32;Uid=kis3admin;Pwd=kis3admin1313#;"

' Global array to store all ControlCabinets table records
Public Cabinets() As ControlCabinetRecord

' Filtered cabinets
Public FilteredControlCabinets() As ControlCabinetRecord

Sub ReadControlCabinets()
    ' Declare connection object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Declare recordset object to store query results
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ErrorHandler
    
    ' Open connection
    conn.Open CONNECTION_STRING
    
    ' SQL query to join control_cabinets with equipment
    Dim sql As String
    sql = "SELECT e.id, e.name, e.model, e.vendor_code, e.description, " & _
          "e.manufacturer_id, e.price, e.currency_id, e.relevance, e.price_date, " & _
          "c.material_id, c.ip_id, c.height_id, c.width_id, c.depth_id " & _
          "FROM equipment e " & _
          "INNER JOIN control_cabinets c ON e.id = c.id " & _
          "WHERE e.discriminator = 'control_cabinet'"
    
    ' Open recordset with static cursor to allow moving
    rs.Open sql, conn, 3, 1 ' 3 = adOpenStatic, 1 = adLockReadOnly
    
    ' Check if query returned any data
    If rs.EOF Then
        MsgBox "No control cabinets found in the database!", vbExclamation, "Error"
        GoTo Cleanup
    End If
    
    ' Count records to dimension the array
    rs.MoveLast
    Dim recordCount As Long
    recordCount = rs.recordCount
    rs.MoveFirst
    
    ' Resize the Cabinets array
    ReDim Cabinets(0 To recordCount - 1)
    
    ' Initialize result string for output
    Dim result As String
    result = "List of Control Cabinets:" & vbCrLf & vbCrLf
    
    ' Loop through records and populate array
    Dim i As Long
    i = 0
    Do Until rs.EOF
        With Cabinets(i)
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
            .materialID = Nz(rs.Fields("material_id").value, 0)
            .ipID = Nz(rs.Fields("ip_id").value, 0)
            .height_id = Nz(rs.Fields("height_id").value, 0)
            .width_id = Nz(rs.Fields("width_id").value, 0)
            .depth_id = Nz(rs.Fields("depth_id").value, 0)
        End With
        
        ' Append to result string for display
        result = result & "ID: " & Cabinets(i).ID & vbCrLf & _
                 "Name: " & Cabinets(i).Name & vbCrLf & _
                 "Model: " & Cabinets(i).Model & vbCrLf & _
                 "Vendor Code: " & Cabinets(i).VendorCode & vbCrLf & _
                 "Description: " & Left(Cabinets(i).Description, 50) & IIf(Len(Cabinets(i).Description) > 50, "...", "") & vbCrLf & _
                 "Manufacturer ID: " & Cabinets(i).manufacturerID & vbCrLf & _
                 "Price: " & Cabinets(i).Price & vbCrLf & _
                 "Currency ID: " & Cabinets(i).CurrencyID & vbCrLf & _
                 "Relevance: " & Cabinets(i).Relevance & vbCrLf & _
                 "Price Date: " & Cabinets(i).PriceDate & vbCrLf & _
                 "Material ID: " & Cabinets(i).materialID & vbCrLf & _
                 "IP ID: " & Cabinets(i).ipID & vbCrLf & _
                 "height_id: " & Cabinets(i).height_id & vbCrLf & _
                 "width_id: " & Cabinets(i).width_id & vbCrLf & _
                 "depth_id: " & Cabinets(i).depth_id & vbCrLf & vbCrLf
        
        i = i + 1
        rs.MoveNext
    Loop
    
    If MyDebug Then
        ' Display results
        MsgBox result, vbInformation, "Control Cabinets List"
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


Sub FilterControlCabinets()
    ' Check if Cabinets array is not empty
    If UBound(Cabinets) < LBound(Cabinets) Then
        ' If array is empty, exit
        Exit Sub
    End If
    
    ' Initialize FilteredControlCabinets with the same size as Cabinets
    ReDim FilteredControlCabinets(LBound(Cabinets) To UBound(Cabinets))
    
    ' Copy all records from Cabinets to FilteredControlCabinets
    Dim i As Long
    For i = LBound(Cabinets) To UBound(Cabinets)
        FilteredControlCabinets(i) = Cabinets(i)
    Next i
    
    ' Get selected manufacturer from ComboBox
    Dim selectedManufacturer As String
    selectedManufacturer = Form_Box_Postgre_v2r0.ComboBox_Manufacturer.Text
    ApplyManufacturerFilter selectedManufacturer
    
    ' Get selected material from ComboBox and apply filter if needed
    Dim selectedMaterial As String
    selectedMaterial = Form_Box_Postgre_v2r0.ComboBox_Material.Text
    ApplyMaterialFilter selectedMaterial
    
    ' Get selected IP from ComboBox and apply filter if needed
    Dim selectedIP As String
    selectedIP = Form_Box_Postgre_v2r0.ComboBox_IP.Text
    ApplyIPFilter selectedIP
    
    ' Get selected model from ComboBox and apply filter if needed
    Dim selectedModel As String
    selectedModel = Form_Box_Postgre_v2r0.ComboBox_Model.Text
    Apply_Model_Filter selectedModel
    
    ' Get selected name from ComboBox and apply filter if needed
    Dim selectedName As String
    selectedName = Form_Box_Postgre_v2r0.ComboBox_Name.Text
    Apply_Name_Filter selectedName
    
    ' Get selected height from ComboBox and apply filter if needed
    Dim selectedHeight As String
    selectedHeight = Form_Box_Postgre_v2r0.ComboBox_H.Text
    Apply_Height_Filter selectedHeight
    
    ' Get selected width from ComboBox and apply filter if needed
    Dim selectedWidth As String
    selectedWidth = Form_Box_Postgre_v2r0.ComboBox_W.Text
    Apply_Width_Filter selectedWidth
    
    ' Get selected depth from ComboBox and apply filter if needed
    Dim selectedDepth As String
    selectedDepth = Form_Box_Postgre_v2r0.ComboBox_D.Text
    Apply_Depth_Filter selectedDepth
    
    ' Display final record count in LabelNum
    Dim recordCount As Long
    On Error Resume Next
    recordCount = UBound(FilteredControlCabinets) - LBound(FilteredControlCabinets) + 1
    If Err.Number <> 0 Then
        recordCount = 0
        Err.Clear
    End If
    On Error GoTo 0
    Form_Box_Postgre_v2r0.LabelNum.Caption = "Number of records: " & recordCount
End Sub

Private Sub ApplyManufacturerFilter(selectedManufacturer As String)
    ' Find ID of selected manufacturer
    Dim targetManufacturerID As Long
    targetManufacturerID = 0
    Dim j As Long
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
    Dim tempArray() As ControlCabinetRecord
    
    ' Create temporary array for filtered records
    ReDim tempArray(LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets))
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).manufacturerID = targetManufacturerID Then
            tempArray(filteredCount) = FilteredControlCabinets(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredControlCabinets to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredControlCabinets(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredControlCabinets(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array with one element (will be unused)
        ReDim FilteredControlCabinets(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredControlCabinets
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub

Private Sub ApplyMaterialFilter(selectedMaterial As String)
    
    ' Find ID of selected material
    Dim targetMaterialID As Long
    targetMaterialID = 0
    Dim j As Long
    For j = LBound(Materials) To UBound(Materials)
        If Materials(j).Name = selectedMaterial Then
            targetMaterialID = Materials(j).ID
            Exit For
        End If
    Next j
    
    ' If material not found, exit
    If targetMaterialID = 0 Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where materialID matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As ControlCabinetRecord
    
    ' Create temporary array for filtered records
    ReDim tempArray(LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets))
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).materialID = targetMaterialID Then
            tempArray(filteredCount) = FilteredControlCabinets(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredControlCabinets to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredControlCabinets(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredControlCabinets(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array with one element (will be unused)
        ReDim FilteredControlCabinets(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredControlCabinets
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub

Private Sub ApplyIPFilter(selectedIP As String)
    
    ' Find ID of selected IP
    Dim targetIPID As Long
    targetIPID = 0
    Dim j As Long
    For j = LBound(IPs) To UBound(IPs)
        If IPs(j).Name = selectedIP Then
            targetIPID = IPs(j).ID
            Exit For
        End If
    Next j
    
    ' If IP not found, exit
    If targetIPID = 0 Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where ipID matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As ControlCabinetRecord
    
    ' Create temporary array for filtered records
    ReDim tempArray(LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets))
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).ipID = targetIPID Then
            tempArray(filteredCount) = FilteredControlCabinets(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredControlCabinets to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredControlCabinets(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredControlCabinets(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array with one element (will be unused)
        ReDim FilteredControlCabinets(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredControlCabinets
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub

Private Sub Apply_Model_Filter(selectedModel As String)
    ' If "all" is selected or empty, skip filtering
    If selectedModel = "all" Or selectedModel = "" Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where Model matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As ControlCabinetRecord
    
    ' Create temporary array for filtered records
    ReDim tempArray(LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets))
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).Model = selectedModel Then
            tempArray(filteredCount) = FilteredControlCabinets(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredControlCabinets to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredControlCabinets(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredControlCabinets(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array with one element (will be unused)
        ReDim FilteredControlCabinets(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredControlCabinets
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub

Private Sub Apply_Name_Filter(selectedName As String)
    ' If "all" is selected or empty, skip filtering
    If selectedName = "all" Or selectedName = "" Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where Name matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As ControlCabinetRecord
    
    ' Create temporary array for filtered records
    ReDim tempArray(LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets))
    
    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).Name = selectedName Then
            tempArray(filteredCount) = FilteredControlCabinets(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredControlCabinets to the number of found records
    If filteredCount > 0 Then
        ReDim FilteredControlCabinets(0 To filteredCount - 1)
        For i = 0 To filteredCount - 1
            FilteredControlCabinets(i) = tempArray(i)
        Next i
    Else
        ' If nothing found, create empty array with one element (will be unused)
        ReDim FilteredControlCabinets(0 To 0)
        ' Set the array to empty by using Erase
        Erase FilteredControlCabinets
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub

Private Sub Apply_Height_Filter(selectedHeight As String)
    ' If "all" is selected or empty, skip filtering
    If selectedHeight = "all" Or selectedHeight = "" Then
        Exit Sub
    End If
    
    ' Find ID of selected height
    Dim targetHeightID As Long
    targetHeightID = 0
    Dim j As Long
    For j = LBound(Heights) To UBound(Heights)
        If Heights(j).value = selectedHeight Then
            targetHeightID = Heights(j).ID
            Exit For
        End If
    Next j
    
    ' If height not found, exit
    If targetHeightID = 0 Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where height_id matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As ControlCabinetRecord
    
    ' Create temporary array for filtered records
    On Error Resume Next
    ReDim tempArray(LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets))
    If Err.Number <> 0 Then
        ' Array is empty, nothing to filter
        Exit Sub
    End If
    On Error GoTo 0

    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).height_id = targetHeightID Then
            tempArray(filteredCount) = FilteredControlCabinets(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredControlCabinets to the number of found records
    If filteredCount > 0 Then
        ReDim Preserve tempArray(0 To filteredCount - 1)
        FilteredControlCabinets = tempArray
    Else
        ' If nothing found, create empty array
        Erase FilteredControlCabinets
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub

Private Sub Apply_Width_Filter(selectedWidth As String)
    ' If "all" is selected or empty, skip filtering
    If selectedWidth = "all" Or selectedWidth = "" Then
        Exit Sub
    End If
    
    ' Find ID of selected width
    Dim targetWidthID As Long
    targetWidthID = 0
    Dim j As Long
    For j = LBound(Widths) To UBound(Widths)
        If Widths(j).value = selectedWidth Then
            targetWidthID = Widths(j).ID
            Exit For
        End If
    Next j
    
    ' If width not found, exit
    If targetWidthID = 0 Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where width_id matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As ControlCabinetRecord
    
    ' Create temporary array for filtered records
    On Error Resume Next
    ReDim tempArray(LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets))
    If Err.Number <> 0 Then
        ' Array is empty, nothing to filter
        Exit Sub
    End If
    On Error GoTo 0

    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).width_id = targetWidthID Then
            tempArray(filteredCount) = FilteredControlCabinets(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredControlCabinets to the number of found records
    If filteredCount > 0 Then
        ReDim Preserve tempArray(0 To filteredCount - 1)
        FilteredControlCabinets = tempArray
    Else
        ' If nothing found, create empty array
        Erase FilteredControlCabinets
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub

Private Sub Apply_Depth_Filter(selectedDepth As String)
    ' If "all" is selected or empty, skip filtering
    If selectedDepth = "all" Or selectedDepth = "" Then
        Exit Sub
    End If
    
    ' Find ID of selected depth
    Dim targetDepthID As Long
    targetDepthID = 0
    Dim j As Long
    For j = LBound(Depths) To UBound(Depths)
        If Depths(j).value = selectedDepth Then
            targetDepthID = Depths(j).ID
            Exit For
        End If
    Next j
    
    ' If depth not found, exit
    If targetDepthID = 0 Then
        Exit Sub
    End If
    
    ' Filter records - keep only those where depth_id matches selected
    Dim filteredCount As Long
    filteredCount = 0
    Dim tempArray() As ControlCabinetRecord
    
    ' Create temporary array for filtered records
    On Error Resume Next
    ReDim tempArray(LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets))
    If Err.Number <> 0 Then
        ' Array is empty, nothing to filter
        Exit Sub
    End If
    On Error GoTo 0

    ' Iterate through all records and copy only matching ones
    Dim i As Long
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).depth_id = targetDepthID Then
            tempArray(filteredCount) = FilteredControlCabinets(i)
            filteredCount = filteredCount + 1
        End If
    Next i
    
    ' Resize FilteredControlCabinets to the number of found records
    If filteredCount > 0 Then
        ReDim Preserve tempArray(0 To filteredCount - 1)
        FilteredControlCabinets = tempArray
    Else
        ' If nothing found, create empty array
        Erase FilteredControlCabinets
    End If
    
    ' Clear temporary array
    Erase tempArray
End Sub









