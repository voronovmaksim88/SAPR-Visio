VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Sensors_PostgreSQL 
   Caption         =   "Choice sensor"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7140
   OleObjectBlob   =   "Form_Sensors_PostgreSQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Sensors_PostgreSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CheckBoxCheckBox_SensorType_Interfeis_Click()

End Sub



Private Sub ComboBox_Manufacturer_Change()
    FilterSensors
    Fill_ComboBox_SensorType
    Fill_ComboBox_MeasuredValue
    Fill_ComboBox_Model
    Fill_ComboBox_Name
End Sub

Private Sub CommandButton_Cancel_Click()
    Unload Form_Sensors_PostgreSQL
End Sub




Private Sub Label5_Click()

End Sub

Private Sub UserForm_Initialize()
    MyDebug = True
    
    ' Load data from Sensors table into array when form starts
    ReadManufacturers
    ReadSensorMeasuredValue
    ReadSensorsType
    ReadSensors
    
    FilterSensors
    
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_SensorType
    Fill_Label_ShapeNum
    Fill_ComboBox_MeasuredValue
    Fill_ComboBox_Model
    Fill_ComboBox_Name
    
    
    

    
End Sub

Private Sub Fill_ComboBox_Manufacturer()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_Manufacturer.ListIndex >= 0 Then
        currentValue = ComboBox_Manufacturer.Text
    End If
    
    ' Declare variables
    Dim i As Long
    Dim j As Long
    Dim manufacturerExists As Boolean
    Dim uniqueManufacturers As Collection
    Set uniqueManufacturers = New Collection
    
    On Error Resume Next ' Handle potential collection key conflicts
    
    ' Loop through Sensors array to collect unique manufacturer IDs
    For i = LBound(FilteredSensors) To UBound(FilteredSensors)
        ' Only add if manufacturer_id is not 0 (non-Null)
        If FilteredSensors(i).manufacturerID <> 0 Then
            uniqueManufacturers.Add FilteredSensors(i).manufacturerID, CStr(FilteredSensors(i).manufacturerID)
        End If
    Next i
    
    On Error GoTo ErrorHandler
    
    ' Clear existing items in ComboBox
    ComboBox_Manufacturer.Clear
    
    ' Add "Any" as the first option
    ComboBox_Manufacturer.AddItem "all"
    
    ' Loop through unique manufacturer IDs and match with Manufacturers array
    For i = 1 To uniqueManufacturers.Count
        ' Find matching manufacturer name
        For j = LBound(Manufacturers) To UBound(Manufacturers)
            If Manufacturers(j).ID = uniqueManufacturers(i) Then
                ' Add manufacturer name to ComboBox
                ComboBox_Manufacturer.AddItem Manufacturers(j).Name
                Exit For
            End If
        Next j
    Next i
    
    ' Try to restore previous selection
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_Manufacturer.ListCount - 1
            If ComboBox_Manufacturer.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If
    
    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_Manufacturer.ListIndex = foundIndex
    Else
        ComboBox_Manufacturer.ListIndex = 0  ' "all"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while filling the manufacturer ComboBox: " & Err.Description, vbCritical, "Error"
    Set uniqueManufacturers = Nothing
End Sub




Private Sub Fill_ComboBox_SensorType()

End Sub



Private Sub Fill_Label_ShapeNum()
' Safely set Label_ShapeNum.Caption if selection and user cells exist
    Dim sel As Object
    Set sel = Nothing
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    On Error GoTo 0
    
    Dim shapeType As String
    Dim shapeNum As String
    shapeType = ""
    shapeNum = ""
    
    If Not sel Is Nothing Then
        On Error Resume Next
        shapeType = sel.PrimaryItem.Cells("User.ShapeType").ResultStr("")
        shapeNum = CStr(CLng(sel.PrimaryItem.Cells("User.ShapeNum").result("")))
        On Error GoTo 0
    End If
    
    Dim result As String
    If shapeType <> "" And shapeNum <> "" Then
        result = shapeType & shapeNum
        Label_ShapeNum.Caption = result
    Else
        Label_ShapeNum.Caption = ""
    End If

End Sub


Private Sub Fill_ComboBox_MeasuredValue()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_MeasuredValue.ListIndex >= 0 Then
        currentValue = ComboBox_MeasuredValue.Text
    End If
    
    ' Clear existing items in ComboBox
    ComboBox_MeasuredValue.Clear
    
    ' Add "all" as the first option
    ComboBox_MeasuredValue.AddItem "all"
    
    ' Add measured values from SensorMeasuredValues array
    Dim i As Long
    For i = LBound(SensorMeasuredValues) To UBound(SensorMeasuredValues)
        ComboBox_MeasuredValue.AddItem SensorMeasuredValues(i).Name
    Next i
    
    ' Try to restore previous selection
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_MeasuredValue.ListCount - 1
            If ComboBox_MeasuredValue.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If
    
    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_MeasuredValue.ListIndex = foundIndex
    Else
        ComboBox_MeasuredValue.ListIndex = 0  ' "all"
    End If
End Sub



Private Sub Fill_ComboBox_Model()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_Model.ListIndex >= 0 Then
        currentValue = ComboBox_Model.Text
    End If
    
    ' Collect unique models from FilteredSensors array
    Dim i As Long, j As Long
    Dim uniqueModels As Collection
    Set uniqueModels = New Collection
    
    On Error Resume Next ' Ignore error when adding duplicates
    For i = LBound(FilteredSensors) To UBound(FilteredSensors)
        If FilteredSensors(i).Model <> "" Then
            uniqueModels.Add FilteredSensors(i).Model, FilteredSensors(i).Model
        End If
    Next i
    On Error GoTo 0
    
    ' Convert collection to array and sort
    Dim modelsArray() As String
    If uniqueModels.Count > 0 Then
        ReDim modelsArray(1 To uniqueModels.Count)
        For i = 1 To uniqueModels.Count
            modelsArray(i) = uniqueModels(i)
        Next i
        ' Bubble sort
        Dim temp As String
        For i = 1 To UBound(modelsArray) - 1
            For j = i + 1 To UBound(modelsArray)
                If modelsArray(i) > modelsArray(j) Then
                    temp = modelsArray(i)
                    modelsArray(i) = modelsArray(j)
                    modelsArray(j) = temp
                End If
            Next j
        Next i
    End If
    
    ' Clear ComboBox
    ComboBox_Model.Clear
    ' Add "all" as the first item
    ComboBox_Model.AddItem "all"
    ' Add sorted unique models
    If uniqueModels.Count > 0 Then
        For i = 1 To UBound(modelsArray)
            ComboBox_Model.AddItem modelsArray(i)
        Next i
    End If
    ' Restore previous selection if possible
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_Model.ListCount - 1
            If ComboBox_Model.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If
    If foundIndex >= 0 Then
        ComboBox_Model.ListIndex = foundIndex
    Else
        ComboBox_Model.ListIndex = 0  ' "all"
    End If
End Sub




Private Sub Fill_ComboBox_Name()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_Name.ListIndex >= 0 Then
        currentValue = ComboBox_Name.Text
    End If
    
    ' Collect unique names from FilteredSensors array
    Dim i As Long, j As Long
    Dim uniqueNames As Collection
    Set uniqueNames = New Collection
    
    On Error Resume Next ' Ignore error when adding duplicates
    For i = LBound(FilteredSensors) To UBound(FilteredSensors)
        If FilteredSensors(i).Name <> "" Then
            uniqueNames.Add FilteredSensors(i).Name, FilteredSensors(i).Name
        End If
    Next i
    On Error GoTo 0
    
    ' Convert collection to array and sort
    Dim namesArray() As String
    If uniqueNames.Count > 0 Then
        ReDim namesArray(1 To uniqueNames.Count)
        For i = 1 To uniqueNames.Count
            namesArray(i) = uniqueNames(i)
        Next i
        ' Bubble sort
        Dim temp As String
        For i = 1 To UBound(namesArray) - 1
            For j = i + 1 To UBound(namesArray)
                If namesArray(i) > namesArray(j) Then
                    temp = namesArray(i)
                    namesArray(i) = namesArray(j)
                    namesArray(j) = temp
                End If
            Next j
        Next i
    End If
    
    ' Clear ComboBox
    ComboBox_Name.Clear
    ' Add "all" as the first item
    ComboBox_Name.AddItem "all"
    ' Add sorted unique names
    If uniqueNames.Count > 0 Then
        For i = 1 To UBound(namesArray)
            ComboBox_Name.AddItem namesArray(i)
        Next i
    End If
    ' Restore previous selection if possible
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_Name.ListCount - 1
            If ComboBox_Name.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If
    If foundIndex >= 0 Then
        ComboBox_Name.ListIndex = foundIndex
    Else
        ComboBox_Name.ListIndex = 0  ' "all"
    End If
End Sub
    
