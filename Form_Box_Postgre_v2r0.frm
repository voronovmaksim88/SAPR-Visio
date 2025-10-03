VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Box_Postgre_v2r0 
   Caption         =   "Form_Box"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   5895
   OleObjectBlob   =   "Form_Box_Postgre_v2r0.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Box_Postgre_v2r0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private isUpdatingComboBoxes As Boolean

Private Sub ComboBox_D_Change()
    If isUpdatingComboBoxes Then Exit Sub
    isUpdatingComboBoxes = True
    FilterControlCabinets

    ' Fill comboboxes
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    Fill_ComboBox_IP
    Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    'Fill_ComboBox_Depths
    Fill_ComboBox_Name
    Fill_ComboBox_Model

    ' If user selected "all", refill this ComboBox with all possible values first
    If ComboBox_D.Text = "all" Then
        Fill_ComboBox_Depths
    End If
    ' Fill description if only one record remains after filtering
    FillDescriptionIfSingleRecord
    
    ' Set single options after updating is complete
    SetSingleOptionIfAvailable
End Sub

Private Sub ComboBox_H_Change()
    If isUpdatingComboBoxes Then Exit Sub
    isUpdatingComboBoxes = True
    FilterControlCabinets

    ' Fill comboboxes
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    Fill_ComboBox_IP
    'Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    Fill_ComboBox_Name
    Fill_ComboBox_Model

    ' If user selected "all", refill this ComboBox with all possible values first
    If ComboBox_H.Text = "all" Then
        Fill_ComboBox_Heights
    End If
    ' Fill description if only one record remains after filtering
    FillDescriptionIfSingleRecord
    
    ' Set single options after updating is complete
    SetSingleOptionIfAvailable
End Sub

Private Sub ComboBox_IP_Change()
    If isUpdatingComboBoxes Then Exit Sub
    isUpdatingComboBoxes = True
    FilterControlCabinets
    
    ' Fill comboboxes
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    'Fill_ComboBox_IP
    Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    Fill_ComboBox_Name
    Fill_ComboBox_Model

    ' If user selected "all", refill this ComboBox with all possible values first
    If ComboBox_IP.Text = "all" Then
        Fill_ComboBox_IP
    End If
    isUpdatingComboBoxes = False
    
    ' Fill description if only one record remains after filtering
    FillDescriptionIfSingleRecord
    
    ' Set single options after updating is complete
    SetSingleOptionIfAvailable
End Sub

Private Sub ComboBox_Manufacturer_Change()
    If isUpdatingComboBoxes Then Exit Sub
    isUpdatingComboBoxes = True
    FilterControlCabinets
    
    ' Fill comboboxes
    'Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    Fill_ComboBox_IP
    Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    Fill_ComboBox_Name
    Fill_ComboBox_Model
    
    ' If user selected "all", refill this ComboBox with all possible values first
    If ComboBox_Manufacturer.Text = "all" Then
        Fill_ComboBox_Manufacturer
    End If
    isUpdatingComboBoxes = False
    
    ' Fill description if only one record remains after filtering
    FillDescriptionIfSingleRecord
    
    ' Set single options after updating is complete
    SetSingleOptionIfAvailable
End Sub



Private Sub ComboBox_Name_Change()
    If isUpdatingComboBoxes Then Exit Sub
    isUpdatingComboBoxes = True
    FilterControlCabinets

    ' Fill comboboxes
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    Fill_ComboBox_IP
    Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    'Fill_ComboBox_Name
    Fill_ComboBox_Model

    ' If user selected "all", refill this ComboBox with all possible values first
    If ComboBox_Name.Text = "all" Then
        Fill_ComboBox_Name
    End If
    isUpdatingComboBoxes = False
    
    ' Fill description if only one record remains after filtering
    FillDescriptionIfSingleRecord
    
    ' Set single options after updating is complete
    SetSingleOptionIfAvailable
End Sub

Private Sub ComboBox_W_Change()
    If isUpdatingComboBoxes Then Exit Sub
    isUpdatingComboBoxes = True
    FilterControlCabinets

    ' Fill comboboxes
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    Fill_ComboBox_IP
    Fill_ComboBox_Heights
    'Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    Fill_ComboBox_Name
    Fill_ComboBox_Model

    ' If user selected "all", refill this ComboBox with all possible values first
    If ComboBox_W.Text = "all" Then
        Fill_ComboBox_Widths
    End If
    isUpdatingComboBoxes = False
    
    ' Fill description if only one record remains after filtering
    FillDescriptionIfSingleRecord
    
    ' Set single options after updating is complete
    SetSingleOptionIfAvailable
End Sub

Private Sub ComboBox_Material_Change()
    If isUpdatingComboBoxes Then Exit Sub
    isUpdatingComboBoxes = True
    FilterControlCabinets

    ' Fill comboboxes
    Fill_ComboBox_Manufacturer
    'Fill_ComboBox_Material
    Fill_ComboBox_IP
    Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    Fill_ComboBox_Name
    Fill_ComboBox_Model

    ' If user selected "all", refill this ComboBox with all possible values first
    If ComboBox_Material.Text = "all" Then
        Fill_ComboBox_Material
    End If
    isUpdatingComboBoxes = False
    
    ' Fill description if only one record remains after filtering
    FillDescriptionIfSingleRecord
    
    ' Set single options after updating is complete
    SetSingleOptionIfAvailable
End Sub



Private Sub ComboBox_Model_Change()
    If isUpdatingComboBoxes Then Exit Sub
    isUpdatingComboBoxes = True
    FilterControlCabinets
    
    
    ' Fill comboboxes
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    Fill_ComboBox_IP
    Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    Fill_ComboBox_Name
    'Fill_ComboBox_Model
    
    
    ' If user selected "all", refill this ComboBox with all possible values first
    If ComboBox_Model.Text = "all" Then
        Fill_ComboBox_Model
    End If
    isUpdatingComboBoxes = False
    
    ' Fill description if only one record remains after filtering
    FillDescriptionIfSingleRecord
    
    ' Set single options after updating is complete
    SetSingleOptionIfAvailable
End Sub

Private Sub CommandButton_Cancel_Click()
    Unload Form_Box_Postgre_v2r0
End Sub

Private Sub CommandButton_OK_Click()
 
 If (ComboBox_Model.Text <> "all" And ComboBox_Model.Name <> "all") Then

    ActiveWindow.Selection.PrimaryItem.Cells("prop.Manufacturer").FormulaU = Chr(34) + ComboBox_Manufacturer.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Note").FormulaU = Chr(34) + TextBox_Description.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Name").FormulaU = Chr(34) + ComboBox_Name.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.IP").FormulaU = Chr(34) + ComboBox_IP.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Material").FormulaU = Chr(34) + ComboBox_Material.Text + Chr(34)
  
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Height").FormulaU = Chr(34) + ComboBox_H.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Width").FormulaU = Chr(34) + ComboBox_W.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Depth").FormulaU = Chr(34) + ComboBox_D.Text + Chr(34)
    ActiveWindow.Selection.PrimaryItem.Cells("Prop.Model").FormulaU = Chr(34) + ComboBox_Model.Text + Chr(34)
    
    ActiveWindow.Selection.PrimaryItem.Cells("Width").FormulaU = CStr(Round(Val(ComboBox_W.Text) / 4)) & " mm"
    ActiveWindow.Selection.PrimaryItem.Cells("Height").FormulaU = CStr(Round(Val(ComboBox_H.Text) / 4)) & " mm"
    
    Form_Box_Postgre_v2r0.Hide
 End If 'Model ComboBox is Empty
 

End Sub

Private Sub CommandButton_Reset_Click()
    ' Reset all filters and reload data
    isUpdatingComboBoxes = True
    
    ' Reset FilteredControlCabinets to contain all records from Cabinets
    If UBound(Cabinets) >= LBound(Cabinets) Then
        ReDim FilteredControlCabinets(LBound(Cabinets) To UBound(Cabinets))
        Dim i As Long
        For i = LBound(Cabinets) To UBound(Cabinets)
            FilteredControlCabinets(i) = Cabinets(i)
        Next i
    End If
    
    ' Refill all ComboBoxes with all available options
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    Fill_ComboBox_IP
    Fill_ComboBox_Model
    Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    Fill_ComboBox_Name
    
    ' Reset all ComboBoxes to "all"
    ComboBox_Manufacturer.ListIndex = 0  ' "all"
    ComboBox_Material.ListIndex = 0       ' "all"
    ComboBox_IP.ListIndex = 0            ' "all"
    ComboBox_Model.ListIndex = 0         ' "all"
    ComboBox_H.ListIndex = 0         ' "all"
    ComboBox_W.ListIndex = 0         ' "all"
    ComboBox_D.ListIndex = 0         ' "all"
    ComboBox_Name.ListIndex = 0         ' "all"
    
    
    ' Clear description when resetting
    TextBox_Description.Text = ""
    
    isUpdatingComboBoxes = False
    
    ' Update record count display
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

Private Sub UserForm_Initialize()
    ' Load data from ControlCabinets table into array when form starts
    ReadControlCabinets
    ReadManufacturers
    ReadMaterials
    ReadIPs
    ReadHeights
    ReadWidths
    ReadDepths
    
    
    FilterControlCabinets
        
    ' Fill comboboxes
    Fill_ComboBox_Manufacturer
    Fill_ComboBox_Material
    Fill_ComboBox_IP
    Fill_ComboBox_Heights
    Fill_ComboBox_Widths
    Fill_ComboBox_Depths
    Fill_ComboBox_Name
    Fill_ComboBox_Model
    
    ' Set single options if available
    SetSingleOptionIfAvailable
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
    
    ' Loop through Cabinets array to collect unique manufacturer IDs
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        ' Only add if manufacturer_id is not 0 (non-Null)
        If Cabinets(i).manufacturerID <> 0 Then
            uniqueManufacturers.Add FilteredControlCabinets(i).manufacturerID, CStr(FilteredControlCabinets(i).manufacturerID)
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

Private Sub Fill_ComboBox_Material()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_Material.ListIndex >= 0 Then
        currentValue = ComboBox_Material.Text
    End If
    
    ' Declare variables
    Dim i As Long
    Dim j As Long
    Dim uniqueMaterials As Collection
    Set uniqueMaterials = New Collection
    
    On Error Resume Next ' Handle potential duplicates when adding to collection
    
    ' Collect unique material_ids from Cabinets array
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).materialID <> 0 Then
            uniqueMaterials.Add FilteredControlCabinets(i).materialID, CStr(FilteredControlCabinets(i).materialID)
        End If
    Next i
    
    On Error GoTo ErrorHandler

    ' Clear combobox
    ComboBox_Material.Clear
    
    ' Add "Any" as the first item
    ComboBox_Material.AddItem "all"
    
    ' Add material names by their IDs
    For i = 1 To uniqueMaterials.Count
        For j = LBound(Materials) To UBound(Materials)
            If Materials(j).ID = uniqueMaterials(i) Then
                ComboBox_Material.AddItem Materials(j).Name
                Exit For
            End If
        Next j
    Next i
    
    ' Try to restore previous selection
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_Material.ListCount - 1
            If ComboBox_Material.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If
    
    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_Material.ListIndex = foundIndex
    Else
        ComboBox_Material.ListIndex = 0  ' "all"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while filling the material ComboBox: " & Err.Description, vbCritical, "Error"
    Set uniqueMaterials = Nothing
End Sub

Private Sub Fill_ComboBox_IP()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_IP.ListIndex >= 0 Then
        currentValue = ComboBox_IP.Text
    End If

    ' Declare variables
    Dim i As Long
    Dim j As Long
    Dim uniqueIPNames As Collection
    Set uniqueIPNames = New Collection

    On Error Resume Next ' Ignore errors when adding duplicates to collection

    ' Collect unique IP names from Cabinets array
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).ipID <> 0 Then
            ' Find the corresponding IP name
            For j = LBound(IPs) To UBound(IPs)
                If IPs(j).ID = FilteredControlCabinets(i).ipID Then
                    ' Add the name to the collection, using name as key for uniqueness
                    uniqueIPNames.Add IPs(j).Name, IPs(j).Name
                    Exit For
                End If
            Next j
        End If
    Next i

    On Error GoTo ErrorHandler

    ' Convert collection to array for sorting
    Dim ipsArray() As String
    If uniqueIPNames.Count > 0 Then
        ReDim ipsArray(1 To uniqueIPNames.Count)
        For i = 1 To uniqueIPNames.Count
            ipsArray(i) = uniqueIPNames(i)
        Next i

        ' Sort array alphabetically using bubble sort
        Dim temp As String
        For i = 1 To UBound(ipsArray) - 1
            For j = i + 1 To UBound(ipsArray)
                If ipsArray(i) > ipsArray(j) Then
                    temp = ipsArray(i)
                    ipsArray(i) = ipsArray(j)
                    ipsArray(j) = temp
                End If
            Next j
        Next i
    End If

    ' Clear combobox
    ComboBox_IP.Clear

    ' Add "all" as the first item
    ComboBox_IP.AddItem "all"

    ' Add sorted unique IP names
    If uniqueIPNames.Count > 0 Then
        For i = 1 To UBound(ipsArray)
            ComboBox_IP.AddItem ipsArray(i)
        Next i
    End If

    ' Try to restore previous selection
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_IP.ListCount - 1
            If ComboBox_IP.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If

    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_IP.ListIndex = foundIndex
    Else
        ComboBox_IP.ListIndex = 0  ' "all"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while filling the IP ComboBox: " & Err.Description, vbCritical, "Error"
    Set uniqueIPNames = Nothing
End Sub




Private Sub Fill_ComboBox_Heights()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_H.ListIndex >= 0 Then
        currentValue = ComboBox_H.Text
    End If

    ' Declare variables
    Dim i As Long
    Dim j As Long
    Dim uniqueHeights As Collection
    Set uniqueHeights = New Collection

    On Error Resume Next ' Ignore errors when adding duplicates to collection

    ' Collect unique height values from Cabinets array
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).height_id <> 0 Then
            ' Find the corresponding height value
            For j = LBound(Heights) To UBound(Heights)
                If Heights(j).ID = FilteredControlCabinets(i).height_id Then
                    ' Add the value to the collection, using value as key for uniqueness
                    uniqueHeights.Add Heights(j).value, CStr(Heights(j).value)
                    Exit For
                End If
            Next j
        End If
    Next i

    On Error GoTo ErrorHandler

    ' Convert collection to array for sorting
    Dim heightsArray() As Variant ' Use Variant to hold strings that are numbers
    If uniqueHeights.Count > 0 Then
        ReDim heightsArray(1 To uniqueHeights.Count)
        For i = 1 To uniqueHeights.Count
            heightsArray(i) = uniqueHeights(i)
        Next i

        ' Sort array numerically using bubble sort
        Dim temp As Variant
        For i = 1 To UBound(heightsArray) - 1
            For j = i + 1 To UBound(heightsArray)
                ' Compare as numbers
                If CLng(heightsArray(i)) > CLng(heightsArray(j)) Then
                    temp = heightsArray(i)
                    heightsArray(i) = heightsArray(j)
                    heightsArray(j) = temp
                End If
            Next j
        Next i
    End If

    ' Clear combobox
    ComboBox_H.Clear

    ' Add "all" as the first item
    ComboBox_H.AddItem "all"

    ' Add sorted unique height values
    If uniqueHeights.Count > 0 Then
        For i = 1 To UBound(heightsArray)
            ComboBox_H.AddItem heightsArray(i)
        Next i
    End If

    ' Try to restore previous selection
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_H.ListCount - 1
            If ComboBox_H.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If

    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_H.ListIndex = foundIndex
    Else
        ComboBox_H.ListIndex = 0  ' "all"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while filling the Heights ComboBox: " & Err.Description, vbCritical, "Error"
    Set uniqueHeights = Nothing
End Sub





Private Sub Fill_ComboBox_Widths()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_W.ListIndex >= 0 Then
        currentValue = ComboBox_W.Text
    End If

    ' Declare variables
    Dim i As Long
    Dim j As Long
    Dim uniqueWidths As Collection
    Set uniqueWidths = New Collection

    On Error Resume Next ' Ignore errors when adding duplicates to collection

    ' Collect unique width values from Cabinets array
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).width_id <> 0 Then
            ' Find the corresponding width value
            For j = LBound(Widths) To UBound(Widths)
                If Widths(j).ID = FilteredControlCabinets(i).width_id Then
                    ' Add the value to the collection, using value as key for uniqueness
                    uniqueWidths.Add Widths(j).value, CStr(Widths(j).value)
                    Exit For
                End If
            Next j
        End If
    Next i

    On Error GoTo ErrorHandler

    ' Convert collection to array for sorting
    Dim widthsArray() As Variant ' Use Variant to hold strings that are numbers
    If uniqueWidths.Count > 0 Then
        ReDim widthsArray(1 To uniqueWidths.Count)
        For i = 1 To uniqueWidths.Count
            widthsArray(i) = uniqueWidths(i)
        Next i

        ' Sort array numerically using bubble sort
        Dim temp As Variant
        For i = 1 To UBound(widthsArray) - 1
            For j = i + 1 To UBound(widthsArray)
                ' Compare as numbers
                If CLng(widthsArray(i)) > CLng(widthsArray(j)) Then
                    temp = widthsArray(i)
                    widthsArray(i) = widthsArray(j)
                    widthsArray(j) = temp
                End If
            Next j
        Next i
    End If

    ' Clear combobox
    ComboBox_W.Clear

    ' Add "all" as the first item
    ComboBox_W.AddItem "all"

    ' Add sorted unique width values
    If uniqueWidths.Count > 0 Then
        For i = 1 To UBound(widthsArray)
            ComboBox_W.AddItem widthsArray(i)
        Next i
    End If

    ' Try to restore previous selection
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_W.ListCount - 1
            If ComboBox_W.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If

    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_W.ListIndex = foundIndex
    Else
        ComboBox_W.ListIndex = 0  ' "all"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while filling the Widths ComboBox: " & Err.Description, vbCritical, "Error"
    Set uniqueWidths = Nothing
End Sub




Private Sub Fill_ComboBox_Depths()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_D.ListIndex >= 0 Then
        currentValue = ComboBox_D.Text
    End If

    ' Declare variables
    Dim i As Long
    Dim j As Long
    Dim uniqueDepths As Collection
    Set uniqueDepths = New Collection

    On Error Resume Next ' Ignore errors when adding duplicates to collection

    ' Collect unique depth values from Cabinets array
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).depth_id <> 0 Then
            ' Find the corresponding depth value
            For j = LBound(Depths) To UBound(Depths)
                If Depths(j).ID = FilteredControlCabinets(i).depth_id Then
                    ' Add the value to the collection, using value as key for uniqueness
                    uniqueDepths.Add Depths(j).value, CStr(Depths(j).value)
                    Exit For
                End If
            Next j
        End If
    Next i

    On Error GoTo ErrorHandler

    ' Convert collection to array for sorting
    Dim depthsArray() As Variant ' Use Variant to hold strings that are numbers
    If uniqueDepths.Count > 0 Then
        ReDim depthsArray(1 To uniqueDepths.Count)
        For i = 1 To uniqueDepths.Count
            depthsArray(i) = uniqueDepths(i)
        Next i

        ' Sort array numerically using bubble sort
        Dim temp As Variant
        For i = 1 To UBound(depthsArray) - 1
            For j = i + 1 To UBound(depthsArray)
                ' Compare as numbers
                If CLng(depthsArray(i)) > CLng(depthsArray(j)) Then
                    temp = depthsArray(i)
                    depthsArray(i) = depthsArray(j)
                    depthsArray(j) = temp
                End If
            Next j
        Next i
    End If

    ' Clear combobox
    ComboBox_D.Clear

    ' Add "all" as the first item
    ComboBox_D.AddItem "all"

    ' Add sorted unique depth values
    If uniqueDepths.Count > 0 Then
        For i = 1 To UBound(depthsArray)
            ComboBox_D.AddItem depthsArray(i)
        Next i
    End If

    ' Try to restore previous selection
    Dim foundIndex As Long
    foundIndex = -1
    If currentValue <> "" Then
        For i = 0 To ComboBox_D.ListCount - 1
            If ComboBox_D.List(i) = currentValue Then
                foundIndex = i
                Exit For
            End If
        Next i
    End If

    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_D.ListIndex = foundIndex
    Else
        ComboBox_D.ListIndex = 0  ' "all"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while filling the Depths ComboBox: " & Err.Description, vbCritical, "Error"
    Set uniqueDepths = Nothing
End Sub




Private Sub Fill_ComboBox_Name()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_Name.ListIndex >= 0 Then
        currentValue = ComboBox_Name.Text
    End If
    
    ' Declare variables
    Dim i As Long, j As Long
    Dim uniqueNames As Collection
    Set uniqueNames = New Collection
    
    On Error Resume Next ' Handle potential duplicates when adding to collection
    
    ' Collect unique names from FilteredControlCabinets array
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).Name <> "" Then
            uniqueNames.Add FilteredControlCabinets(i).Name, FilteredControlCabinets(i).Name
        End If
    Next i
    
    On Error GoTo ErrorHandler

    ' Convert collection to array for sorting
    Dim namesArray() As String
    If uniqueNames.Count > 0 Then
        ReDim namesArray(1 To uniqueNames.Count)
        For i = 1 To uniqueNames.Count
            namesArray(i) = uniqueNames(i)
        Next i
        
        ' Sort array alphabetically using bubble sort
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

    ' Clear combobox
    ComboBox_Name.Clear
    
    ' Add "all" as the first item
    ComboBox_Name.AddItem "all"
    
    ' Add sorted unique name values
    If uniqueNames.Count > 0 Then
        For i = 1 To UBound(namesArray)
            ComboBox_Name.AddItem namesArray(i)
        Next i
    End If
    
    ' Try to restore previous selection
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
    
    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_Name.ListIndex = foundIndex
    Else
        ComboBox_Name.ListIndex = 0  ' "all"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while filling the Name ComboBox: " & Err.Description, vbCritical, "Error"
    Set uniqueNames = Nothing
End Sub



Private Sub Fill_ComboBox_Model()
    ' Save current selected value
    Dim currentValue As String
    currentValue = ""
    If ComboBox_Model.ListIndex >= 0 Then
        currentValue = ComboBox_Model.Text
    End If
    
    ' Declare variables
    Dim i As Long, j As Long
    Dim uniqueModels As Collection
    Set uniqueModels = New Collection
    
    On Error Resume Next ' Handle potential duplicates when adding to collection
    
    ' Collect unique models from FilteredControlCabinets array
    For i = LBound(FilteredControlCabinets) To UBound(FilteredControlCabinets)
        If FilteredControlCabinets(i).Model <> "" Then
            uniqueModels.Add FilteredControlCabinets(i).Model, FilteredControlCabinets(i).Model
        End If
    Next i
    
    On Error GoTo ErrorHandler

    ' Convert collection to array for sorting
    Dim modelsArray() As String
    If uniqueModels.Count > 0 Then
        ReDim modelsArray(1 To uniqueModels.Count)
        For i = 1 To uniqueModels.Count
            modelsArray(i) = uniqueModels(i)
        Next i
        
        ' Sort array alphabetically using bubble sort
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

    ' Clear combobox
    ComboBox_Model.Clear
    
    ' Add "all" as the first item
    ComboBox_Model.AddItem "all"
    
    ' Add sorted unique model names
    If uniqueModels.Count > 0 Then
        For i = 1 To UBound(modelsArray)
            ComboBox_Model.AddItem modelsArray(i)
        Next i
    End If
    
    ' Try to restore previous selection
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
    
    ' Set selection: restore previous if found, otherwise set to "all"
    If foundIndex >= 0 Then
        ComboBox_Model.ListIndex = foundIndex
    Else
        ComboBox_Model.ListIndex = 0  ' "all"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while filling the Model ComboBox: " & Err.Description, vbCritical, "Error"
    Set uniqueModels = Nothing
End Sub

Private Sub FillDescriptionIfSingleRecord()
    ' Check if FilteredControlCabinets contains exactly one record
    Dim recordCount As Long
    On Error Resume Next
    recordCount = UBound(FilteredControlCabinets) - LBound(FilteredControlCabinets) + 1
    If Err.Number <> 0 Then
        recordCount = 0
        Err.Clear
        On Error GoTo 0
        ' Clear description if no records
        TextBox_Description.Text = ""
        Exit Sub
    End If
    On Error GoTo 0
    
    ' If exactly one record remains, fill the description
    If recordCount = 1 Then
        TextBox_Description.Text = FilteredControlCabinets(LBound(FilteredControlCabinets)).Description
    Else
        ' Clear description if multiple or no records
        TextBox_Description.Text = ""
    End If
End Sub

Private Sub SetSingleOptionIfAvailable()
    ' Set single option for each ComboBox if only 2 items exist (all + 1 option)
    isUpdatingComboBoxes = True
    
    ' Check Manufacturer ComboBox
    If ComboBox_Manufacturer.ListCount = 2 And ComboBox_Manufacturer.ListIndex = 0 Then
        ComboBox_Manufacturer.ListIndex = 1
    End If
    
    ' Check Material ComboBox
    If ComboBox_Material.ListCount = 2 And ComboBox_Material.ListIndex = 0 Then
        ComboBox_Material.ListIndex = 1
    End If
    
    ' Check IP ComboBox
    If ComboBox_IP.ListCount = 2 And ComboBox_IP.ListIndex = 0 Then
        ComboBox_IP.ListIndex = 1
    End If
    
    ' Check Model ComboBox
    If ComboBox_Model.ListCount = 2 And ComboBox_Model.ListIndex = 0 Then
        ComboBox_Model.ListIndex = 1
    End If
    
    ' Check Heights ComboBox
    If ComboBox_H.ListCount = 2 And ComboBox_H.ListIndex = 0 Then
        ComboBox_H.ListIndex = 1
    End If
    
    ' Check Widths ComboBox
    If ComboBox_W.ListCount = 2 And ComboBox_W.ListIndex = 0 Then
        ComboBox_W.ListIndex = 1
    End If
    
    ' Check Depths ComboBox
    If ComboBox_D.ListCount = 2 And ComboBox_D.ListIndex = 0 Then
        ComboBox_D.ListIndex = 1
    End If
    
    ' Check Name ComboBox
    If ComboBox_Name.ListCount = 2 And ComboBox_Name.ListIndex = 0 Then
        ComboBox_Name.ListIndex = 1
    End If
    
    isUpdatingComboBoxes = False
End Sub

' Helper function to handle Null values
Private Function Nz(value As Variant, default As Variant) As Variant
    If IsNull(value) Then
        Nz = default
    Else
        Nz = value
    End If
End Function



