Attribute VB_Name = "PostgeSQL_GlobalTypes"
' ===================================
' Module: GlobalTypes
' Purpose: Common user-defined types (UDT)
' ===================================

' Define user type for storing one record from ControlCabinets table
' (including inherited fields from Equipment)
Public MyDebug As Boolean

Public Type ControlCabinetRecord
    ID As Long
    Name As String
    Model As String
    VendorCode As String
    Description As String
    manufacturerID As Long
    Price As Long
    CurrencyID As Long
    Relevance As Boolean
    PriceDate As Date
    materialID As Long
    ipID As Long
    height_id As Long
    width_id As Long
    depth_id As Long
End Type

' Define user type for storing one record from Manufacturers table
Public Type ManufacturerRecord
    ID As Long          ' id (primary key)
    Name As String      ' name from manufacturers
End Type

' Define user type for storing one record from ControlCabinetMaterial table
Public Type MaterialRecord
    ID As Long          ' id (primary key)
    Name As String      ' name from control_cabinet_materials
End Type

' Define user type for storing one record from IP table
Public Type IpRecord
    ID As Long          ' id (primary key)
    Name As String      ' name from IP
End Type

' Define user type for storing one record from control_cabinet_heights table
Public Type HeightRecord
    ID As Long          ' id (primary key)
    value As String      ' name from control_cabinet_heights
End Type

' Define user type for storing one record from control_cabinet_widths table
Public Type WidthRecord
    ID As Long          ' id (primary key)
    value As String      ' name from control_cabinet_widths
End Type

' Define user type for storing one record from control_cabinet_depths table
Public Type DepthRecord
    ID As Long          ' id (primary key)
    value As String      ' name from control_cabinet_depths
End Type






' For Sensors
' Define user type for storing one record from Manufacturers table
Public Type SensorsManufacturerRecord
    ID As Long          ' id (primary key)
    Name As String      ' name from manufacturers
End Type

Public Type SensorsType
    ID As Long          ' id (primary key)
    Name As String      ' name from manufacturers
End Type

Public Type SensorMeasuredValue
    ID As Long          ' id (primary key)
    Name As String
End Type

Public Type SensorRecord
    ID As Long
    Name As String
    Model As String
    VendorCode As String
    Description As String
    manufacturerID As Long
    Price As Long
    CurrencyID As Long
    Relevance As Boolean
    PriceDate As Date
    SensorTypes() As Long ' Array of Sensor Type IDs (Long)
    SensorMeasuredValues() As Long ' Array of Measured Values IDs (Long)
End Type





