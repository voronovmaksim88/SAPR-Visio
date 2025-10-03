Attribute VB_Name = "PostgeSQL_0_test_connection"
Option Explicit
' Option Explicit в VBA (Visual Basic for Applications) Ч это директива,
' котора€ заставл€ет разработчика €вно объ€вл€ть все переменные перед их использованием.
' ≈сли эта строка присутствует в начале модул€, VBA выдаст ошибку компил€ции,
' если вы попытаетесь использовать переменную,
' котора€ не была предварительно объ€влена с помощью Dim, Public, Private или другого ключевого слова объ€влени€.


' Connection string constant
Const CONNECTION_STRING As String = "DSN=PostgreSQL_Vizio_x32;Uid=kis3admin;Pwd=kis3admin1313#;"

Sub ConnectToPostgreSQL()
    ' Declare connection object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connection string using DSN
    ' Replace with your actual credentials
    conn.Open CONNECTION_STRING
    
    ' Declare recordset object to store query results
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Execute SQL query to get city names
    ' Using only the 'name' column for efficiency
    rs.Open "SELECT name FROM cities", conn
    
    ' Check if query returned any data
    If rs.EOF Then
        ' Show warning if table is empty
        MsgBox "The cities table is empty or not found!", vbExclamation, "Error"
    Else
        ' Initialize result string with header
        Dim result As String
        result = "List of cities:" & vbCrLf & vbCrLf ' Header with line breaks
        
        ' Loop through all records
        Do Until rs.EOF
            ' Append each city name with line break
            result = result & rs.Fields("name").value & vbCrLf
            rs.MoveNext ' Move to next record
        Loop
        
        ' Display results in message box
        MsgBox result, vbInformation, "City List"
    End If
    
    ' Clean up objects
    rs.Close
    conn.Close
    
    ' Release memory
    Set rs = Nothing
    Set conn = Nothing
End Sub


