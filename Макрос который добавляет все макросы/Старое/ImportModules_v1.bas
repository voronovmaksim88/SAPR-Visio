Attribute VB_Name = "ImportModules_v1"
' 2017-10-16 Воронов МВ
' Добавил AutoNakleiki.bas

' 2017-12-12 Воронов МВ
' Добавил AutoPageNum

' 2018-04-20 Воронов МВ
' Добавил Form_Box

' 2018-06-06 Воронов МВ
' Form_KL.frm
' Form_M.frm

' 2018-06-20 Воронов МВ
' обновил AutoNakleiki
' закоментил окно об обновлении чтоб не траить лишние 3 секунды

' 2018-07-31 Воронов МВ
'добавил Form_SA , эот форма для переключателей, её ещё допилить надо будет

Sub ImportAllModules()

Const PATH As String = "D:\SynologyDrive\work_main\Для СХЕМ\VB+Visio\Sapr\"
Const FILES_NUM As Integer = 28

Dim File(1 To FILES_NUM) As String
Dim успешныйИмпорт As Boolean
Dim количествоИмпортированных As Integer

успешныйИмпорт = True
количествоИмпортированных = 0

File(1) = "Autospec_v8.bas"
File(2) = "FormShow.bas"
File(3) = "AutoNum_v2.bas"
File(4) = "Form_E.frm"
File(5) = "Form_HL_v3.frm"
File(6) = "Form_QF_v1.frm"
File(7) = "Form_Spec.frm"
File(8) = "AutoNakleiki_v2.bas"
File(9) = "AutoEskiz.bas"
File(10) = "EditAllPages.bas"
File(11) = "Form_EditAllPages.frm"
File(12) = "AutoPictureSize.bas"
File(13) = "Form_KM.frm"
File(14) = "Form_Area.frm"
File(15) = "AutoPageNum.bas"
File(16) = "Form_CB.frm"
File(17) = "Form_Box_Postgre_v2r0.frm"
File(18) = "Form_M.frm"
File(19) = "Form_SA.frm"
File(20) = "Clear_Prop.bas"
File(21) = "Form_All_Macros.frm"
File(22) = "PostgeSQL_0_test_connection.bas"
File(23) = "PostgeSQL_ControlCabinets.bas"
File(24) = "PostgeSQL_HWD.bas"
File(25) = "PostgeSQL_IP.bas"
File(26) = "PostgeSQL_Manufacturers.bas"
File(27) = "PostgeSQL_Material.bas"
File(28) = "PostgeSQL_GlobalTypes.bas"


Dim oDoc As Visio.Document
Dim VBProj As VBIDE.VBProject
Dim cmpComponents As VBIDE.VBComponents
Dim cmpComp As VBIDE.VBComponent
Dim i As Integer

Set oDoc = ActiveDocument
Set VBProj = ActiveDocument.VBProject
Set cmpComponents = VBProj.VBComponents

' Проверяем наличие всех файлов перед началом импорта
For i = 1 To FILES_NUM
    If Dir(PATH & File(i)) = "" Then
        MsgBox "Отсутствует файл " & File(i), vbExclamation, "Ошибка импорта"
        успешныйИмпорт = False
        Exit Sub
    End If
Next i

' Удаляем существующие компоненты, которые совпадают с нашими файлами импорта
On Error Resume Next
For Each cmpComp In VBProj.VBComponents
    For i = 1 To FILES_NUM
        If (cmpComp.Name = Left(File(i), Len(File(i)) - 4)) Then
            'MsgBox ("Обновлён " & cmpComp.Name)
            cmpComponents.Remove cmpComponents.item(cmpComp.Name)
            Exit For
        End If
    Next i
Next
On Error GoTo 0

' Импортируем все компоненты
On Error Resume Next
For i = 1 To FILES_NUM
    cmpComponents.Import (PATH & File(i))
    If Err.Number = 0 Then
        количествоИмпортированных = количествоИмпортированных + 1
    Else
        успешныйИмпорт = False
        MsgBox "Ошибка при импорте файла " & File(i) & vbCrLf & _
               "Ошибка: " & Err.Description, vbExclamation, "Ошибка импорта"
        Err.Clear
    End If
Next i
On Error GoTo 0

' Отображаем сообщение об успешном импорте, если все прошло без ошибок
If успешныйИмпорт And количествоИмпортированных = FILES_NUM Then
    MsgBox "Все модули успешно импортированы!" & vbCrLf & _
           "Импортировано файлов: " & количествоИмпортированных & " из " & FILES_NUM, _
           vbInformation, "Импорт завершен"
ElseIf количествоИмпортированных > 0 Then
    MsgBox "Импорт завершен с предупреждениями." & vbCrLf & _
           "Импортировано файлов: " & количествоИмпортированных & " из " & FILES_NUM, _
           vbExclamation, "Импорт завершен"
End If

End Sub
