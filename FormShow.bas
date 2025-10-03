Attribute VB_Name = "FormShow"
' Воронов МВ
' добавил KL это клеммы

Public Manufacturer As String

Sub QF(shpObj As Visio.Shape)
    Form_QF_v1.Show
End Sub

Sub HL(shpObj As Visio.Shape)
    Form_HL_v3.Show
End Sub

Sub KM(shpObj As Visio.Shape)
    Form_KM.Show
End Sub

Sub E(shpObj As Visio.Shape)
    Form_E.Show
End Sub

Sub Area(shpObj As Visio.Shape)
    Form_Area.Show
End Sub

Sub CB(shpObj As Visio.Shape)
    Form_CB.Show
End Sub

Sub Box(shpObj As Visio.Shape)
    Form_Box_Postgre_v2r0.Show
End Sub

Sub KL(shpObj As Visio.Shape)
    Form_KL.Show
End Sub


Sub M(shpObj As Visio.Shape)
    Form_M.Show
End Sub

Sub SA(shpObj As Visio.Shape)
    Form_SA.Show
End Sub

Sub Show_Form_All_Macros()
    Form_All_Macros.Show
End Sub

Sub Show_Form_Sensors_PostgreSQL(shpObj As Visio.Shape)
    Form_Sensors_PostgreSQL.Show
End Sub

