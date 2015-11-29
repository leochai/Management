Imports System.Data.OleDb

Public Class DBMethord
    Public Shared Sub Initial(ByRef db As OleDbConnection, ByRef unit() As LHUnit)
        Dim dbcom As New OleDbCommand
        dbcom.Connection = db
        dbcom.CommandText = "select * from 单元状态 order by 老化单元号"
        Dim reader As OleDbDataReader = dbcom.ExecuteReader()
        While reader.Read
            With unit(reader("老化单元号") - 1)
                .器件类型 = reader("器件类型")
                .isTesting = reader("运行状态")
                .address = reader("单元地址")
                '-----待补充
            End With
        End While
    End Sub

End Class
