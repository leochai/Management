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
                .lastHour = reader("最近上传时间")
                .座子类型 = reader("座子类型")
                .电压规格 = reader("电压规格")
                '.质量等级 = reader("质量等级")
                '.生产批号 = reader("生产批号")
                '.标准号 = reader("标准号")
                '.试验编号 = reader("试验编号")
                '.开机日期 = reader("开机日期")
                '-----待补充
            End With
        End While

        For i = 0 To 23
            dbcom.CommandText = "select 老化位号,器件编号 from 对位表 where 老化单元号 = " & i + 1
            reader.Close()
            reader = dbcom.ExecuteReader()
            While reader.Read
                unit(i).对位表(reader("老化位号") - 1) = reader("器件编号")
            End While
        Next
    End Sub

    Public Shared Sub WriteResult()

    End Sub
End Class
