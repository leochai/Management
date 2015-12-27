Imports System.Data.OleDb

Public Class DBMethord
    Public Shared Sub Initial(ByRef db As OleDbConnection, ByRef unit() As LHUnit)
        Dim dbcom As New OleDbCommand
        dbcom.Connection = db
        dbcom.CommandText = "select * from 单元状态 left join 试验记录 on 单元状态.试验编号 = 试验记录.试验编号"
        Dim reader As OleDbDataReader = dbcom.ExecuteReader()
        If reader.HasRows Then
            While reader.Read
                With unit(reader("老化单元号") - 1)
                    If Not reader("器件类型") Is DBNull.Value Then
                        .器件类型 = reader("器件类型")
                    End If
                    If Not reader("运行状态") Is DBNull.Value Then
                        .isTesting = reader("运行状态")
                    End If
                    If Not reader("单元地址") Is DBNull.Value Then
                        .address = reader("单元地址")
                    End If
                    If Not reader("最近上传时间") Is DBNull.Value Then
                        .lastHour = reader("最近上传时间")
                    End If
                    If Not reader("座子类型") Is DBNull.Value Then
                        .座子类型 = reader("座子类型")
                    End If
                    If Not reader("电压规格") Is DBNull.Value Then
                        .电压规格 = reader("电压规格")
                    End If
                    If Not reader("质量等级") Is DBNull.Value Then
                        .质量等级 = reader("质量等级")
                    End If
                    If Not reader("生产批号") Is DBNull.Value Then
                        .生产批号 = reader("生产批号")
                    End If
                    If Not reader("标准号") Is DBNull.Value Then
                        .标准号 = reader("标准号")
                    End If
                    If Not reader("单元状态.试验编号") Is DBNull.Value Then
                        .试验编号 = reader("单元状态.试验编号")
                    End If
                    If Not reader("开机日期") Is DBNull.Value Then
                        .开机日期 = reader("开机日期")
                    End If
                    If Not reader("操作员") Is DBNull.Value Then
                        .操作员 = reader("操作员")
                    End If
                End With
            End While
        End If
        For i = 0 To 23
            dbcom.CommandText = "select 老化位号,器件编号 from 对位表 where 老化单元号 = " & i + 1
            reader.Close()
            reader = dbcom.ExecuteReader()
            While reader.Read
                unit(i).对位表(reader("老化位号") - 1) = reader("器件编号")
            End While
        Next
    End Sub

    Public Shared Sub WriteResult(ByVal testnum As String, ByVal chipnum As Byte, ByVal num As Byte, _
                                  ByVal hour As Byte, ByVal volt As Single, ByVal power As Single)
        Dim dbcmd As New OleDbCommand
        dbcmd.Connection = _DBconn

        Dim rdate As Date = Now
        rdate = rdate.AddHours(hour - rdate.Hour)
        If hour > Now.Hour Then
            rdate = rdate.AddDays(-1)
        End If

        dbcmd.CommandText = "insert into 试验结果 values('" _
                            & testnum & "','" _
                            & chipnum & "','" _
                            & num & "','" _
                            & rdate & "','" _
                            & volt & "','" _
                            & power & "')"
        dbcmd.ExecuteNonQuery()
    End Sub
End Class
