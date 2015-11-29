﻿Public Class DownloadCmd

    Public Shared Sub Polling(ByVal com As LHSerialPort, ByVal unit As LHUnit)
        '轮询命令
        If Not unit.isTesting Then
            Exit Sub    '跳过不在运行的单元
        Else
            com.WriteOrdinary(unit.address)
        End If
    End Sub

    Public Shared Sub Integral(ByVal com As LHSerialPort, ByVal unit As LHUnit, _
                               ByVal part As Byte, ByVal time As Byte)
        '整点召回
        com.WriteIntegral(unit.address, part, time)
    End Sub

    Public Shared Sub Distribute(ByVal com As LHSerialPort, ByVal unit As LHUnit)
        '参数下发
        Dim param As New prmDistribute
        param.type = (unit.器件类型 << 4) + unit.电压规格
        Select Case unit.电压规格
            Case 0  '21V
                param.max = &H8     '+0.40V
                param.mini = &H88   '-0.40V
            Case 1  '25V
                param.max = &HA     '+0.50V
                param.mini = &H8A   '-0.50V
            Case 2  '28V
                param.max = &HB     '+0.55V
                param.mini = &H8B   '-0.55V
        End Select
        param.hour = Num2BCD(Now.Hour)
        param.minute = Num2BCD(Now.Minute)
        param.second = Num2BCD(Now.Second)
        com.WriteDistribute(unit.address, param)
    End Sub

    Public Shared Sub Startup(ByVal com As LHSerialPort, ByVal unit As LHUnit)
        '启动单元
        If unit.isTesting Then
            If MsgBox("是否结束现有试验重新开始？", MsgBoxStyle.YesNo, "确认") = vbNo Then
                Exit Sub
            End If
        End If
        com.WriteStartup(unit.address)
    End Sub

    Public Shared Sub Synchronous()
        '对时——等待协议进一步完善
    End Sub

    Public Shared Function Num2BCD(ByVal num As Byte) As Byte
        '数字转换为8421BCD码
        Return (num Mod 10) + ((num \ 10) << 4)
    End Function

    Public Shared Function BCD2Num(ByVal bcd As Byte) As Byte
        '8421BCD码转换为数字
        Return (bcd And &HF) + ((bcd >> 4) * 10)
    End Function
End Class
