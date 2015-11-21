Imports System.IO.Ports

Public Class LHSerialPort
    Inherits SerialPort

    Public Shared beginning As Byte = &HA5   '起始符
    Public Shared terminal As Byte = &H16   '结束符

    '命令域
    Public Shared cmdOrdinary As Byte = &H3 '一般请求
    Public Shared cmdIntegral As Byte = &HC '整点数据请求
    Public Shared cmdDistribute As Byte = &HF   '参数下发
    Public Shared cmdStartup As Byte = &H30 '启动

    Private Function CS(ByVal input() As Byte) As Byte
        '计算校验码
        Dim input2(input.Length - 1) As Byte
        For i = 0 To input.Length - 1
            input2(i) = input(i)
        Next

        Dim num As Byte = 0
        For i = 0 To input2.Length - 1 - 2
            While (input2(i))
                input2(i) = input2(i) And (input2(i) - 1)
                num += 1
            End While
        Next
        Return num
    End Function

    '发送一般请求帧
    Public Sub WriteOrdinary(ByVal address As Byte)
        Dim wbuffer(6) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 2
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdOrdinary
        wbuffer(5) = CS(wbuffer)
        wbuffer(6) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)
    End Sub

    '发送启动帧
    Public Sub WriteStartup(ByVal address As Byte)
        Dim wbuffer(6) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 2
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdStartup
        wbuffer(5) = CS(wbuffer)
        wbuffer(6) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)
    End Sub

    '发送参数下发帧
    Public Sub WriteDistribute(ByVal address As Byte, ByVal prm As prmDistribute)
        Dim wbuffer(18) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 14
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdDistribute
        wbuffer(5) = prm.second
        wbuffer(6) = prm.minute
        wbuffer(7) = prm.hour
        For i = 0 To 5
            wbuffer(8 + i) = prm.pos(i)
        Next
        wbuffer(14) = prm.type
        wbuffer(15) = prm.max
        wbuffer(16) = prm.mini
        wbuffer(17) = CS(wbuffer)
        wbuffer(18) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)
    End Sub

    '发送整点数据请求帧
    Public Sub WriteIntegral(ByVal address As Byte, ByVal part As Byte, ByVal time As Byte)
        Dim wbuffer(7) As Byte

        wbuffer(0) = beginning
        wbuffer(1) = 3
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdIntegral
        wbuffer(5) = (part << 6) + time
        wbuffer(6) = CS(wbuffer)
        wbuffer(7) = terminal

        MyBase.Write(wbuffer, 0, wbuffer.Length)
    End Sub
End Class
