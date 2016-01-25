
Public Class LHSerialPort
    Inherits System.IO.Ports.SerialPort

    Public outputbuffer(31) As Byte
    Public outputlength As Byte

    Public Shared beginning As Byte = &HA5   '起始符
    Public Shared terminal As Byte = &H16   '结束符

    '命令域
    Public Shared cmdOrdinary As Byte = &H3 '一般请求
    Public Shared cmdIntegral As Byte = &HC '整点数据请求
    Public Shared cmdDistribute As Byte = &HF   '参数下发
    Public Shared cmdStartup As Byte = &H30 '启动
    Public Shared cmdReboot340 As Byte = &HCF '340状态直接重启
    Public Shared cmdReboot As Byte = &HCC '强制重启
    Public Shared cmdTimeModify As Byte = &H3F '对时
    Public Shared cmdLHTimeModify As Byte = &HC0 '老化时间校正
    Public Shared cmdRecord As Byte = &HC3 '招记录信息
    Public Shared cmdNegative As Byte = &H33 '否认应答
    Public Shared cmdNotSame As Byte = &H3C '器件类型不符

    Public Sub New(ByVal portName As String, ByVal baudRate As Integer, _
                   ByVal parity As IO.Ports.Parity, ByVal dataBits As Integer, ByVal stopBits As IO.Ports.StopBits)
        MyBase.new(portName, baudRate, parity, dataBits, stopBits)
    End Sub
    Private Function CS(ByVal input() As Byte, ByVal len As Integer) As Byte
        '计算校验码
        Dim input2(len - 1) As Byte
        For i = 0 To len - 1
            input2(i) = input(i)
        Next

        Dim num As Integer = 0
        For i = 0 To input2.Length - 1
            num += input2(i)
        Next
        Dim re As Byte = num Mod 256
        Return re
    End Function

    '发送轮询帧
    Public Sub WriteOrdinary(ByVal address As Byte)
        Dim wbuffer(6) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 2
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdOrdinary
        wbuffer(5) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(6) = terminal

        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
    End Sub

    '发送启动帧
    Public Sub WriteStartup(ByVal address As Byte)
        Dim wbuffer(6) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 2
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdStartup
        wbuffer(5) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(6) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
    End Sub

    '发送参数下发帧
    Public Sub WriteDistribute(ByVal address As Byte, ByVal prm As prmDistribute)
        Dim wbuffer(28) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 24
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdDistribute
        wbuffer(5) = prm.second
        wbuffer(6) = prm.minute
        wbuffer(7) = prm.hour
        wbuffer(8) = prm.week
        wbuffer(9) = prm.day
        wbuffer(10) = prm.month
        wbuffer(11) = prm.year
        For i = 0 To 11
            wbuffer(12 + i) = prm.pos(i)
        Next
        wbuffer(24) = prm.type
        wbuffer(25) = prm.max
        wbuffer(26) = prm.mini
        wbuffer(27) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(28) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
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
        wbuffer(6) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(7) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
    End Sub

    '发送对时帧
    Public Sub WriteTimeModify(ByVal address As Byte, ByVal second As Byte, ByVal minute As Byte, ByVal hour As Byte)
        Dim wbuffer(9) As Byte

        wbuffer(0) = beginning
        wbuffer(1) = 5
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdTimeModify
        wbuffer(5) = second
        wbuffer(6) = minute
        wbuffer(7) = hour
        wbuffer(8) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(9) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
    End Sub

    '发送老化时间对时帧
    Public Sub WriteLHTimeModify(ByVal address As Byte, ByVal second As Byte, _
                                 ByVal minute As Byte, ByVal lhour As Byte, ByVal hhour As Byte)
        Dim wbuffer(10) As Byte

        wbuffer(0) = beginning
        wbuffer(1) = 6
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdLHTimeModify
        wbuffer(5) = second
        wbuffer(6) = minute
        wbuffer(7) = lhour
        wbuffer(8) = hhour
        wbuffer(9) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(10) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
    End Sub

    '发送招记录信息帧
    Public Sub WriteRecord(ByVal address As Byte)
        Dim wbuffer(6) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 2
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdRecord
        wbuffer(5) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(6) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
    End Sub

    '发送340状态重启帧
    Public Sub WriteReboot340(ByVal address As Byte)
        Dim wbuffer(6) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 2
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdReboot340
        wbuffer(5) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(6) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
    End Sub

    '发送强制重启帧
    Public Sub WriteReboot(ByVal address As Byte)
        Dim wbuffer(6) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 2
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdReboot
        wbuffer(5) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(6) = terminal
        MyBase.Write(wbuffer, 0, wbuffer.Length)

        outputbuffer = wbuffer
        outputlength = wbuffer.Length
    End Sub

    Public Function ReadUp(ByRef buffer() As Byte) As Byte
        For i = 0 To buffer.Length - 1
            buffer(i) = 0
        Next
        Dim len As Byte
        If BytesToRead > 0 Then
            While Not MyBase.ReadByte() = &H63  '查找起始字符
                If BytesToRead = 0 Then Return False
            End While

            len = MyBase.ReadByte()             '地址域+命令域+数据域，比帧长度小5字节
            buffer(0) = &H63
            buffer(1) = len
            MyBase.Read(buffer, 2, len + 3)

            If buffer(len + 4) <> &H16 Then Return &HFF '不是以结束符结尾
            If CS(buffer, len + 3) <> buffer(len + 3) Then Return &HFF '校验错误

            Return buffer(4)
        End If
        Return &HFF '未收到数据
    End Function

End Class
