
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

        Dim num As Byte = 0
        For i = 0 To input2.Length - 1
            While (input2(i))
                input2(i) = input2(i) And (input2(i) - 1)
                If num = 255 Then
                    num = 0
                Else
                    num += 1
                End If
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
        Dim wbuffer(24) As Byte
        wbuffer(0) = beginning
        wbuffer(1) = 20
        wbuffer(2) = beginning
        wbuffer(3) = address
        wbuffer(4) = cmdDistribute
        wbuffer(5) = prm.second
        wbuffer(6) = prm.minute
        wbuffer(7) = prm.hour
        For i = 0 To 11
            wbuffer(8 + i) = prm.pos(i)
        Next
        wbuffer(20) = prm.type
        wbuffer(21) = prm.max
        wbuffer(22) = prm.mini
        wbuffer(23) = CS(wbuffer, wbuffer.Length - 2)
        wbuffer(24) = terminal
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

    Public Function ReadUp(ByRef buffer() As Byte) As Boolean
        Dim len As Byte
        If BytesToRead > 0 Then
            While Not MyBase.ReadByte() = &H63  '查找起始字符
                If BytesToRead = 0 Then Return False
            End While

            len = MyBase.ReadByte()             '地址域+命令域+数据域，比帧长度小5字节
            buffer(0) = &H63
            buffer(1) = len
            MyBase.Read(buffer, 2, len + 3)

            If buffer(len + 4) <> &H16 Then Return False '不是以结束符结尾
            If CS(buffer, len + 3) <> buffer(len + 3) Then Return False '校验错误

            Return True
        End If
        Return False
    End Function


End Class
