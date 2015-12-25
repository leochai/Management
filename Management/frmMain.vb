Imports System.Data.OleDb
Imports System.IO.Ports
Imports System.Threading
Imports LeoControls

Public Class frmMain
    Dim ShowList(23) As ArrayList
    Delegate Sub TextCallback(ByVal cmd() As Byte, ByVal len As Byte)
    Delegate Sub Polling_dg1(ByVal volt As Single, ByVal power As Single, _
                             ByVal unitnum As Byte, ByVal cellnum As Byte)
    Delegate Sub Polling_dg4(ByVal volt As Single, ByVal power As Single, _
                             ByVal unitnum As Byte, ByVal cellnum As Byte, ByVal cellpos As Byte)

    Public RS485 As New LHSerialPort("COM3", 1200, Parity.Odd, 8, 1)

    Dim CommThread As New Thread(AddressOf CommTask)


    Private Sub frmMain_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        '_integralFlag = True
        'ShowList(6).Item(5).setResult(5.5, 6.6)
        'Dim v As Single = 5.5
        'Dim mA As Single = 6.6
        'Dim i As Byte = 5
        'Dim pos As Byte = 50
        ''PollingShow(v, v * mA, i, (pos - 1) \ 4, (pos - 1) Mod 4)
        'PollingShow(v, v * mA, i + 1, (pos - 1) \ 4 + 1, (pos - 1) Mod 4)

        OneSec.Enabled = Not OneSec.Enabled
        'Dim cmd As New OleDbCommand
        'cmd.Connection = _DBconn
        'cmd.CommandText = "delete * from 试验结果"
        'cmd.ExecuteNonQuery()
        'Button1.Text = "ok"
    End Sub '测试

    Private Sub PaintShow()
        For k = 0 To 23
            ShowList(k) = New ArrayList
            If _unit(k).座子类型 Then       '1位器件
                Dim sshow(47) As OneShow
                For j = 0 To 3
                    For i = 0 To 11
                        sshow(j * 12 + i) = New OneShow
                        With sshow(j * 12 + i)
                            .Left = 69 * i + 8
                            .Top = 118 * j + 50
                            .ShowNum = j * 12 + i + 1
                            .Parent = TabControl1.TabPages(k)
                            .Enabled = True
                        End With
                        ShowList(k).Add(sshow(j * 12 + i))
                    Next
                Next
            Else                        '2、4位器件
                Dim sshow(23) As FourShow
                For j = 0 To 3
                    For i = 0 To 5
                        sshow(j * 6 + i) = New FourShow
                        With sshow(j * 6 + i)
                            .Left = 138 * i + 8
                            .Top = 118 * j + 50
                            .ShowNum = j * 6 + i + 1
                            .Parent = TabControl1.TabPages(k)
                            .Enabled = True
                        End With
                        ShowList(k).Add(sshow(j * 6 + i))
                    Next
                Next
            End If
            TabControl1.TabPages(k).Show()
        Next
    End Sub '画界面
    Private Sub ThreadInit()
        _commFlag.polling = False
        _commFlag.startup = False
        _commFlag.integral = False
        _commFlag.unitNo = 0
        CommThread.Start()
    End Sub '线程初始化

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        _DBconn.Close()
        CommThread.Abort()
        RS485.Close()
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '------遮挡
        Dim fs As New frmRoot
        fs.Show()
        fs.Label1.Text = "加载数据..."
        fs.Refresh()

        '------设置串口
        RS485.WriteBufferSize = 2048
        RS485.ReadTimeout = 200
        RS485.ReceivedBytesThreshold = 3
        RS485.RtsEnable = True
        Try
            RS485.Open()
        Catch ex As Exception
            MsgBox("请检查串口连接后再打开程序！")
            Me.Close()
        End Try

        '------打开数据库，初始化单元对象
        _DBconn.Open()
        For i = 0 To 23
            _unit(i) = New LHUnit
        Next
        DBMethord.Initial(_DBconn, _unit)

        PaintShow()     '绘制界面
        ThreadInit()    '线程初始化
        OneSec.Enabled = True
        'OneMin.Enabled = True
        fs.Close()
    End Sub


    Private Sub 操作员管理ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 操作员管理ToolStripMenuItem.Click
        frmLogin.ShowDialog()
    End Sub

    Private Sub 试验数据检索ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 试验数据检索ToolStripMenuItem.Click
        frm数据检索.ShowDialog()
    End Sub

    Public Sub PollingShow1(ByVal volt As Single, ByVal power As Single, _
                           ByVal unitnum As Byte, ByVal cellnum As Byte)
        ShowList(unitnum - 1).Item(cellnum - 1).setResult(volt, power)
    End Sub '更新实时数据显示1位

    Public Sub PollingShow4(ByVal volt As Single, ByVal power As Single, _
                           ByVal unitnum As Byte, ByVal cellnum As Byte, ByVal cellpos As Byte)
        ShowList(unitnum - 1).Item(cellnum - 1).setResult(volt, power, cellpos)
    End Sub '更新实时数据显示4位

    Private Sub btnMin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMin.Click
        TabControl1.Hide()
        GroupBox1.Hide()
        PictureBox1.Hide()
        MenuStrip1.Hide()
        btnMin.Hide()
        TextBox1.Hide()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.Size = New Point(280, 720)
        Me.Left = 1000
        Me.Top = 0

        btnMax.Top = 350
        btnMax.Left = 10
        btnMax.Show()
    End Sub '切换为边栏显示

    Private Sub btnMax_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMax.Click
        Me.Size = New Point(1280, 720)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        Me.Left = 0
        Me.Top = 0
        btnMin.Show()
        TextBox1.Show()
        GroupBox1.Show()
        PictureBox1.Show()
        MenuStrip1.Show()
        TabControl1.Show()

        btnMax.Hide()
    End Sub '切换为全屏显示

    Private Sub OneSec_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OneSec.Tick
        Static i As Byte = 0
        While _unit(i).isTesting = False
            i = (i + 1) Mod 24
        End While
        _commFlag.unitNo = i
        _commFlag.polling = True
        i = (i + 1) Mod 24
    End Sub

    Private Sub OneMin_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OneMin.Tick
        Select Case Now.Minute
            Case 5
                OneSec.Enabled = False
                GroupBox1.Enabled = False
                _commFlag.integral = True
            Case 54

            Case 55

            Case 56
        End Select
    End Sub

    Private Sub CommTask()
        While True
            If _commFlag.integral Then
                IntegralTask()
                _commFlag.integral = False
            ElseIf _commFlag.startup Then
                StartupTask(_commFlag.unitNo)
                _commFlag.startup = False
            ElseIf _commFlag.polling Then
                PollingTask(_commFlag.unitNo)
                _commFlag.polling = False
            End If
        End While
    End Sub
    Private Sub PollingTask(ByVal unitNo As Byte)   '轮询任务
        Dim i As Byte
        For i = 0 To 2      '如果没有收到回复，重复发送三遍
            DownloadCmd.Polling(RS485, _unit(unitNo))
            Me.Invoke(New TextCallback(AddressOf showbyte), RS485.outputbuffer, RS485.outputlength)
            Thread.Sleep(300 + 210)
            If RS485.ReadUp(_readBuffer) Then Exit For
        Next
        If i <= 2 Then ReceievedTackle()
        'DownloadCmd.Polling(RS485, _unit(unitNo))
        'Me.Invoke(New TextCallback(AddressOf showbyte), RS485.outputbuffer, RS485.outputlength)
        'Thread.Sleep(500)
        'RS485.ReadUp(_readBuffer)
        'ReceievedTackle()
    End Sub

    Private Sub StartupTask(ByVal unitNo As Byte)   '启动任务
        Dim i As Byte
        DistributeTask(unitNo)
        For i = 0 To 2
            DownloadCmd.Startup(RS485, _unit(unitNo))
            'Me.Invoke(New TextCallback(AddressOf showbyte), RS485.outputbuffer, RS485.outputlength)
            Thread.Sleep(300 + 120)
            If RS485.ReadUp(_readBuffer) Then Exit For
        Next
        If i <= 2 Then ReceievedTackle()
    End Sub

    Private Sub DistributeTask(ByVal unitNo As Byte)
        Dim i As Byte
        For i = 0 To 2
            DownloadCmd.Distribute(RS485, _unit(unitNo))
            Me.Invoke(New TextCallback(AddressOf showbyte), RS485.outputbuffer, RS485.outputlength)
            Thread.Sleep(300 + 260)
            If RS485.ReadUp(_readBuffer) Then Exit For
        Next
        If i <= 2 Then ReceievedTackle()
    End Sub

    Private Sub IntegralTask()
        Dim i As Byte
        Dim k As Byte
        Dim h As Byte = Now.Hour
        For k = 0 To 23
            If _unit(k).isTesting Then
                While _unit(k).lastHour <> h
                    _unit(k).lastHour = (_unit(k).lastHour + 1) Mod 24
                    '每个小时数据需要4帧问答
                    For i = 0 To 2
                        DownloadCmd.Integral(RS485, _unit(k), 0, _unit(k).lastHour)
                        Me.Invoke(New TextCallback(AddressOf showbyte), RS485.outputbuffer, RS485.outputlength)
                        Thread.Sleep(300 + 320)
                        If RS485.ReadUp(_readBuffer) Then Exit For
                    Next
                    If i <= 2 Then ReceievedTackle()
                    For i = 0 To 2
                        DownloadCmd.Integral(RS485, _unit(k), 1, _unit(k).lastHour)
                        Me.Invoke(New TextCallback(AddressOf showbyte), RS485.outputbuffer, RS485.outputlength)
                        Thread.Sleep(300 + 320)
                        If RS485.ReadUp(_readBuffer) Then Exit For
                    Next
                    If i <= 2 Then ReceievedTackle()
                    For i = 0 To 2
                        DownloadCmd.Integral(RS485, _unit(k), 2, _unit(k).lastHour)
                        Me.Invoke(New TextCallback(AddressOf showbyte), RS485.outputbuffer, RS485.outputlength)
                        Thread.Sleep(300 + 320)
                        If RS485.ReadUp(_readBuffer) Then Exit For
                    Next
                    If i <= 2 Then ReceievedTackle()
                    For i = 0 To 2
                        DownloadCmd.Integral(RS485, _unit(k), 3, _unit(k).lastHour)
                        Me.Invoke(New TextCallback(AddressOf showbyte), RS485.outputbuffer, RS485.outputlength)
                        Thread.Sleep(300 + 320)
                        If RS485.ReadUp(_readBuffer) Then Exit For
                    Next
                    If i <= 2 Then ReceievedTackle()
                End While
            End If
        Next
        OneSec.Enabled = True       '此处有问题，子线程控制控件
        GroupBox1.Enabled = True
    End Sub

    Private Sub ReceievedTackle()   '处理接收到的帧数据，分发给各处理单元
        Dim datalen As Byte = _readBuffer(1)
        Dim address As Byte = _readBuffer(3)
        Dim cmd As Byte = _readBuffer(4)
        Dim data() As Byte
        If datalen Then
            ReDim data(datalen - 3)
            For i = 0 To datalen - 3
                data(i) = _readBuffer(i + 5)
            Next
        End If
        Select Case cmd
            Case &H33 : ReplyNegative(address)     '否认应答
            Case &H3 : ReplyOrdinary(address, data)  '一般请求应答
            Case &HC : ReplyIntegral(address, data)  '整点数据应答
            Case &HF : ReplyDistribute(address)    '参数下发确认
            Case &H30 : ReplyStartup(address)      '启动确认
            Case &H3C : ReplyNotSame(address)      '器件类型不符
        End Select
        BufferReset(_readBuffer)    '处理完毕清空缓存
    End Sub

    Private Sub BufferReset(ByRef buffer() As Byte)
        For i = 0 To buffer.Length - 1
            buffer(i) = 0
        Next
    End Sub

    Private Sub ReplyNegative(ByVal address As Byte)
        
    End Sub
    Private Sub ReplyOrdinary(ByVal address As Byte, ByVal data() As Byte)
        Dim status As Byte = data(0)        '读取状态
        Dim i As Byte                       '循环变量
        Select Case status
            Case &H0    '正常
                Dim type As Byte = (data(11) >> 4)  '读取器件位数
                Dim v_plus As Byte = data(2)        '实测电压偏移量
                Dim pos As Byte = data(1)           '老化位
                Dim v_std As Byte                   '电压标准
                Dim v As Single                     '实际电压
                Dim mA As Single = 0.01             '电流值
                Select Case (data(11) And &H3)
                    Case 0 : v_std = 21
                    Case 1 : v_std = 25
                    Case 2 : v_std = 28
                End Select
                If v_plus >> 7 Then
                    v = v_std - (v_plus And &H8F) * 0.05
                Else
                    v = v_std + (v_plus And &H8F) * 0.05
                End If
                For i = 0 To 23
                    If _unit(i).address = address Then Exit For
                Next

                If type = 0 Then    '1位器件的显示
                    'PollingShow(v, v * mA, i + 1, pos)
                    Me.Invoke(New Polling_dg1(AddressOf PollingShow1), _
                              v, v * mA, CByte(i + 1), pos)
                Else
                    'PollingShow(v, v * mA, i + 1, (pos - 1) \ 4 + 1, (pos - 1) Mod 4)
                    Me.Invoke(New Polling_dg4(AddressOf PollingShow4), _
                             v, v * mA, CByte(i + 1), CByte((pos - 1) \ 4 + 1), CByte((pos - 1) Mod 4))
                End If
            Case &H3    '请求参数
                For i = 0 To 23
                    If _unit(i).address = address Then
                        DistributeTask(i)
                        Exit For
                    End If
                Next
            Case &HC    '340暂停

            Case &H30   '1000停止
                For i = 0 To 23
                    If _unit(i).address = address Then
                        _unit(i).isTesting = False
                    End If
                Next
        End Select
    End Sub
    Private Sub ReplyIntegral(ByVal address As Byte, ByVal data() As Byte)
        Dim i As Byte
        Dim dbcmd As New OleDbCommand
        dbcmd.Connection = _DBconn

        For i = 0 To 23
            If address = _unit(i).address Then Exit For
        Next

        Dim part As Byte = data(0) >> 6
        Dim hour As Byte = data(0) And &H1F

        For k = 1 To 24
            If _unit(i).对位表(k + part * 24) Then
                Dim volt As Single
                If data(k) >> 7 Then
                    volt = _unit(i).单元电压 - (data(k) And &H7F) * 0.05
                Else
                    volt = _unit(i).单元电压 + (data(k) And &H7F) * 0.05
                End If
                Dim power As Single = volt * (0.075 / _unit(i).单元电压)
                DBMethord.WriteResult(_unit(i).试验编号, _unit(i).对位表(k + part * 24), _
                                      hour, volt, power)
            End If
        Next
    End Sub
    Private Sub ReplyDistribute(ByVal address As Byte)
        '保留
    End Sub
    Private Sub ReplyStartup(ByVal address As Byte)
        For i = 0 To 23
            If _unit(i).address = address Then
                _unit(i).isTesting = True
                If Now.Hour = 0 Then
                    _unit(i).lastHour = 23
                Else
                    _unit(i).lastHour = Now.Hour - 1
                End If
                Exit For
            End If
        Next
    End Sub
    Private Sub ReplyNotSame(ByVal address As Byte)
        Dim i As Byte
        Dim str As String
        For i = 0 To 23
            If _unit(i).address = address Then Exit For
        Next
        str = "第" & i + 1 & "号单元器件类型不符，请检查后重新启动！"
        MsgBox(str)
    End Sub

    Private Sub showbyte(ByVal cmd() As Byte, ByVal len As Byte)
        For i = 1 To len
            TextBox1.Text += cmd(i - 1).ToString("X2") & " "
        Next
        TextBox1.Text += vbNewLine
    End Sub

    
    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        TextBox1.SelectionStart = TextBox1.TextLength
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub btnBegin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBegin.Click
        _commFlag.unitNo = 0
        _commFlag.startup = True
    End Sub

    Private Sub cmbUnitNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbUnitNo.SelectedIndexChanged
        Dim index As Byte = cmbUnitNo.SelectedIndex
        TabControl1.SelectedIndex = index
        lblVolt.Text = _unit(index).电压规格 & " V"
        If _unit(index).座子类型 Then
            lblLeg.Text = "一位"
        Else
            lblLeg.Text = "四位"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        RS485.ReadUp(_readBuffer)
        For i = 0 To _readBuffer.Length - 1
            TextBox1.Text += _readBuffer(i).ToString("X2") & " "
        Next
    End Sub
End Class
