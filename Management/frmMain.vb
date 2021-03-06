﻿Imports System.Data.OleDb
Imports System.IO.Ports
Imports System.Threading
Imports LeoControls

Public Class frmMain
    Dim ShowList(23) As ArrayList
    Delegate Sub TextCallback(ByVal txt As TextBox, ByVal cmd() As Byte, ByVal len As Byte)
    Delegate Sub Polling_dg1(ByVal volt As Single, ByVal power As Single, _
                             ByVal unitnum As Byte, ByVal cellnum As Byte)
    Delegate Sub Polling_dg4(ByVal volt As Single, ByVal power As Single, _
                             ByVal unitnum As Byte, ByVal cellnum As Byte, ByVal cellpos As Byte)

    Public RS485 As New LHSerialPort("COM3", 1200, Parity.Even, 8, 1)

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

        'OneSec.Enabled = Not OneSec.Enabled
        '_commFlag.integral = True
        _commFlag.unitNo = 0
        _commFlag.startup = True
        'Dim cmd As New OleDbCommand
        'cmd.Connection = _DBconn
        'cmd.CommandText = "delete * from 试验结果"
        'cmd.ExecuteNonQuery()
        'Button1.Text = "ok"
    End Sub '测试

    Private Sub PaintShow()
        For k = 0 To 23
            ShowList(k) = New ArrayList
            If Not _unit(k).座子类型 Then       '1位器件
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
        _commFlag.timeModify = False
        _commFlag.reboot340 = False
        _commFlag.reboot = False
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
        'RS485.RtsEnable = True
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
        txtSend.Hide()
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
        txtSend.Show()
        GroupBox1.Show()
        PictureBox1.Show()
        MenuStrip1.Show()
        TabControl1.Show()

        btnMax.Hide()
    End Sub '切换为全屏显示

    Private Sub OneSec_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OneSec.Tick
        Static i As Byte = 0
        While _unit(i).Testing <> 0
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
            End If
            If _commFlag.startup Then
                StartupTask(_commFlag.unitNo)
                _commFlag.startup = False
            End If
            If _commFlag.polling Then
                PollingTask(_commFlag.unitNo)
                _commFlag.polling = False
            End If
            If _commFlag.timeModify Then
                TimeModifyTask()
                _commFlag.timeModify = False
            End If
            If _commFlag.reboot340 Then
                Reboot340Task(_commFlag.unitNo)
                _commFlag.reboot340 = False
            End If
            If _commFlag.reboot Then
                RebootTask(_commFlag.unitNo)
                _commFlag.reboot = False
            End If
        End While
    End Sub

    Private Sub PollingTask(ByVal unitNo As Byte)   '轮询任务
        For i = 0 To 2      '如果没有收到回复，重复发送三遍
            DownloadCmd.Polling(RS485, _unit(unitNo))
            Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
            Thread.Sleep(300 + 210)
            Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
            Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
            If  cmdBack= &H3 Then
                ReplyPolling()
                Exit For
            End If
        Next
    End Sub

    Private Sub StartupTask(ByVal unitNo As Byte)   '启动任务
        Dim i As Byte
        For i = 0 To 2
            DownloadCmd.Distribute(RS485, _unit(unitNo))
            Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
            Thread.Sleep(300 + 260)
            Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
            Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
            If cmdBack = LHSerialPort.cmdNotSame Then
                MsgBox("器件类型与老化单元不符，请检查后重新启动！")
                Exit Sub
            End If
            If cmdBack = LHSerialPort.cmdDistribute Then
                DownloadCmd.Startup(RS485, _unit(unitNo))
                Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
                Thread.Sleep(300 + 120)
                Dim cmd As Byte = RS485.ReadUp(_readBuffer)
                Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
                If cmd = LHSerialPort.cmdStartup Then
                    With _unit(unitNo)
                        .Testing = &H0
                        .lastHour = Now.Hour - 1
                    End With
                    '还未写回数据库
                    Exit Sub
                End If
            End If
        Next
    End Sub

    Private Sub DistributeTask(ByVal unitNo As Byte)
        For i = 0 To 2
            DownloadCmd.Distribute(RS485, _unit(unitNo))
            Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
            Thread.Sleep(300 + 260)
            Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
            Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
            If cmdBack = LHSerialPort.cmdDistribute Then Exit For
        Next
    End Sub

    Private Sub IntegralTask()
        '招记录信息的用法


        Dim h As Byte = Now.Hour
        For k = 0 To 23
            If _unit(k).Testing <> 0 Then   '该判断依据需要修改
                While _unit(k).lastHour <> h
                    _unit(k).lastHour = (_unit(k).lastHour + 1) Mod 24
                    '每个小时数据需要4帧问答
                    For i = 0 To 2
                        DownloadCmd.Integral(RS485, _unit(k), 0, _unit(k).lastHour)
                        Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
                        Thread.Sleep(300 + 320)
                        Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
                        Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
                        If cmdBack = LHSerialPort.cmdIntegral Then
                            ReplyIntegral()
                            Exit For
                        End If
                    Next
                    For i = 0 To 2S
                        DownloadCmd.Integral(RS485, _unit(k), 1, _unit(k).lastHour)
                        Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
                        Thread.Sleep(300 + 320)
                        Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
                        Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
                        If cmdBack = LHSerialPort.cmdIntegral Then
                            ReplyIntegral()
                            Exit For
                        End If
                    Next
                    
                    For i = 0 To 2
                        DownloadCmd.Integral(RS485, _unit(k), 2, _unit(k).lastHour)
                        Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
                        Thread.Sleep(300 + 320)
                        Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
                        Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
                        If cmdBack = LHSerialPort.cmdIntegral Then
                            ReplyIntegral()
                            Exit For
                        End If
                    Next
                    For i = 0 To 2
                        DownloadCmd.Integral(RS485, _unit(k), 3, _unit(k).lastHour)
                        Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
                        Thread.Sleep(300 + 320)
                        Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
                        Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
                        If cmdBack = LHSerialPort.cmdIntegral Then
                            ReplyIntegral()
                            Exit For
                        End If
                    Next
                End While
                DBMethord.UpdateHour(k + 1, _unit(k).lastHour)
            End If
        Next
        _commFlag.integral = False

        OneSec.Enabled = True       '此处可能有问题，子线程控制控件
        GroupBox1.Enabled = True
    End Sub

    Private Sub TimeModifyTask()
        DownloadCmd.TimeModify(RS485)
        Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
    End Sub

    Private Sub RebootTask(ByVal unitNo As Byte)
        Dim i As Byte
        For i = 0 To 2
            DownloadCmd.Distribute(RS485, _unit(unitNo))
            Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
            Thread.Sleep(300 + 260)
            Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
            Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
            If cmdBack = LHSerialPort.cmdNotSame Then
                MsgBox("器件类型与老化单元不符，请检查后重新启动！")
                Exit Sub
            End If
            If cmdBack = LHSerialPort.cmdDistribute Then
                DownloadCmd.Reboot(RS485, _unit(unitNo))
                Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
                Thread.Sleep(300 + 120)
                Dim cmd As Byte = RS485.ReadUp(_readBuffer)
                Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
                If cmd = LHSerialPort.cmdReboot Then
                    With _unit(unitNo)
                        .Testing = &H0
                        .lastHour = Now.Hour - 1
                    End With
                    MsgBox("强制重启成功！")
                    '还未写回数据库
                    Exit Sub
                End If
            End If
        Next
    End Sub

    Private Sub Reboot340Task(ByVal unitNo As Byte)
        Dim i As Byte
        For i = 0 To 2
            DownloadCmd.Distribute(RS485, _unit(unitNo))
            Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
            Thread.Sleep(300 + 260)
            Dim cmdBack As Byte = RS485.ReadUp(_readBuffer)
            Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
            If cmdBack = LHSerialPort.cmdNotSame Then
                MsgBox("器件类型与老化单元不符，请检查后重新启动！")
                Exit Sub
            End If
            If cmdBack = LHSerialPort.cmdDistribute Then
                DownloadCmd.Reboot340(RS485, _unit(unitNo))
                Me.Invoke(New TextCallback(AddressOf showbyte), txtSend, RS485.outputbuffer, RS485.outputlength)
                Thread.Sleep(300 + 120)
                Dim cmd As Byte = RS485.ReadUp(_readBuffer)
                Me.Invoke(New TextCallback(AddressOf showbyte), txtRecv, _readBuffer, _readBuffer(1) + 5)
                If cmd = LHSerialPort.cmdReboot340 Then
                    With _unit(unitNo)
                        .Testing = &H0
                        .lastHour = Now.Hour - 1
                    End With
                    MsgBox("放弃340小时后试验并重启成功！")
                    '还未写回数据库
                    Exit Sub
                End If
            End If
        Next
    End Sub

    Private Sub ReplyPolling()
        Dim datalen As Byte = _readBuffer(1)
        Dim address As Byte = _readBuffer(3)
        Dim Data(datalen - 3) As Byte
        For j = 0 To datalen - 3
            Data(j) = _readBuffer(j + 5)
        Next

        Dim status As Byte = Data(0)        '读取状态
        Dim i As Byte                       '循环变量
        Select Case status
            Case &H0    '正常
                Dim type As Byte = (Data(11) >> 4)  '读取器件位数
                Dim v_plus As Byte = Data(2)        '实测电压偏移量
                Dim pos As Byte = Data(1)           '老化位
                Dim volt As Single                  '实际电压
                Dim power As Single                 '实际功率
                Select Case (Data(11) And &H3)
                    Case 0
                        volt = (v_plus * 1175 + 387200) / 25600
                        power = 75 * volt / 21
                    Case 1
                        volt = (v_plus * 1375 + 464000) / 25600
                        power = 75 * volt / 25
                    Case 2
                        volt = (v_plus * 1525 + 521600) / 25600
                        power = 75 * volt / 28
                End Select

                For i = 0 To 23
                    If _unit(i).address = address Then Exit For
                Next

                If type = 0 Then    '1位器件的显示
                    'PollingShow(v, v * mA, i + 1, pos)
                    Me.Invoke(New Polling_dg1(AddressOf PollingShow1), _
                              volt, power, CByte(i + 1), pos)
                Else
                    'PollingShow(v, v * mA, i + 1, (pos - 1) \ 4 + 1, (pos - 1) Mod 4)
                    Me.Invoke(New Polling_dg4(AddressOf PollingShow4), _
                             volt, power, CByte(i + 1), CByte((pos - 1) \ 4 + 1), CByte((pos - 1) Mod 4))
                End If
            Case &H3    '请求参数
                For i = 0 To 23
                    If _unit(i).address = address Then
                        DistributeTask(i)
                        Exit For
                    End If
                Next
            Case &HC    '340暂停
                For i = 0 To 23
                    If _unit(i).address = address Then
                        _unit(i).Testing = &HC
                        Exit For
                    End If
                Next
            Case &H30   '1000停止
                For i = 0 To 23
                    If _unit(i).address = address Then
                        _unit(i).Testing = &H30
                        Exit For
                    End If
                Next
        End Select
    End Sub

    Private Sub ReplyIntegral()
        Dim datalen As Byte = _readBuffer(1)
        Dim address As Byte = _readBuffer(3)
        Dim Data(datalen - 3) As Byte
        For j = 0 To datalen - 3
            Data(j) = _readBuffer(j + 5)
        Next

        Dim i As Byte
        For i = 0 To 23
            If address = _unit(i).address Then Exit For
        Next

        Dim part As Byte = data(0) >> 6
        Dim hour As Byte = data(0) And &H1F

        If _unit(i).器件类型 = 0 Then       '1位器件的数据压入
            For k = 0 To 23
                If _unit(i).对位表(k + part * 24) <> 0 Then
                    Dim volt As Single, power As Single
                    Select Case _unit(i).单元电压
                        Case 21
                            volt = (Data(k) * 1175 + 387200) / 25600
                            power = 75 * volt / 21
                        Case 25
                            volt = (Data(k) * 1375 + 464000) / 25600
                            power = 75 * volt / 25
                        Case 28
                            volt = (Data(k) * 1525 + 521600) / 25600
                            power = 75 * volt / 28
                    End Select

                    DBMethord.WriteResult(_unit(i).试验编号, _unit(i).对位表(k + part * 24), _
                                          CByte(1), hour, volt, power)
                End If
            Next
        End If

        If _unit(i).器件类型 = 1 Then       '2位器件的数据压入
            Dim isFirst As Boolean = True
            For k = 0 To 23
                If _unit(i).对位表(k + part * 24) Then
                    Dim volt As Single, power As Single
                    Select Case _unit(i).单元电压
                        Case 21
                            volt = (data(k) * 1175 + 387200) / 25600
                            power = 75 * volt / 21
                        Case 25
                            volt = (data(k) * 1375 + 464000) / 25600
                            power = 75 * volt / 25
                        Case 28
                            volt = (data(k) * 1525 + 521600) / 25600
                            power = 75 * volt / 28
                    End Select

                    Dim subnum As Byte
                    If isFirst Then
                        subnum = 1
                    Else
                        subnum = 2
                    End If
                    DBMethord.WriteResult(_unit(i).试验编号, _unit(i).对位表(k + part * 24), _
                                          subnum, hour, volt, power)
                    isFirst = Not isFirst
                End If
            Next
        End If

        If _unit(i).器件类型 = 2 Then       '4位器件的数据压入
            For k = 0 To 23
                If _unit(i).对位表(k + part * 24) Then
                    Dim volt As Single, power As Single
                    Select Case _unit(i).单元电压
                        Case 21
                            volt = (data(k) * 1175 + 387200) / 25600
                            power = 75 * volt / 21
                        Case 25
                            volt = (data(k) * 1375 + 464000) / 25600
                            power = 75 * volt / 25
                        Case 28
                            volt = (data(k) * 1525 + 521600) / 25600
                            power = 75 * volt / 28
                    End Select

                    DBMethord.WriteResult(_unit(i).试验编号, _unit(i).对位表(k + part * 24), _
                                          CByte((k + part * 24) Mod 4 + 1), hour, volt, power)
                End If
            Next
        End If
    End Sub

    Private Sub showbyte(ByVal txt As TextBox, ByVal cmd() As Byte, ByVal len As Byte)
        For i = 1 To len
            txt.Text += cmd(i - 1).ToString("X2") & " "
        Next
        txt.Text += vbNewLine
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSend.TextChanged
        txtSend.SelectionStart = txtSend.TextLength
        txtSend.ScrollToCaret()
    End Sub

    Private Sub btnBegin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBegin.Click
        Dim i As Byte = cmbUnitNo.SelectedIndex
        _unit(i).产品型号 = cmbType.SelectedItem
        _unit(i).生产批号 = txtManufact.Text
        _unit(i).标准号 = txtStandard.Text
        _unit(i).试验编号 = txtTestNo.Text
        _unit(i).操作员 = txtOperator.Text


        _commFlag.unitNo = 0
        _commFlag.startup = True
    End Sub

    Private Sub cmbUnitNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbUnitNo.SelectedIndexChanged
        Dim index As Byte = cmbUnitNo.SelectedIndex
        TabControl1.SelectedIndex = index
        lblVolt.Text = _unit(index).电压规格 & " V"
        If Not _unit(index).座子类型 Then
            lblSeatLeg.Text = "一位"
        Else
            lblSeatLeg.Text = "四位"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        RS485.ReadUp(_readBuffer)
        For i = 0 To _readBuffer.Length - 1
            txtSend.Text += _readBuffer(i).ToString("X2") & " "
        Next
    End Sub

    Private Sub txtOperator_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOperator.Click
        frmInputOperator.ShowDialog()
    End Sub

    Private Sub btnPos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPos.Click
        Dim i As Byte = cmbUnitNo.SelectedIndex
        If _unit(i).座子类型 = 0 And _unit(i).器件类型 = 0 Then '单位插单位
            Dim frm As New frmPosChart1
            frm.unitNo = i
            frm.Show()
        End If
        If _unit(i).座子类型 = 1 And _unit(i).器件类型 = 1 Then '四位插双位
            Dim frm As New frmPosChart2
            frm.unitNo = i
            frm.Show()
        End If
        If _unit(i).座子类型 = 1 And _unit(i).器件类型 = 2 Then '四位插四位
            Dim frm As New frmPosChart4
            frm.unitNo = i
            frm.Show()
        End If
    End Sub
End Class
