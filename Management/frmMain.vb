Imports System.Data.OleDb
Imports System.IO.Ports
Imports System.Threading
Imports LeoControls

Public Class frmMain
    Dim ShowList(23) As ArrayList
    Public WithEvents RS232 As New LHSerialPort("COM3", 1200, Parity.Odd, 8, 1)

    Private Sub frmMain_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        For Each item In _unit
            If item.isTesting Then
                DownloadCmd.Polling(_COM, item)
                Threading.Thread.Sleep(20)
            End If
        Next
    End Sub
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '------打开数据库
        _DBconn.Open()
        RS232.WriteBufferSize = 2048
        RS232.ReadTimeout = 1000
        RS232.ReceivedBytesThreshold = 3
        RS232.Open()

        Dim fs As New frmRoot
        fs.Show()
        fs.Label1.Text = "加载数据..."
        fs.Refresh()

        For i = 0 To 23
            _unit(i) = New LHUnit
        Next
        DBMethord.Initial(_DBconn, _unit)

        For k = 0 To 23
            ShowList(k) = New ArrayList
            If _unit(k).器件类型 = 0 Then       '1位器件
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
        'OneSec.Enabled = True
        fs.Close()
    End Sub


    Private Sub 操作员管理ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 操作员管理ToolStripMenuItem.Click
        frmLogin.ShowDialog()
    End Sub

    Private Sub 试验数据检索ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 试验数据检索ToolStripMenuItem.Click
        frm数据检索.ShowDialog()
    End Sub

    Private Sub frmMain_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        
    End Sub

    Public Sub RefreshData(ByVal volt As Single, ByVal power As Single, _
                           ByVal unitnum As Byte, ByVal cellnum As Byte)
        ShowList(unitnum - 1).Item(cellnum - 1).setResult(volt, power)
    End Sub '更新实时数据显示

    Public Sub RefreshData(ByVal volt As Single, ByVal power As Single, _
                           ByVal unitnum As Byte, ByVal cellnum As Byte, ByVal cellpos As Byte)
        ShowList(unitnum - 1).Item(cellnum - 1).setResult(volt, power, cellpos)
    End Sub '更新实时数据显示

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
        For Each item In _unit      '轮询
            If item.isTesting Then
                DownloadCmd.Polling(_COM, item)
            End If
        Next

    End Sub

    Private Sub RS232_DataReceived(ByVal sender As Object, _
                                  ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles RS232.DataReceived
        If RS232.ReadUp(_readBuffer) Then

        End If
    End Sub

    Private Sub ReceievedTackle()
        Dim datalen As Byte = _readBuffer(1)
        Dim address As Byte = _readBuffer(3)
        Dim cmd As Byte = _readBuffer(4)
        Dim data() As Byte
        If datalen Then
            ReDim data(datalen - 1)
            For i = 0 To datalen - 1
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
        BufferReset(_readBuffer)
    End Sub

    Private Sub BufferReset(ByRef buffer() As Byte)
        For i = 0 To buffer.Length - 1
            buffer(i) = 0
        Next
    End Sub

    Private Sub ReplyNegative(ByVal address As Byte)

    End Sub
    Private Sub ReplyOrdinary(ByVal address As Byte, ByVal data() As Byte)

    End Sub
    Private Sub ReplyIntegral(ByVal address As Byte, ByVal data() As Byte)

    End Sub
    Private Sub ReplyDistribute(ByVal address As Byte)

    End Sub
    Private Sub ReplyStartup(ByVal address As Byte)

    End Sub
    Private Sub ReplyNotSame(ByVal address As Byte)

    End Sub
End Class
