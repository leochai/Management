Imports System.Data.OleDb
Imports System.IO.Ports
Imports System.Threading
Imports LeoControls

Public Class frmMain
    Dim ShowList(23) As ArrayList
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '数据库操作
        _DBconn.Open()

        Dim fs As New frmRoot
        fs.Show()
        fs.Label1.Text = "加载数据..."
        fs.Refresh()

        Dim com As New OleDbCommand
        com.Connection = _DBconn
        com.CommandText = "select 器件类型 from 单元状态"
        Dim reader As OleDbDataReader = com.ExecuteReader()
        Dim mytype(23) As Byte
        Dim a As Byte = 0
        While reader.Read
            mytype(a) = reader("器件类型")
            a += 1
        End While

        For k = 0 To 23
            ShowList(k) = New ArrayList
            If mytype(k) = 1 Then
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
            Else
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
        'fs.Refresh()
        'fs.Label1.Text = "加载完毕."
        'fs.Refresh()
        'System.Threading.Thread.Sleep(1000)
        'fs.Close()
    End Sub


    Private Sub 操作员管理ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 操作员管理ToolStripMenuItem.Click
        frmLogin.ShowDialog()
    End Sub

    Private Sub 试验数据检索ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 试验数据检索ToolStripMenuItem.Click
        frm数据检索.ShowDialog()
    End Sub

    Private Sub frmMain_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        RefreshData(36.11, 12.34, 15, 24)
        RefreshData(36.11, 12.34, 1, 22, 4)
        RefreshData(36.11, 12.34, 2, 1, 1)
    End Sub

    Public Sub RefreshData(ByVal volt As Single, ByVal power As Single, _
                           ByVal unitnum As Byte, ByVal cellnum As Byte)
        ShowList(unitnum - 1).Item(cellnum - 1).setResult(volt, power)
    End Sub

    Public Sub RefreshData(ByVal volt As Single, ByVal power As Single, _
                           ByVal unitnum As Byte, ByVal cellnum As Byte, ByVal cellpos As Byte)
        ShowList(unitnum - 1).Item(cellnum - 1).setResult(volt, power, cellpos)
    End Sub
End Class
