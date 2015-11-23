Imports System.Data.OleDb
Imports System.IO.Ports
Imports System.Threading
Imports LeoControls

Public Class frmMain
    Dim ShowList(23) As ArrayList
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '数据库操作
        _DBconn.Open()

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
                        sshow(j * 6 + i) = New OneShow
                        With sshow(j * 6 + i)
                            .Left = 69 * i + 8
                            .Top = 118 * j + 50
                            .Parent = TabControl1.TabPages(k)
                            .ShowNum = j * 6 + i + 1
                        End With
                        ShowList(k).Add(sshow(j * 6 + i))
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
                            .Parent = TabControl1.TabPages(k)
                            .ShowNum = j * 6 + i + 1
                        End With
                        ShowList(k).Add(sshow(j * 6 + i))
                    Next
                Next
            End If
        Next
        
    End Sub


    Private Sub 操作员管理ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 操作员管理ToolStripMenuItem.Click
        frmLogin.ShowDialog()
    End Sub

    Private Sub 试验数据检索ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 试验数据检索ToolStripMenuItem.Click
        frm数据检索.ShowDialog()
    End Sub

    Private Sub frmMain_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        ShowList(12).Item(10).setResult(5.5, 6.6, 2)
    End Sub
End Class
