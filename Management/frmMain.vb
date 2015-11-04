Option Explicit On
Imports System.Data.OleDb

Imports System.Threading

Public Class frmMain
    Dim flag As Boolean
    Private Sub frmMain_mousedown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseDown
        flag = Not flag
        SwitchLight.TurnOnOff(flag)

    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        flag = False

        For i = 1 To 48
            dgv编号.Rows.Add(1)
            dgv编号(0, i - 1).Value = i '先列后行
        Next
        cbo单元号.SelectedIndex = 0
        cbo型号.SelectedIndex = 0


        '数据库操作

        _DBconn.Open()


    End Sub

    Private Sub btn自动编号_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn自动编号.Click
        Dim row As Integer
        Dim myvalue As Integer


        row = dgv编号.CurrentCell.RowIndex
        myvalue = dgv编号.CurrentCell.Value

        For i = row + 1 To 47
            myvalue += 1
            dgv编号(1, i).Value = myvalue
        Next

    End Sub


    Private Sub 操作员管理ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 操作员管理ToolStripMenuItem.Click
        frmLogin.ShowDialog()
    End Sub

    Private Sub 试验数据检索ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 试验数据检索ToolStripMenuItem.Click
        frm数据检索.ShowDialog()
    End Sub
End Class
