﻿Imports System.Data.OleDb

Public Class frmInputOperator

    ' TODO: 插入代码，以使用提供的用户名和密码执行自定义的身份验证
    ' (请参见 http://go.microsoft.com/fwlink/?LinkId=35339)。 
    ' 随后自定义主体可附加到当前线程的主体，如下所示: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' 其中 CustomPrincipal 是用于执行身份验证的 IPrincipal 实现。 
    ' 随后，My.User 将返回 CustomPrincipal 对象中封装的标识信息
    ' 如用户名、显示名等

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim dbcmd As New OleDbCommand
        Dim reader As OleDbDataReader
        dbcmd.Connection = _DBconn
        dbcmd.CommandText = "select 密码 from 操作员 where 姓名 = '" & UsernameTextBox.Text & "'"
        reader = dbcmd.ExecuteReader
        If Not reader.HasRows Then
            MsgBox("姓名错误！")
            UsernameTextBox.Text = ""
            PasswordTextBox.Text = ""
        Else
            reader.Read()
            If reader("密码") = PasswordTextBox.Text Then
                frmMain.txtOperator.Text = UsernameTextBox.Text
                Me.Close()
            Else
                MsgBox("密码错误！")
                UsernameTextBox.Text = ""
                PasswordTextBox.Text = ""
            End If
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
