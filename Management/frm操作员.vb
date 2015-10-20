Imports System.Data.OleDb

Public Class frm操作员
    '调试完成后删除
    Dim _DBconn As New OleDbConnection("Provider=Microsoft.Ace.OleDb.12.0;Data Source=D:/老化台.accdb")

    Private Sub frm操作员_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        '测试模块
        Dim DBcmd As New OleDbCommand()
        DBcmd.Connection = _DBconn
        DBcmd.CommandText = "select * from 操作员 where 姓名 = '鸿和'"
        txt新增姓名.Text = DBcmd.ExecuteNonQuery().ToString
    End Sub

    Private Sub frm操作员_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        '调试完成后删除
        _DBconn.Close()
    End Sub
    Private Sub txt修改旧密码_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt修改旧密码.KeyPress
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txt修改新密码_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt修改新密码.KeyPress
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txt修改重复密码_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt修改重复密码.KeyPress
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txt新增密码_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt新增密码.KeyPress
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txt新增重复密码_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt新增重复密码.KeyPress
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txt删除密码_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt删除密码.KeyPress
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txt新增姓名_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt新增姓名.KeyPress
        If e.KeyChar = Chr(32) Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub frm操作员_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '=============
        _DBconn.Open()
        '=============调试完成后删除
        updatelst()
    End Sub

    Private Sub btn新增_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn新增.Click
        Dim DBcmd As New OleDbCommand()
        DBcmd.Connection = _DBconn
        
        '处理姓名为空的情况
        If txt新增姓名.Text = "" Then
            MsgBox("姓名不能为空！", MsgBoxStyle.Exclamation, "警告")
            txt新增姓名.Text = ""
            txt新增密码.Text = ""
            txt新增重复密码.Text = ""
            Exit Sub
        End If
        '处理密码不足8位的情况
        If txt新增密码.TextLength < 8 Then
            MsgBox("请输入8位数字密码！", MsgBoxStyle.Exclamation, "警告")
            txt新增密码.Text = ""
            txt新增重复密码.Text = ""
            Exit Sub
        End If
        '处理存在同名用户的情况
        DBcmd.CommandText = "select * from 操作员 where 姓名 = '" & txt新增姓名.Text & "'"
        If DBcmd.ExecuteReader().HasRows Then
            MsgBox("已存在同名用户", MsgBoxStyle.Exclamation, "警告")
            txt新增姓名.Text = ""
            txt新增密码.Text = ""
            txt新增重复密码.Text = ""
            Exit Sub
        End If
        '处理两次密码输入不同
        If txt新增密码.Text = txt新增重复密码.Text Then
            DBcmd.CommandText = "insert into 操作员 values('" & txt新增姓名.Text & "','" & txt新增密码.Text & "')"
            DBcmd.ExecuteNonQuery()
            txt新增姓名.Text = ""
            txt新增密码.Text = ""
            txt新增重复密码.Text = ""
            updatelst()
        Else
            MsgBox("两次输入密码不相同！", MsgBoxStyle.Exclamation, "警告")
            txt新增密码.Text = ""
            txt新增重复密码.Text = ""
        End If
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Dispose()
    End Sub

    Private Sub updatelst()
        Dim DBreader As OleDbDataReader
        Dim DBcmd As New OleDbCommand()
        DBcmd.Connection = _DBconn
        DBcmd.CommandText = "select 姓名 from 操作员"
        DBreader = DBcmd.ExecuteReader()
        lst操作员.Items.Clear()

        If DBreader.HasRows Then
            While DBreader.Read
                lst操作员.Items.Add(DBreader("姓名"))
            End While
        End If
    End Sub


    Private Sub btn修改_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn修改.Click
        Dim DBcmd As New OleDbCommand()
        Dim DBreader As OleDbDataReader
        DBcmd.Connection = _DBconn

        '处理姓名为空情况
        If txt修改姓名.Text = "" Then
            MsgBox("请从左侧列表选择用户！", MsgBoxStyle.Exclamation, "警告")
            txt修改旧密码.Text = ""
            txt修改新密码.Text = ""
            txt修改重复密码.Text = ""
            Exit Sub
        End If

        '处理原始密码错误的情况
        DBcmd.CommandText = "select 密码 from 操作员 where 姓名 = '" & txt修改姓名.Text & "'"
        DBreader = DBcmd.ExecuteReader()
        DBreader.Read()
        If Not (DBreader("密码") = txt修改旧密码.Text) Then
            MsgBox("旧密码输入错误！", MsgBoxStyle.Exclamation, "警告")
            txt修改旧密码.Text = ""
            txt修改新密码.Text = ""
            txt修改重复密码.Text = ""
            Exit Sub
        End If
        DBreader.Close()

        '处理密码不足8位的情况
        If txt修改新密码.TextLength < 8 Then
            MsgBox("请输入8位新密码！", MsgBoxStyle.Exclamation, "警告")
            txt修改旧密码.Text = ""
            txt修改新密码.Text = ""
            txt修改重复密码.Text = ""
            Exit Sub
        End If

        '处理两次密码输入不同
        If txt修改新密码.Text = txt修改重复密码.Text Then
            DBcmd.CommandText = "update 操作员 set 密码 = '" & txt修改新密码.Text & "' where 姓名='" & txt修改姓名.Text & "'"
            DBcmd.ExecuteNonQuery()
            txt修改姓名.Text = ""
            txt修改旧密码.Text = ""
            txt修改新密码.Text = ""
            txt修改重复密码.Text = ""
            MsgBox("密码修改成功！", MsgBoxStyle.Information, "提示")
        Else
            MsgBox("两次输入密码不相同！", MsgBoxStyle.Exclamation, "警告")
            txt修改旧密码.Text = ""
            txt修改新密码.Text = ""
            txt修改重复密码.Text = ""
        End If
    End Sub

    Private Sub btn删除_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn删除.Click
        Dim DBcmd As New OleDbCommand()
        Dim DBreader As OleDbDataReader
        DBcmd.Connection = _DBconn

        '处理姓名为空情况
        If txt删除姓名.Text = "" Then
            MsgBox("请从左侧列表选择用户！", MsgBoxStyle.Exclamation, "警告")
            txt删除密码.Text = ""
            Exit Sub
        End If

        '处理密码错误的情况
        DBcmd.CommandText = "select 密码 from 操作员 where 姓名='" & txt删除姓名.Text & "'"
        DBreader = DBcmd.ExecuteReader
        DBreader.Read()
        If txt删除密码.Text = _str管理员密码 Or txt删除密码.Text = DBreader("密码") Then
            DBreader.Close()
            DBcmd.CommandText = "delete from 操作员 where 姓名='" & txt删除姓名.Text & "'"
            DBcmd.ExecuteNonQuery()
            txt删除姓名.Text = ""
            txt删除密码.Text = ""
            updatelst()
        Else
            MsgBox("密码错误！", MsgBoxStyle.Exclamation, "警告")
            txt删除密码.Text = ""
        End If
    End Sub
    Private Sub lst操作员_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lst操作员.Click
        If TabControl1.SelectedIndex = 1 Then
            txt修改姓名.Text = lst操作员.Text
        ElseIf TabControl1.SelectedIndex = 2 Then
            txt删除姓名.Text = lst操作员.Text
        End If
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        If TabControl1.SelectedIndex = 0 Then
            txt修改姓名.Text = ""
            txt修改旧密码.Text = ""
            txt修改新密码.Text = ""
            txt修改重复密码.Text = ""
            txt删除姓名.Text = ""
            txt删除密码.Text = ""
        ElseIf TabControl1.SelectedIndex = 1 Then
            txt新增姓名.Text = ""
            txt新增密码.Text = ""
            txt新增重复密码.Text = ""
            txt删除姓名.Text = ""
            txt删除密码.Text = ""
        ElseIf TabControl1.SelectedIndex = 2 Then
            txt新增姓名.Text = ""
            txt新增密码.Text = ""
            txt新增重复密码.Text = ""
            txt修改姓名.Text = ""
            txt修改旧密码.Text = ""
            txt修改新密码.Text = ""
            txt修改重复密码.Text = ""
        End If
    End Sub
End Class