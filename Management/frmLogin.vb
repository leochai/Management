Public Class frmLogin
    Dim isOpen As Boolean
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        If _管理员密码 = txtPsw.Text Then
            isOpen = True
            Me.Dispose()
        Else
            txtPsw.Text = ""
            txtPsw.Focus()
            lblWarning.Text = "密码错误，请重新输入！"
            isOpen = False
        End If
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        isOpen = False
        Me.Dispose()
    End Sub

    Private Sub frmLogin_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If isOpen Then frm操作员.ShowDialog()
    End Sub

    Private Sub frmLogin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        isOpen = False
        txtPsw.Focus()
    End Sub

End Class