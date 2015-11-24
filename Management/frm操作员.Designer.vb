<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm操作员
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lst操作员 = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Label5 = New System.Windows.Forms.Label
        Me.btn新增 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txt新增重复密码 = New System.Windows.Forms.TextBox
        Me.txt新增密码 = New System.Windows.Forms.TextBox
        Me.txt新增姓名 = New System.Windows.Forms.TextBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.txt修改重复密码 = New System.Windows.Forms.TextBox
        Me.txt修改新密码 = New System.Windows.Forms.TextBox
        Me.txt修改旧密码 = New System.Windows.Forms.TextBox
        Me.txt修改姓名 = New System.Windows.Forms.TextBox
        Me.btn修改 = New System.Windows.Forms.Button
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.Label13 = New System.Windows.Forms.Label
        Me.txt删除密码 = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.txt删除姓名 = New System.Windows.Forms.TextBox
        Me.btn删除 = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'lst操作员
        '
        
        Me.lst操作员.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.lst操作员.FormattingEnabled = True
        Me.lst操作员.ItemHeight = 16
        Me.lst操作员.Location = New System.Drawing.Point(31, 49)
        Me.lst操作员.Name = "lst操作员"
        Me.lst操作员.Size = New System.Drawing.Size(81, 260)
        Me.lst操作员.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(43, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "操作员"
        '
        'btnExit
        '
        Me.btnExit.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btnExit.Location = New System.Drawing.Point(354, 289)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(88, 36)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "退  出"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(148, 22)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(294, 261)
        Me.TabControl1.TabIndex = 4
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Transparent
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.btn新增)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.txt新增重复密码)
        Me.TabPage1.Controls.Add(Me.txt新增密码)
        Me.TabPage1.Controls.Add(Me.txt新增姓名)
        Me.TabPage1.Location = New System.Drawing.Point(4, 26)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(286, 231)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "新增"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(107, 115)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(144, 16)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "请输入8位数字密码"
        '
        'btn新增
        '
        Me.btn新增.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btn新增.Location = New System.Drawing.Point(211, 196)
        Me.btn新增.Name = "btn新增"
        Me.btn新增.Size = New System.Drawing.Size(72, 29)
        Me.btn新增.TabIndex = 3
        Me.btn新增.Text = "确认"
        Me.btn新增.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(26, 147)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 16)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "重复密码"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(34, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "密  码"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(34, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "姓  名"
        '
        'txt新增重复密码
        '
        Me.txt新增重复密码.BackColor = System.Drawing.SystemColors.Window
        Me.txt新增重复密码.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt新增重复密码.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txt新增重复密码.Location = New System.Drawing.Point(110, 143)
        Me.txt新增重复密码.MaxLength = 8
        Me.txt新增重复密码.Name = "txt新增重复密码"
        Me.txt新增重复密码.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt新增重复密码.Size = New System.Drawing.Size(100, 26)
        Me.txt新增重复密码.TabIndex = 2
        '
        'txt新增密码
        '
        Me.txt新增密码.BackColor = System.Drawing.SystemColors.Window
        Me.txt新增密码.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt新增密码.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txt新增密码.Location = New System.Drawing.Point(110, 75)
        Me.txt新增密码.MaxLength = 8
        Me.txt新增密码.Name = "txt新增密码"
        Me.txt新增密码.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt新增密码.Size = New System.Drawing.Size(100, 26)
        Me.txt新增密码.TabIndex = 1
        '
        'txt新增姓名
        '
        Me.txt新增姓名.BackColor = System.Drawing.SystemColors.Window
        Me.txt新增姓名.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt新增姓名.Location = New System.Drawing.Point(110, 34)
        Me.txt新增姓名.MaxLength = 4
        Me.txt新增姓名.Name = "txt新增姓名"
        Me.txt新增姓名.Size = New System.Drawing.Size(100, 26)
        Me.txt新增姓名.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label7)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.Label10)
        Me.TabPage2.Controls.Add(Me.Label8)
        Me.TabPage2.Controls.Add(Me.Label9)
        Me.TabPage2.Controls.Add(Me.txt修改重复密码)
        Me.TabPage2.Controls.Add(Me.txt修改新密码)
        Me.TabPage2.Controls.Add(Me.txt修改旧密码)
        Me.TabPage2.Controls.Add(Me.txt修改姓名)
        Me.TabPage2.Controls.Add(Me.btn修改)
        Me.TabPage2.Location = New System.Drawing.Point(4, 26)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(286, 231)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "修改"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(13, 165)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "重复新密码"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(34, 132)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 16)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "新密码"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(107, 69)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(168, 16)
        Me.Label10.TabIndex = 10
        Me.Label10.Text = "请从操作员列表中选择"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(34, 97)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 16)
        Me.Label8.TabIndex = 10
        Me.Label8.Text = "旧密码"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(34, 38)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 16)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "姓  名"
        '
        'txt修改重复密码
        '
        Me.txt修改重复密码.BackColor = System.Drawing.SystemColors.Window
        Me.txt修改重复密码.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt修改重复密码.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txt修改重复密码.Location = New System.Drawing.Point(110, 161)
        Me.txt修改重复密码.MaxLength = 8
        Me.txt修改重复密码.Name = "txt修改重复密码"
        Me.txt修改重复密码.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt修改重复密码.Size = New System.Drawing.Size(100, 26)
        Me.txt修改重复密码.TabIndex = 3
        '
        'txt修改新密码
        '
        Me.txt修改新密码.BackColor = System.Drawing.SystemColors.Window
        Me.txt修改新密码.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt修改新密码.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txt修改新密码.Location = New System.Drawing.Point(110, 127)
        Me.txt修改新密码.MaxLength = 8
        Me.txt修改新密码.Name = "txt修改新密码"
        Me.txt修改新密码.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt修改新密码.Size = New System.Drawing.Size(100, 26)
        Me.txt修改新密码.TabIndex = 2
        '
        'txt修改旧密码
        '
        Me.txt修改旧密码.BackColor = System.Drawing.SystemColors.Window
        Me.txt修改旧密码.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt修改旧密码.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txt修改旧密码.Location = New System.Drawing.Point(110, 93)
        Me.txt修改旧密码.MaxLength = 8
        Me.txt修改旧密码.Name = "txt修改旧密码"
        Me.txt修改旧密码.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt修改旧密码.Size = New System.Drawing.Size(100, 26)
        Me.txt修改旧密码.TabIndex = 1
        '
        'txt修改姓名
        '
        Me.txt修改姓名.BackColor = System.Drawing.SystemColors.Window
        Me.txt修改姓名.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt修改姓名.Location = New System.Drawing.Point(110, 34)
        Me.txt修改姓名.MaxLength = 4
        Me.txt修改姓名.Name = "txt修改姓名"
        Me.txt修改姓名.ReadOnly = True
        Me.txt修改姓名.Size = New System.Drawing.Size(100, 26)
        Me.txt修改姓名.TabIndex = 0
        '
        'btn修改
        '
        Me.btn修改.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btn修改.Location = New System.Drawing.Point(211, 196)
        Me.btn修改.Name = "btn修改"
        Me.btn修改.Size = New System.Drawing.Size(72, 29)
        Me.btn修改.TabIndex = 4
        Me.btn修改.Text = "确认"
        Me.btn修改.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.Label13)
        Me.TabPage3.Controls.Add(Me.txt删除密码)
        Me.TabPage3.Controls.Add(Me.Label14)
        Me.TabPage3.Controls.Add(Me.Label12)
        Me.TabPage3.Controls.Add(Me.Label11)
        Me.TabPage3.Controls.Add(Me.txt删除姓名)
        Me.TabPage3.Controls.Add(Me.btn删除)
        Me.TabPage3.Location = New System.Drawing.Point(4, 26)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(286, 231)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "删除"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(34, 97)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 16)
        Me.Label13.TabIndex = 16
        Me.Label13.Text = "密  码"
        '
        'txt删除密码
        '
        Me.txt删除密码.BackColor = System.Drawing.SystemColors.Window
        Me.txt删除密码.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt删除密码.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txt删除密码.Location = New System.Drawing.Point(110, 93)
        Me.txt删除密码.MaxLength = 8
        Me.txt删除密码.Name = "txt删除密码"
        Me.txt删除密码.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt删除密码.Size = New System.Drawing.Size(100, 26)
        Me.txt删除密码.TabIndex = 1
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(107, 131)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(168, 36)
        Me.Label14.TabIndex = 14
        Me.Label14.Text = "可输入操作员密码或管理员密码"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(107, 69)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(168, 16)
        Me.Label12.TabIndex = 14
        Me.Label12.Text = "请从操作员列表中选择"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(34, 38)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 16)
        Me.Label11.TabIndex = 13
        Me.Label11.Text = "姓  名"
        '
        'txt删除姓名
        '
        Me.txt删除姓名.BackColor = System.Drawing.SystemColors.Window
        Me.txt删除姓名.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt删除姓名.Location = New System.Drawing.Point(110, 34)
        Me.txt删除姓名.Name = "txt删除姓名"
        Me.txt删除姓名.ReadOnly = True
        Me.txt删除姓名.Size = New System.Drawing.Size(100, 26)
        Me.txt删除姓名.TabIndex = 0
        '
        'btn删除
        '
        Me.btn删除.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btn删除.Location = New System.Drawing.Point(211, 196)
        Me.btn删除.Name = "btn删除"
        Me.btn删除.Size = New System.Drawing.Size(72, 29)
        Me.btn删除.TabIndex = 2
        Me.btn删除.Text = "确认"
        Me.btn删除.UseVisualStyleBackColor = True
        '
        'frm操作员
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(473, 336)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lst操作员)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm操作员"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "操作员管理"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lst操作员 As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents txt新增姓名 As System.Windows.Forms.TextBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txt新增重复密码 As System.Windows.Forms.TextBox
    Friend WithEvents txt新增密码 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btn新增 As System.Windows.Forms.Button
    Friend WithEvents btn修改 As System.Windows.Forms.Button
    Friend WithEvents btn删除 As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txt修改重复密码 As System.Windows.Forms.TextBox
    Friend WithEvents txt修改新密码 As System.Windows.Forms.TextBox
    Friend WithEvents txt修改旧密码 As System.Windows.Forms.TextBox
    Friend WithEvents txt修改姓名 As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txt删除密码 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txt删除姓名 As System.Windows.Forms.TextBox
End Class
