Imports System.Data.OleDb

Public Structure prmDistribute
    Dim second As Byte
    Dim minute As Byte
    Dim hour As Byte
    Dim pos() As Byte
    Dim type As Byte
    Dim max As Byte
    Dim mini As Byte
End Structure
Module pubdeclare
    Public _DBconn As New OleDbConnection("Provider=Microsoft.Ace.OleDb.12.0;Data Source=D:/老化台.accdb")
    Public _管理员密码 As String = "65108280"


End Module
