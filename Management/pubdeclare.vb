Imports System.Data.OleDb
Imports System.IO.Ports

Public Class prmDistribute
    Public second As Byte
    Public minute As Byte
    Public hour As Byte
    Public week As Byte
    Public day As Byte
    Public month As Byte
    Public year As Byte
    Public pos(11) As Byte
    Public type As Byte
    Public max As Byte
    Public mini As Byte
End Class

Public Class commFlag
    Public polling As Boolean
    Public startup As Boolean
    Public integral As Boolean
    Public timeModify As Boolean
    Public reboot340 As Boolean
    public reboot As Boolean
    Public unitNo As Byte
End Class

Module pubdeclare
    Public _DBconn As New OleDbConnection("Provider=Microsoft.Ace.OleDb.12.0;Data Source=" & Application.StartupPath & "/db/老化台.accdb")
    Public _管理员密码 As String = "65108280"
    Public _unit(23) As LHUnit
    Public _readBuffer(32) As Byte
    Public _commFlag As New commFlag
    Public _rts前沿 As Integer = 2
    Public _rts后沿 As Integer = 10
End Module
