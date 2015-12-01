Imports System.Data.OleDb
Imports System.IO.Ports

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
    Public _unit(23) As LHUnit
    Public _readBuffer(29) As Byte
    Public _pollingFlag As Boolean
    Public _startupFlag As Byte
    Public _distributeFlag As Byte
    Public _integralFlag As Boolean
End Module
