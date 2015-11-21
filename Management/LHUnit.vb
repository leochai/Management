Public Class LHUnit
    Public 产品名称 As String
    Public 产品型号 As String
    Public 质量等级 As String
    Public 标准号 As String
    Public 电压 As Single
    Public 功率 As Single
    Public 生产批号 As String
    Public 例试编号 As String
    Public 单元号 As Byte
    Public 开机时间 As Date
    Public 操作员 As String
    Public 对位表(47) As Short


    Public Sub startTest()      '开始试验

    End Sub
    Public Sub resumeTest()     '继续试验

    End Sub
    Public Sub endTest()        '结束试验

    End Sub
    Public Sub getOperatorName() '获取操作员姓名

    End Sub
End Class
