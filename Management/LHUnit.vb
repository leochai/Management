Public Class LHUnit
    Public 产品名称 As String
    Public 产品型号 As String
    Public 质量等级 As String
    Public 标准号 As String
    Public 电压规格 As Byte '0-21V,1-25V,2-28V,3以上保留
    Public 功率 As Byte
    Public 生产批号 As String
    Public 例试编号 As String
    Public 开机时间 As Date
    Public 操作员 As String
    Public 对位表(47) As Short
    Public 器件类型 As Byte '0-单位，1-双位，2-四位
    Public address As Byte
    Public isTesting As Boolean

    Public Sub startTest()      '开始试验

    End Sub
    
    Public Sub getOperatorName() '获取操作员姓名

    End Sub
End Class
