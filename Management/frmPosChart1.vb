Imports LeoControls

Public Class frmPosChart1
    Dim cell(47) As OneCell

    Private Sub frmPosChart1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For j = 0 To 3
            For i = 0 To 11
                cell(12 * j + i) = New OneCell
                cell(12 * j + i).Left = 30 + 100 * i
                cell(12 * j + i).Top = 40 + 140 * j
                cell(12 * j + i).Parent = Me
                cell(12 * j + i).CellLabel = (12 * j + i + 1).ToString
            Next
        Next


    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        For i = 0 To 47
            cell(i).Reset()
        Next
    End Sub

    Private Sub btnAuto_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAuto.Click
        Dim a As String

        For i = 0 To 47
            If cell(i).isUsed Then
                a = cell(i).CellNum
                Exit For
            End If
        Next

        For i = 0 To 47
            If cell(i).isUsed Then
                cell(i).CellNum = a
                a += 1
            End If
        Next
    End Sub
End Class