Imports LeoControls

Public Class frmPosChart2
    Dim cell(23) As TwoCell

    Private Sub frmPosChart2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For j = 0 To 3
            For i = 0 To 5
                cell(6 * j + i) = New TwoCell
                cell(6 * j + i).Left = 30 + 200 * i
                cell(6 * j + i).Top = 40 + 140 * j
                cell(6 * j + i).Parent = Me
                cell(6 * j + i).CellLabel = (6 * j + i + 1).ToString
            Next
        Next
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        For i = 0 To 23
            cell(i).Reset()
        Next
    End Sub

    Private Sub btnAuto_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAuto.Click
        Dim a As String

        For i = 0 To 23
            If cell(i).isUsed Then
                If cell(i).CellNum = "" Then
                    a = 1
                Else
                    a = cell(i).CellNum
                End If
                Exit For
            End If
        Next

        For i = 0 To 23
            If cell(i).isUsed Then
                cell(i).CellNum = a
                a += 1
            End If
        Next
    End Sub
End Class