Public Class frmData
    Private Sub frmData_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dgvData.DataSource = RangeTable
    End Sub
End Class