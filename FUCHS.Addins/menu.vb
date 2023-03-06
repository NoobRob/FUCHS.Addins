Imports Microsoft.Office.Tools.Ribbon
Imports System.Data
Imports System.Windows.Forms

Public Class menu

    Private Sub menu_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdExportMark_Click(sender As Object, e As RibbonControlEventArgs) Handles cmdExportMark.Click
        Dim dt As DataTable = ExcelDataTableConverter.SelectedRangeToDataTable
        RangeTable = dt
        Dim f As New frmData
        f.ShowDialog()
    End Sub
End Class
