Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data

Public Class ExcelDataTableConverter

    Public Shared Function SelectedRangeToDataTable() As DataTable
        Dim selectedRange As Excel.Range = Globals.ThisAddIn.Application.ActiveWindow.Selection 'erstmal die selektion holen
        Dim dataTable As New DataTable() 'dann neue dt dimen


        For i As Integer = 1 To selectedRange.Columns.Count 'für alle spalten auch spalten in der dt definieren und richtig benennen
            dataTable.Columns.Add(String.Format(selectedRange.Cells(1, i).Value, i))
        Next

        For Each row As Excel.Range In selectedRange.Rows 'dann neue zeilen pro excelzeile
            Dim dataRow As DataRow = dataTable.NewRow()

            For i As Integer = 1 To selectedRange.Columns.Count 'zählen ab 1, denn 0 ist überschrift
                dataRow(i - 1) = row.Cells(1, i).Value
            Next

            dataTable.Rows.Add(dataRow)
        Next

        Return dataTable
    End Function

End Class
