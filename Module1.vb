Imports Microsoft.Office.Interop
Module Module1
    Public Enum ColumnVisibility
        Hide
        Show

    End Enum
    Public Sub setColumnVisibility(ByVal filename As String, ByVal visibility As ColumnVisibility)
        Dim xlApp As Excel.Application = New Excel.Application
        Dim xlWb As Excel.Workbook
        Dim xlPreviousActiveSheet As Excel.Worksheet
        Dim xlSheet As Excel.Worksheet
        Dim xlRng As Excel.Range
        'start Excel and get Application object
        xlApp = CreateObject("Excel.Application")
        'change to False if you don't
        'want Excel to be visible
        xlApp.Visible = True
        'open workbook
        xlWb = xlApp.Workbooks.Open(filename)
        'get previously active sheet
        'so we can make it the active sheet
        'again before we close the file
        xlPreviousActiveSheet = xlWb.ActiveSheet
        For i As Integer = 1 To xlWb.Sheets.Count
            xlSheet = xlApp.Sheets(i)
            'activate current sheet
            'needed for "Select"
            xlSheet.Activate()
            'get range of sheet
            xlRng = xlSheet.Cells()
            'Console.WriteLine("Total Rows: " & xlRng.Rows.Count)
            'Console.WriteLine("Total Columns: " & xlRng.Columns.Count)
            'select range
            xlSheet.Range(xlSheet.Cells(1, 10), xlSheet.Cells(xlRng.Rows.Count, xlRng.Columns.Count)).Select()
            If visibility = ColumnVisibility.Show Then
                'show columns in range
                xlApp.Selection.EntireColumn.Hidden = False
            Else
                'hide columns in range
                xlApp.Selection.EntireColumn.Hidden = True
            End If
            'undo selection
            'xlApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
            xlSheet.Range("A1").Select()
        Next
        'make previous active sheet
        'the active sheet again
        'before we close the file
        xlPreviousActiveSheet.Activate()
        'close and save changes
        'xlWb.Close(SaveChanges:=True)
        'quit Excel
        xlApp.Quit()
    End Sub
End Module

