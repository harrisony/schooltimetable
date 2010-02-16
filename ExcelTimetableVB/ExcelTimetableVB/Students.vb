Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports System.IO
Module Students
    Public xlApp As Excel.Application
    Public xlWorkBook As Excel.Workbook
    Public xlRange As Excel.Range
    Public xlWorksheet As Excel.Worksheet
    Sub Main()
        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\harrisony\Downloads\Current_Yr11_Student_Subjects.xls")
        xlWorkSheet = xlWorkBook.Worksheets("Current_Yr11_Student_Subjects")
        xlRange = xlWorksheet.UsedRange
        Dim outputfile As StreamWriter = New StreamWriter("students.txt")
        For row As Integer = 2 To (xlRange.Rows.Count) ' skip the titles row
            For col As Integer = 1 To (xlRange.Columns.Count)
                outputfile.Write(xlRange.Cells(row, col).Value & ",")
            Next
            outputfile.Write(vbNewLine)
        Next
        outputfile.Close()
    End Sub
End Module
