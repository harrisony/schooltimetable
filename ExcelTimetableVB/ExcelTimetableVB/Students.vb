Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Module Students
    Public xlApp As Excel.Application
    Public xlWorkBook As Excel.Workbook
    Public xlRange As Excel.Range
    Sub Main()
        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\harrisony\Downloads\Current_Yr11_Student_Subjects.xls")
        xlRange = xlWorkBook.Worksheets("Current_Yr11_Student_Subjects").UsedRange
        Call classes()
    End Sub
    Sub classes()
        Dim outputfile As StreamWriter = New StreamWriter("classes.txt")
        Dim classes As New ArrayList
        For row As Integer = 2 To (xlRange.Rows.Count) ' skip the titles row
            Dim code As String = xlRange.Cells(row, 5).Value
            Dim fullname As String = xlRange.Cells(row, 6).Value
            If Not classes.Contains(code) Then
                classes.Add(code)
                outputfile.WriteLine(String.Format("{0},{1}", code, fullname))
            End If

        Next row
        outputfile.Close()
    End Sub

    Sub all()
        Dim outputfile As StreamWriter = New StreamWriter("students.txt")
        For row As Integer = 2 To (xlRange.Rows.Count) ' skip the titles row
            For col As Integer = 1 To (xlRange.Columns.Count)
                If col = xlRange.Columns.Count Then
                    ' last column we don't want a comma
                    outputfile.Write(xlRange.Cells(row, col).Value)
                Else
                    outputfile.Write(xlRange.Cells(row, col).Value & ",")
                End If
            Next
            outputfile.Write(vbNewLine)
        Next
        outputfile.Close()
    End Sub


End Module
