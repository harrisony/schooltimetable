Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Module Students
    Public xlApp As Excel.Application
    Public xlWorkBook As Excel.Workbook
    Public xlRange As Excel.Range
    Sub Main()
        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\harrisony\Downloads\Current_Yr11_Student_Subjects.xls", , True)
        xlRange = xlWorkBook.Worksheets("Current_Yr11_Student_Subjects").UsedRange
        Console.WriteLine("Classes")
        Call classes()
        Console.WriteLine("Students")
        Call students()
        Console.Write("Students and Classes")
        Call matchstudentswithclasses()
    End Sub
    Sub classes()
        Dim outputfile As StreamWriter = New StreamWriter("classes.txt")
        Dim classes As New ArrayList
        For row As Integer = 2 To (xlRange.Rows.Count) ' skip the titles row
            Dim code As String = xlRange.Cells(row, 5).Value
            Dim fullname As String = xlRange.Cells(row, 6).Value
            Dim course As String
            If fullname.Split()(0) = "IB" Then
                course = "IB"
            Else
                course = String.Format("HSC{0}{1}", xlRange.Cells(row, 7).Value, xlRange.Cells(row, 8).Value)
            End If
            If Not classes.Contains(code) Then
                classes.Add(code)
                outputfile.WriteLine(String.Format("{0},{1},{2}", code, fullname, course))
            End If
        Next row
        outputfile.Close()
    End Sub
    Sub students()
        Dim outputfile As StreamWriter = New StreamWriter("students.txt")
        Dim students As New ArrayList
        For row As Integer = 2 To (xlRange.Rows.Count)
            Dim stunumber As Integer = xlRange.Cells(row, 1).Value
            Dim surname As String = xlRange.Cells(row, 2).Value
            Dim cname As String = xlRange.Cells(row, 3).Value
            Dim house As String = xlRange.Cells(row, 4).Value
            ' Do we need to add year here?
            If Not students.Contains(stunumber) And Not house = "NEW" Then ' NEW students already have houses
                students.Add(stunumber)
                Dim q As String = String.Format("{0},{1},{2},{3}", stunumber, surname, cname, house)
                outputfile.WriteLine(q)
                Console.WriteLine(q)
            End If
        Next row
    End Sub
    Sub matchstudentswithclasses()
        Dim outputfile As StreamWriter = New StreamWriter("studentandclass.txt")
        Dim sandc As New Dictionary(Of Integer, ArrayList)

        For row As Integer = 2 To (xlRange.Rows.Count)
            If Not sandc.Keys.Contains(xlRange.Cells(row, 1).Value) Then
                sandc.Add((xlRange.Cells(row, 1).Value), New ArrayList)
            End If
            sandc(xlRange.Cells(row, 1).Value).Add(xlRange.Cells(row, 5).Value)
        Next row

        For Each item As KeyValuePair(Of Integer, ArrayList) In sandc
            outputfile.Write(String.Format("{0},{1}", item.Key, Chr(34)))
            For i As Integer = 0 To item.Value.Count - 1
                If i = item.Value.Count - 1 Then
                    outputfile.WriteLine(item.Value(i) & Chr(34)) ' Chr(34) = "
                Else
                    outputfile.Write(item.Value(i) & ",")
                End If
            Next i
        Next item
        outputfile.Close()
    End Sub
End Module
