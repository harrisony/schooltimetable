Imports MSExcel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports System.IO
Module Timetable
    Sub Main()
        Dim yearandrows As Dictionary(Of Integer, Array) ' A mapping of the Year level to how many rows
        Dim outputfile As StreamWriter = New StreamWriter("timetable.txt")
        For Each xlWorkSheet As MSExcel.Worksheet In Excel.timetable.Worksheets
            Dim xlRange As MSExcel.Range = xlWorkSheet.UsedRange
            yearandrows = MatchYearsAndRows(xlRange)
            Dim day As String = GetDate(xlRange)

            Console.WriteLine(day)
            outputfile.WriteLine(String.Format("DAY:{0}", day))

            For col As Integer = 2 To (xlRange.Columns.Count - 1) Step 3
                Dim period As String = (xlRange.Cells(3, col).Value) ' "Period 1"
                Console.WriteLine(vbTab & period)
                For Each year As Integer In yearandrows.Keys
                    Dim q As Array = yearandrows(year) ' an array with the row it starts from and number of classes
                    outputfile.WriteLine(String.Format("YER:{0}", year))
                    For row As Integer = q(0) To (q(0) + q(1))
                        If (xlRange.Cells(row, col).Value <> vbNullString) Then
                            Dim aclass As New SClass
                            aclass.classid = xlRange.Cells(row, col).Value
                            aclass.classroom = xlRange.Cells(row, col + 1).Value
                            aclass.teacher = xlRange.Cells(row, col + 2).Value
                            Select Case period
                                Case "Before School"
                                    aclass.period = "A"
                                Case "Lunch"
                                    aclass.period = "Lunch"
                                Case "After School"
                                    aclass.period = "C"
                                Case Else
                                    aclass.period = period.Split(" ")(1)
                            End Select

                            outputfile.WriteLine(String.Format("PER:{0}", aclass.tofile()))

                        End If
                    Next row
                Next year
            Next col
        Next xlWorkSheet
        outputfile.Close()
    End Sub
    Function MatchYearsAndRows(ByVal spreadsheet As MSExcel.Range) As Dictionary(Of Integer, Array)
        Dim c1, c2 As Integer
        Dim d As New Dictionary(Of Integer, Array)
        For row As Integer = 4 To Excel.timetable.Rows.Count
            If Excel.timetable.Cells(row, 1).Value <> "" Then
                c2 = c1
                c1 = row
                'MsgBox(Excel.timetable.Cells(currentrow, 1).Value)
                If c2 <> 0 And c1 <> 0 Then
                    Dim k(1) As Integer
                    k(0) = c2
                    k(1) = c1 - c2 - 1
                    d.Add((Excel.timetable.Cells(c2, 1).Value), k)
                    c2 = 0
                End If
            End If
        Next
        Return d
    End Function
    Private Function GetDate(ByVal spreadsheet As MSExcel.Range) As String
        Return Regex.Split(spreadsheet.Cells(2, 1).Value, "  ")(2)
    End Function
End Module

Public Class SClass
    Public classid As String
    Public classroom As String
    Public teacher As String
    Public period As String
    Function tofile()
        Return String.Format("{0},{1},{2},{3}", classid, classroom, teacher, period)
    End Function
End Class