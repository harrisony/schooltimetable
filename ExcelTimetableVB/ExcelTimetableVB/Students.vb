Imports MSExcel = Microsoft.Office.Interop.Excel
Imports System.IO
Module Students
    Sub Main()
        Console.WriteLine("Creating Database")
        Call SQL.createnewdb()
        Console.WriteLine("Classes")
        Call classes()
        Console.WriteLine("Students")
        Call students()
        Console.Write("Students and Classes")
        Call matchstudentswithclasses()
    End Sub
    Sub classes()
        db.Open()
        Dim classes As New ArrayList
        For row As Integer = 2 To Excel.students.Rows.Count ' skip the titles row
            Using dbquery As New SQLite.SQLiteCommand(db)
                dbquery.CommandText = "INSERT INTO Classes VALUES(@code,@fullname,@course,@units)"
                Dim k(3) As SQLite.SQLiteParameter
                k(0) = New SQLite.SQLiteParameter("@code", Excel.students.Cells(row, 5).Value) ' "11ENGADVA"
                k(1) = New SQLite.SQLiteParameter("@fullname", Excel.students.Cells(row, 6).Value) ' "English Advanced 2 Unit" 
                k(3) = New SQLite.SQLiteParameter("@units", CStr(Excel.students.Cells(row, 7).Value) + CStr(Excel.students.Cells(row, 8).Value))
                If k(1).Value.Split()(0) = "IB" Then
                    k(2) = New SQLite.SQLiteParameter("@course", "IB")
                Else
                    k(2) = New SQLite.SQLiteParameter("@course", "HSC")
                End If
                If Not classes.Contains(k(0).Value) Then
                    classes.Add(k(0).Value)
                    dbquery.Parameters.AddRange(k)
                    dbquery.ExecuteNonQuery()
                    Console.WriteLine(vbTab & k(0).Value)
                End If
            End Using
        Next row
        db.Close()
    End Sub
    Sub students()
        db.Open()
        Dim students As New ArrayList
        For row As Integer = 2 To Excel.students.Rows.Count
            Using dbquery As New SQLite.SQLiteCommand(db)
                dbquery.CommandText = "INSERT INTO Students VALUES(@stunumber,@surname,@cname,@house)"
                Dim k(3) As SQLite.SQLiteParameter
                k(0) = New SQLite.SQLiteParameter("@stunumber", Excel.students.Cells(row, 1).Value)
                k(1) = New SQLite.SQLiteParameter("@surname", Excel.students.Cells(row, 2).Value)
                k(2) = New SQLite.SQLiteParameter("@cname", Excel.students.Cells(row, 3).Value)
                k(3) = New SQLite.SQLiteParameter("@house", Excel.students.Cells(row, 4).Value)

                ' Do we need to add year here?
                If Not students.Contains(k(0).Value) And Not k(3).Value = "NEW" Then ' NEW students already have houses
                    students.Add(k(0).Value)
                    Console.WriteLine(vbTab & String.Format("{0},{1},{2},{3}", k(0).Value, k(1).Value, k(2).Value, k(3).Value))
                    dbquery.Parameters.AddRange(k)
                    dbquery.ExecuteNonQuery()
                End If
            End Using
            Using dbquery As New SQLite.SQLiteCommand(db)
                dbquery.CommandText = "INSERT INTO StudentsClasses VALUES(@stunumber,@class)"
                Dim k(1) As SQLite.SQLiteParameter
                k(0) = New SQLite.SQLiteParameter("@stunumber", Excel.students.Cells(row, 1).Value)
                k(1) = New SQLite.SQLiteParameter("@class", Excel.students.Cells(row, 5).Value)
                Console.WriteLine(vbTab & String.Format("{0} {1}", CStr(k(0).Value), k(1).Value))
                dbquery.Parameters.AddRange(k)
                dbquery.ExecuteNonQuery()
            End Using
        Next row
        db.Close()
    End Sub
End Module
