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
        Console.WriteLine("Creating Database")
        Call createnewdb()
        Console.WriteLine("Classes")
        Call classes()
        Console.WriteLine("Students")
        Call students()
        Console.Write("Students and Classes")
        Call matchstudentswithclasses()
    End Sub
    Sub classes()
        Dim db As New SQLite.SQLiteConnection("data source=students.db")
        db.Open()
        Dim classes As New ArrayList
        For row As Integer = 2 To (xlRange.Rows.Count) ' skip the titles row

            Dim dbquery As New SQLite.SQLiteCommand(db)
            dbquery.CommandText = "INSERT INTO Classes VALUES(@code,@fullname,@course,@units)"
            Dim k(3) As SQLite.SQLiteParameter
            k(0) = New SQLite.SQLiteParameter("@code", xlRange.Cells(row, 5).Value) ' "11ENGADVA"
            k(1) = New SQLite.SQLiteParameter("@fullname", xlRange.Cells(row, 6).Value) ' "English Advanced 2 Unit" 
            k(3) = New SQLite.SQLiteParameter("@units", CStr(xlRange.Cells(row, 7).Value) + CStr(xlRange.Cells(row, 8).Value))
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
        Next row
        db.Close()
    End Sub
    Sub students()
        Dim db As New SQLite.SQLiteConnection("data source=students.db")
        db.Open()
        Dim students As New ArrayList
        For row As Integer = 2 To (xlRange.Rows.Count)
            Dim dbquery As New SQLite.SQLiteCommand(db)
            dbquery.CommandText = "INSERT INTO Students VALUES(@stunumber,@surname,@cname,@house)"
            Dim k(3) As SQLite.SQLiteParameter
            k(0) = New SQLite.SQLiteParameter("@stunumber", xlRange.Cells(row, 1).Value)
            k(1) = New SQLite.SQLiteParameter("@surname", xlRange.Cells(row, 2).Value)
            k(2) = New SQLite.SQLiteParameter("@cname", xlRange.Cells(row, 3).Value)
            k(3) = New SQLite.SQLiteParameter("@house", xlRange.Cells(row, 4).Value)

            ' Do we need to add year here?
            If Not students.Contains(k(0).Value) And Not k(3).Value = "NEW" Then ' NEW students already have houses
                students.Add(k(0).Value)
                Console.WriteLine(vbTab & String.Format("{0},{1},{2},{3}", k(0).Value, k(1).Value, k(2).Value, k(3).Value))
                dbquery.Parameters.AddRange(k)
                dbquery.ExecuteNonQuery()
            End If
        Next row
        db.Close()
    End Sub
    Sub matchstudentswithclasses()
        Dim db As New SQLite.SQLiteConnection("data source=students.db")
        db.Open()
        For row As Integer = 2 To (xlRange.Rows.Count)
            Dim dbquery As New SQLite.SQLiteCommand(db)
            dbquery.CommandText = "INSERT INTO StudentsClasses VALUES(@stunumber,@class)"
            Dim k(1) As SQLite.SQLiteParameter
            k(0) = New SQLite.SQLiteParameter("@stunumber", xlRange.Cells(row, 1).Value)
            k(1) = New SQLite.SQLiteParameter("@class", xlRange.Cells(row, 5).Value)
            Console.WriteLine(vbTab & String.Format("{0} {1}", CStr(k(0).Value), k(1).Value))
            dbquery.Parameters.AddRange(k)
            dbquery.ExecuteNonQuery()
        Next row
        db.Close()
    End Sub
    Sub createnewdb()
        '' DEBUG ONLY ''
        Dim fi As New FileInfo("students.db")
        fi.Delete()
        '' DEBUG ONLY ''
        ' I don't like this section
        Dim fileexist As Boolean
        Try
            GetAttr("students.db")
            fileexist = True
        Catch ex As Exception
            fileexist = False
        End Try
        If Not fileexist Then
            Dim db As New SQLite.SQLiteConnection("data source=students.db")
            db.Open()
            Dim k As New SQLite.SQLiteCommand(db)
            k.CommandText = "CREATE TABLE [Students] ( [ComputerNumber] integer PRIMARY KEY NOT NULL, [Surname] text NOT NULL, [CName] text NOT NULL, [House] text NOT NULL)"
            k.ExecuteNonQuery()
            k.CommandText = "CREATE TABLE [Classes] ([Code] TEXT  NULL PRIMARY KEY,[Name] TEXT  NULL,[Course] TEXT  NULL,[Units] TEXT  NULL)"
            k.ExecuteNonQuery()
            k.CommandText = "CREATE TABLE [StudentsClasses] ( [ComputerNumber] INTEGER  NULL,[Class] TEXT  NULL)"
            k.ExecuteNonQuery()
            db.Close()
        End If

    End Sub
End Module
