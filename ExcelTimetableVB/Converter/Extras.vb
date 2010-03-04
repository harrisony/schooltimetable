Module SQL
    Public db As New SQLite.SQLiteConnection("data source=students.db")
    Sub createnewdb()
        '' DEBUG ONLY ''
        Dim fi As New System.IO.FileInfo("students.db")
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
Module Excel
    Dim xlApp As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.ApplicationClass
    Public students As Microsoft.Office.Interop.Excel.Range _
    = xlApp.Workbooks.Open("C:\Users\harrisony\Downloads\Current_Yr11_Student_Subjects.xls", , True) _
      .Worksheets("Current_Yr11_Student_Subjects").UsedRange
    Public timetable As Microsoft.Office.Interop.Excel.Workbook _
    = xlApp.Workbooks.Open("C:\Users\harrisony\Downloads\Mater TT Term 1  4 Feb.xls", , True)


End Module