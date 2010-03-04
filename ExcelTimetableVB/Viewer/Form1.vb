Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim db As SQLite.SQLiteConnection = New SQLite.SQLiteConnection("data source=students.db")
        db.Open()
        Using query As New SQLite.SQLiteCommand(db)
            query.CommandText = "SELECT Class, Classes.Name from StudentsClasses JOIN Classes on (StudentsClasses.Class=Classes.Code) WHERE ComputerNumber = @compnumber"
            query.Parameters.Add(New SQLite.SQLiteParameter("@compnumber", TextBox1.Text))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            ListBox1.Items.Clear()
            If response.HasRows = True Then
                While response.Read()
                    ListBox1.Items.Add(String.Format("{0} {1}", response.GetValue(0), response.GetValue(1)))
                End While
            Else
                ListBox1.Items.Add("No student found")
            End If
        End Using

        Using query As New SQLite.SQLiteCommand(db)
            Label1.Text = vbNullString
            query.CommandText = "SELECT Surname, CName from Students WHERE ComputerNumber = @compnumber"
            query.Parameters.Add(New SQLite.SQLiteParameter("@compnumber", TextBox1.Text))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            Label1.Text = String.Format("{0} {1}", response.GetValue(1), response.GetValue(0))
        End Using


    End Sub

End Class
