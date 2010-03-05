Public Class Form1
    Public db As SQLite.SQLiteConnection = New SQLite.SQLiteConnection()

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim source As String

        If Not db.State = ConnectionState.Open Then
            'TODO:  I'm not keen on the below line, I'd rather we use the existing connection via ADO
            ' It should also use ConnectionStrings("studentsConnectionString1") but it doesn't work for some reason
            db.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings(1).ConnectionString
            db.Open()
        End If
        If Not Len(TextBox1.Text) = 0 Then
            source = TextBox1.Text
            ComboBox1.SelectedValue = 0  ' clear the combobox to remove confusion
        Else
            source = ComboBox1.SelectedValue.ToString()
        End If
        Using query As New SQLite.SQLiteCommand(db)
            query.CommandText = "SELECT Class, Classes.Name from StudentsClasses JOIN Classes on (StudentsClasses.Class=Classes.Code) WHERE ComputerNumber = @compnumber"
            query.Parameters.Add(New SQLite.SQLiteParameter("@compnumber", source))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            ListBox1.Items.Clear()
            If response.HasRows = True Then
                While response.Read()
                    ListBox1.Items.Add(String.Format("{0} - {1}", response.GetValue(0), response.GetValue(1)))
                End While
            Else
                ListBox1.Items.Add("No student found")
            End If
        End Using

        Using query As New SQLite.SQLiteCommand(db)
            Label1.Text = vbNullString
            query.CommandText = "SELECT Surname, CName from Students WHERE ComputerNumber = @compnumber"
            query.Parameters.Add(New SQLite.SQLiteParameter("@compnumber", source))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            Label1.Text = response.GetValue(1) & " " & response.GetValue(0)
        End Using


    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'StudentsDataSet.Students' table. You can move, or remove it, as needed.
        If Not db.State = ConnectionState.Open Then
            'TODO:  I'm not keen on the below line, I'd rather we use the existing connection via ADO
            ' It should also use ConnectionStrings("studentsConnectionString1") but it doesn't work for some reason
            db.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings(1).ConnectionString
            db.Open()
        End If
        Using query As New SQLite.SQLiteCommand(db)
            query.CommandText = "SELECT Code from Classes"
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            ComboBox2.Items.Clear()
            If response.HasRows = True Then
                While response.Read()
                    ComboBox2.Items.Add(response.GetValue(0))
                End While
            Else
            End If
        End Using
        Me.StudentsTableAdapter.Fill(Me.StudentsDataSet.Students)

    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Using query As New SQLite.SQLiteCommand(db)
            query.CommandText = "SELECT Students.CName || ' ' || Students.Surname as Test from StudentsClasses JOIN Students on (StudentsClasses.ComputerNumber = Students.ComputerNumber) WHERE Class = @class ORDER BY Students.Surname ASC"
            query.Parameters.Add(New SQLite.SQLiteParameter("@class", ComboBox2.SelectedItem))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            ListBox1.Items.Clear()
            If response.HasRows = True Then
                While response.Read()
                    ListBox1.Items.Add(response.GetValue(0))
                End While
            Else
            End If
        End Using
        Using query As New SQLite.SQLiteCommand(db)
            Label1.Text = vbNullString
            query.CommandText = "SELECT Name FROM Classes WHERE Code = @code"
            query.Parameters.Add(New SQLite.SQLiteParameter("@code", ComboBox2.SelectedItem))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            Label1.Text = response.GetValue(0)
        End Using

    End Sub
End Class
