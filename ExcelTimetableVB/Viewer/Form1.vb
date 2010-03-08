Public Class Form1
    Public db As SQLite.SQLiteConnection = New SQLite.SQLiteConnection()

    Private Sub StudentClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim source As String

        If sender.GetType.ToString() = "System.Windows.Forms.ListBox" Then
            source = sender.SelectedItem.Value()

        ElseIf IsNumeric(ComboBox1.Text) Then
            source = ComboBox1.Text
        Else
            source = ComboBox1.SelectedValue
        End If

        Using query As New SQLite.SQLiteCommand(db)
            query.CommandText = "SELECT Class, Classes.Name from StudentsClasses JOIN Classes on (StudentsClasses.Class=Classes.Code) WHERE ComputerNumber = @compnumber"
            query.Parameters.Add(New SQLite.SQLiteParameter("@compnumber", source))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            ListBox1.Items.Clear()
            If response.HasRows = True Then
                While response.Read()
                    Dim test As New cValue
                    test.Display = String.Format("{0} - {1}", response.GetValue(0), response.GetValue(1))
                    test.Value = response.GetValue(0)
                    ListBox1.Items.Add(test)

                End While
            Else
                ListBox1.Items.Add("No student found")
            End If
        End Using
        ComboBox1.SelectedValue = 0  ' clear the combobox to remove confusion


        Using query As New SQLite.SQLiteCommand(db)
            Label1.Text = vbNullString
            query.CommandText = "SELECT Surname, CName from Students WHERE ComputerNumber = @compnumber"
            query.Parameters.Add(New SQLite.SQLiteParameter("@compnumber", source))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            Label1.Text = String.Format("{0} {1} - {2}", response.GetValue(1), response.GetValue(0), source)
        End Using
        ListBox1.Tag = "Student"

    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not db.State = ConnectionState.Open Then
            'TODO:  I'm not keen on the below line, I'd rather we use the existing connection via ADO
            ' It should also use ConnectionStrings("studentsConnectionString1") but it doesn't work for some reason
            db.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings(1).ConnectionString
            db.Open()
        End If
        ' Call adddatasource() 

        Me.ClassesTableAdapter.Fill(Me.ClassesDataSet.Classes)
        ComboBox1.SelectedValue = 0
        ComboBox2.SelectedValue = 0
        Me.StudentsTableAdapter.Fill(Me.StudentsDataSet.Students)
        ComboBox1.SelectedValue = 0
        ComboBox2.SelectedValue = 0
    End Sub

    Private Sub ClassClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim source As String
        If sender.GetType.ToString() = "System.Windows.Forms.ListBox" Then
            source = sender.SelectedItem.Value
        Else
            source = ComboBox2.SelectedValue
        End If

        Using query As New SQLite.SQLiteCommand(db)
            query.CommandText = "SELECT Students.CName || ' ' || Students.Surname as Test, Students.ComputerNumber from StudentsClasses JOIN Students on (StudentsClasses.ComputerNumber = Students.ComputerNumber) WHERE Class = @class ORDER BY Students.Surname ASC"
            query.Parameters.Add(New SQLite.SQLiteParameter("@class", source))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            ListBox1.Items.Clear()
            If response.HasRows = True Then
                While response.Read()
                    Dim test As New cValue
                    test.Value = response.GetValue(1)
                    test.Display = response.GetValue(0)
                    ListBox1.Items.Add(test)
                End While
            Else
            End If
        End Using
        Using query As New SQLite.SQLiteCommand(db)
            Label1.Text = vbNullString
            query.CommandText = "SELECT Name FROM Classes WHERE Code = @code"
            query.Parameters.Add(New SQLite.SQLiteParameter("@code", source))
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            Label1.Text = String.Format("{0} - {1}", response.GetValue(0), source)
            ' Have I mentioned how much I love String.Format()
        End Using
        ListBox1.Tag = "Class"

    End Sub

    Private Sub adddatasource()
        Dim ds As New DataSet
        Dim dt As DataTable
        Dim dr As DataRow
        Dim idCoulumn As DataColumn
        Dim nameCoulumn As DataColumn

        dt = New DataTable()
        idCoulumn = New DataColumn("Code", Type.GetType("System.String"))
        nameCoulumn = New DataColumn("Name", Type.GetType("System.String"))

        dt.Columns.Add(idCoulumn)
        dt.Columns.Add(nameCoulumn)
        Using query As New SQLite.SQLiteCommand(db)
            query.CommandText = "SELECT Name, Code FROM Classes"
            Dim response As SQLite.SQLiteDataReader = query.ExecuteReader()
            ListBox1.Items.Clear()
            If response.HasRows = True Then
                While response.Read()
                    dr = dt.NewRow()
                    dr("Code") = response.GetValue(1)
                    dr("Name") = response.GetValue(0)
                    dt.Rows.Add(dr)
                End While
            Else
            End If
        End Using
        ds.Tables.Add(dt)
        ComboBox2.DataSource = ds.Tables(0)
        ComboBox2.DisplayMember = "Code"
        ComboBox2.ValueMember = "Code"

    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.Tag = "Student" Then
            ClassClick(sender, New EventArgs())
        ElseIf ListBox1.Tag = "Class" Then
            StudentClick(sender, New EventArgs())
        End If

    End Sub
End Class
