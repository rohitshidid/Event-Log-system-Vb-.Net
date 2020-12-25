
Imports System.Data.OleDb

Public Class view_data
    Dim conn As New OleDbConnection
    Dim comm As New OleDb.OleDbCommand

    Dim rank As Integer = 0
    Dim abs As Integer



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        pass_view.Close()

        Form1.Show()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        conn = New OleDb.OleDbConnection
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = my_db.accdb"
        conn.Open()
        comm.Connection = conn
        'aaaaaaaaaaaaaaaaaaaa
        see_the_data()

    End Sub

    Private Sub view_data_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn = New OleDb.OleDbConnection
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = my_db.accdb"
        conn.Open()
        comm.Connection = conn
        RD()

        Label18.Text = My.Settings.pass

    End Sub

    Public Sub BindGrid()
        conn = New OleDb.OleDbConnection
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = my_db.accdb"
        conn.Open()
    End Sub


    Public Sub RD()
        Dim dt As New DataTable()
        Dim ds As New DataSet
        ds.Tables.Add(dt)

        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("SELECT * FROM log_table", conn)
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        conn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        End

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If rank < (abs - 1) Then
            rank = rank + 1
            see_the_data()
        Else
            MessageBox.Show("We are at the bottom!")
        End If



    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If rank > 0 Then
            rank = rank - 1
            see_the_data()
        Else
            MessageBox.Show("We are on the top!")
        End If
    End Sub


    Public Sub see_the_data()

        Dim dt As New DataTable()
        Dim ds As New DataSet
        ds.Tables.Add(dt)

        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("SELECT * FROM log_table where [Roll NO] = '" & TextBox11.Text & "'", conn)
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        conn.Close()

        If dt.Rows.Count > 0 Then

            ComboBox1.Text = dt.Rows(rank)(1).ToString
            TextBox8.Text = dt.Rows(rank)(2).ToString
            TextBox7.Text = dt.Rows(rank)(3).ToString
            TextBox13.Text = dt.Rows(rank)(4).ToString
            ComboBox2.Text = dt.Rows(rank)(5).ToString
            TextBox9.Text = dt.Rows(rank)(6).ToString
            TextBox2.Text = dt.Rows(rank)(7).ToString
            TextBox1.Text = dt.Rows(rank)(8).ToString
            TextBox3.Text = dt.Rows(rank)(9).ToString
            TextBox4.Text = dt.Rows(rank)(10).ToString
            TextBox5.Text = dt.Rows(rank)(11).ToString
            DateTimePicker1.Text = dt.Rows(rank)(12).ToString
            DateTimePicker2.Text = dt.Rows(rank)(13).ToString
            ComboBox4.Text = dt.Rows(rank)(14).ToString
        End If


        abs = dt.Rows.Count
    End Sub


    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            ToolTip1.Show("This is not a number", sender, 5000)
            e.KeyChar = Nothing
        End If
    End Sub


    ' Seach View Code

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If ComboBox3.Text = "Roll No" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [Roll No] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "Semester" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [Semester] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "Phone No" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [Phone No] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "Email" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [Email] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "Description" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [Description] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "Gender" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [Gender] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "P_Name" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [P_Name] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "E_Venue" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [E_Venue] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "E_Name" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [E_Name] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "Level" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [Level] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "College Name" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [College Name] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "Sponsors" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [Sponsors] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "S_Date" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [S_Date] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If

        If ComboBox3.Text = "E_Date" Then
            conn = New OleDb.OleDbConnection
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my_db.accdb"
            Dim da As New OleDbDataAdapter("SELECT * FROM log_table where [E_Date] Like'" + TextBox6.Text + "%' order by [Roll No] asc", conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            Me.DataGridView1.DataSource = dt
        End If




    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim dt As New DataTable()
        Dim ds As New DataSet
        ds.Tables.Add(dt)

        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("SELECT * FROM log_table", conn)
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        conn.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        My.Settings.pass = TextBox10.Text
        Label18.Text = My.Settings.pass
        MsgBox("Password Changed", MsgBoxStyle.Exclamation)
        TextBox10.Text = ""

    End Sub
End Class