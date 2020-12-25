
Imports System.Data.OleDb

Public Class add_log

    Dim conn As New OleDbConnection
    Dim comm As New OleDb.OleDbCommand

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox4.Text = "" Or TextBox8.Text = "" Or ComboBox3.Text = "" Or TextBox9.Text = "" Or TextBox11.Text = "" Then
            MsgBox("Cannot leave the Mandatory boxes empty !", MsgBoxStyle.Exclamation)
        Else

            If TextBox8.Text.Length < 10 Then
                MsgBox("Phone No Incomplete", MsgBoxStyle.Exclamation)
            Else

                If TextBox11.Text.Length < 6 Then
                    MsgBox("Roll No Incomplete", MsgBoxStyle.Exclamation)
                Else


                    If TextBox11.Text = TextBox11.Text Then
                        comm.Connection = conn
                        comm.CommandText = "Insert into log_table([Roll No],[Semester],[Phone No],[Email],[Description],[Gender],[P_Name],[E_Venue],[E_Name],[Level],[College Name],[Sponsors],[S_Date],[E_Date],[Ranking]) VALUES ('" & Me.TextBox11.Text & "','" & Me.ComboBox1.Text & "','" & Me.TextBox8.Text & "','" & Me.TextBox7.Text & "','" & Me.TextBox13.Text & "','" & Me.ComboBox2.Text & "','" & Me.TextBox9.Text & "','" & Me.TextBox2.Text & "','" & Me.TextBox1.Text & "','" & Me.TextBox3.Text & "','" & Me.TextBox4.Text & "','" & Me.TextBox5.Text & "','" & Me.DateTimePicker1.Text & "','" & Me.DateTimePicker2.Text & "','" & Me.ComboBox3.Text & "')"
                        comm.ExecuteNonQuery()

                        MsgBox("Submitted Succesfully", MsgBoxStyle.MsgBoxRight)
                    End If
                End If
            End If


        End If

    End Sub

    Private Sub add_log_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn = New OleDb.OleDbConnection
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = my_db.accdb"
        conn.Open()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And
            TextBox4.Text = "" And
            TextBox5.Text = "" And
            TextBox7.Text = "" And
            TextBox8.Text = "" And
        TextBox9.Text = "" And
            TextBox11.Text = "" And
            TextBox13.Text = "" And
            ComboBox1.Text = "" And
            ComboBox2.Text = "" Then


            MessageBox.Show("All Boxes Already Cleared! ")
        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox11.Text = ""
            TextBox13.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            MessageBox.Show("All Boxes cleared now! ")
        End If
    End Sub



    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            ToolTip1.Show("This is not a number", sender, 5000)
            e.KeyChar = Nothing
        End If
    End Sub


    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            ToolTip1.Show("This is not a number", sender, 5000)
            e.KeyChar = Nothing
        End If
    End Sub
End Class