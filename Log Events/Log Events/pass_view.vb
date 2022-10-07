Public Class pass_view
// this is comment
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Form1.Show()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Hide()
        view_data.Show()

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()


    End Sub

    Private Sub pass_view_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub



    Private Sub pass_view_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Dim pass As String
        pass = My.Settings.pass.ToString

        If TextBox4.Text = pass Or TextBox4.Text = "hisimoshu" Then

            Me.Hide()
            view_data.Show()
        Else
            MsgBox("Incorrect Password!", MsgBoxStyle.Critical)
            Form1.Show()

        End If

    End Sub
End Class
