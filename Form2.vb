Public Class Form2
    Private Sub butConfig_Click(sender As Object, e As EventArgs) Handles butConfig.Click
        Dim configPwrd = My.Settings.Config_Pwrd
        If txtConfigPasswrd.Text = configPwrd Then
            Me.txtConfigPasswrd.Text = ""
            Me.Hide()
        Else
            MsgBox("Incorrect password!", MsgBoxStyle.Exclamation)
            Me.txtConfigPasswrd.Text = ""
            Form1.TabControl1.SelectedIndex = 0
            Me.Hide()
        End If
    End Sub

    Private Sub Form2_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form1.TabControl1.SelectedIndex = 0
    End Sub
End Class