Imports ADOX
Imports System.Data.OleDb
'Imports System.IO

Public Class bdUpdate
    Dim pmPath As String = My.Settings.pm_Path 'Application settings
    Dim chklstPath As String = My.Settings.pm_Checklist
    Dim pmData As String = My.Settings.pm_File
    'Public myConnToAccess As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & pmPath & pmData & ";Extended Properties=""Excel 12.0;HDR=YES""")
    Public myConnToAccess As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & pmPath & pmData)


    Private Sub txtSvcCost_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSvcCost.KeyPress
        Dim FullStop As Char
        FullStop = "."

        ' if the '.' key was pressed see if there already is a '.' in the string
        ' if so, dont handle the keypress
        If e.KeyChar = FullStop And txtSvcCost.Text.IndexOf(FullStop) <> -1 Then
            e.Handled = True
            Return
        End If

        ' If the key aint a digit
        If Not Char.IsDigit(e.KeyChar) Then
            ' verify whether special keys were pressed
            ' (i.e. all allowed non digit keys - in this example
            ' only space and the '.' are validated)
            If (e.KeyChar <> FullStop) And (e.KeyChar <> Convert.ToChar(8)) Then
                ' if its a non-allowed key, dont handle the keypress
                e.Handled = True
                Return
            End If
        End If

    End Sub

    Private Sub butUpdateCloseRec_Click(sender As Object, e As EventArgs) Handles butUpdateCloseRec.Click

        If txtupdateFaultSymptom.Text = "" Then
            MsgBox("Description of the breakdown is required. Thanks!")
            Exit Sub
        ElseIf txtUpdateSoln.Text = "" Then
            MsgBox("A corrective action or solution for breakdown is required to close the record. Thanks!")
            Exit Sub
        End If
        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim brkDwnDesc As String = txtupdateFaultSymptom.Text
            Dim brkDwnSln As String = txtUpdateSoln.Text
            Dim clsDate As String = dtCloseRec.Value.ToString("d/M/yyyy")
            Dim stDate As String = lblStartDate.Text
            Dim clsID As String = lblUpdateID.Text
            Dim clsSQL = "UPDATE [Fault] SET End_Date='" & clsDate & "',Eq_Status='ACTIVE',Breakdown='" & brkDwnDesc & "',Solutions='" & brkDwnSln &
            "',Svc_Cost=" & txtSvcCost.Text & ",Svc_PO='" & txtSvcPO.Text & "' WHERE Eqp_ID='" & clsID & "' AND Eq_Status='B-DOWN'"
            'Dim clsSQL = "UPDATE [Fault] SET End_Date='" & clsDate & "',Eq_Status='ACTIVE' WHERE Eqp_ID='" & clsID & "' AND Eq_Status='B-DOWN'"
            myConnToAccess.Open()
            Dim clsCommand As New OleDbCommand(clsSQL, myConnToAccess)
            clsCommand.ExecuteNonQuery()
            myConnToAccess.Close()

            clsSQL = "UPDATE [GageMasterEntry] SET Eqp_Status='ACTIVE' WHERE Eqp_ID='" & clsID & "'"
            myConnToAccess.Open()
            Dim updatePMcommand As New OleDbCommand(clsSQL, myConnToAccess)
            updatePMcommand.ExecuteNonQuery()
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default

            MsgBox("Update for " & clsID & " done. Breakdown record closed. Thanks.")
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message & ". Please capture screenshot and email Aik Koon. Thanks.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default

            Exit Sub
        End Try

    End Sub

    Private Sub bdUpdate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dim ds As DataSet
        Dim dA As OleDbDataAdapter
        Dim Tables As DataTableCollection
        Dim bkDwnID = Form3.lblBrkDwnID.Text
        Dim getBDSQL = "SELECT ID,Fault,Start_Date,Eq_Status,Svc_Cost FROM [Fault] WHERE Eqp_ID='" & bkDwnID & "'"

        Try

            ds = New DataSet
            Tables = ds.Tables
            dA = New OleDbDataAdapter(getBDSQL, myConnToAccess)
            dA.Fill(ds, "Fault")
            myConnToAccess.Close()

            Form3.DataGridBrkDwn.DataSource = ds
            Form3.DataGridBrkDwn.DataMember = "Fault"
            Form3.txtFaultSymptom.Text = ""
            Form3.txtSoln.Text = ""
            If Application.OpenForms().OfType(Of Form3).Any = False Then
                Form3.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Please capture screen shot and notify Aik Koon. Thanks.")
        End Try
    End Sub

    Private Sub bdUpdate_HelpButtonClicked(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.HelpButtonClicked
        bdUpdatehelp.Show()
    End Sub
End Class