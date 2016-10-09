Imports ADOX
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports Microsoft.Office

Public Class frmAddNewBD
    Dim pmPath As String = My.Settings.pm_Path 'Application settings
    Dim pmData As String = My.Settings.pm_File
    'Public myConnToAccess As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & pmPath & pmData & ";Extended Properties=""Excel 12.0;HDR=YES""")
    Public myConnToAccess As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & pmPath & pmData)


    Private Sub butAddBD_Click(sender As Object, e As EventArgs) Handles butAddBD.Click
        Try

            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim newBDtxt As String = txtAddNewBD.Text
            myConnToAccess.Open()
            'myCommand.Connection = myConnToAccess
            'insert new equipment into PMList.

            'Rev 6 Change
            'Dim insSQL As String = "INSERT INTO [Fault] (Eqp_ID,Fault,Start_Date,Eq_Status,Breakdown) VALUES('" & lblNewRec.Text & "'" &
            '    ",'" & cmbFaultCat.SelectedValue.ToString & "','" & dtStart.Value.ToShortDateString & "','B-DOWN','" & newBDtxt & "')"
            Dim insSQL As String = "INSERT INTO [Fault] (Eqp_ID,Fault,Start_Date,Eq_Status,Breakdown) VALUES('" & lblNewRec.Text & "'" &
                ",'" & cmbFaultCat.SelectedValue.ToString & "','" & dtStart.Value.ToString("d/M/yyyy") & "','B-DOWN','" & newBDtxt & "')"

            Dim myCommand As New OleDbCommand(insSQL, myConnToAccess)
            myCommand.ExecuteNonQuery()
            myConnToAccess.Close()

            myConnToAccess.Open()
            Dim updateSQL As String = "UPDATE [GageMasterEntry] SET Eqp_Status='B-DOWN' WHERE Eqp_ID='" & lblNewRec.Text & "'"
            myCommand = New OleDbCommand(updateSQL, myConnToAccess)
            myCommand.ExecuteNonQuery()
            myConnToAccess.Close()

            MsgBox("New Breakdown updated for: " & lblNewRec.Text & ". Thank you!")

            'get support ID and send notification email
            Dim ds_SelSupport As DataSet
            Dim da As OleDbDataAdapter
            'Dim supTbls As DataTableCollection
            Dim idSql As String = "SELECT [Master_Lookup].Notify_ID,[Master_Lookup].Eqp_Line from [Master_Lookup] LEFT JOIN [GageMasterEntry] ON" &
                " [GageMasterEntry].Eqp_Line=[Master_Lookup].Eqp_Line WHERE [GageMasterEntry].Eqp_ID='" & lblNewRec.Text & "'"
            ds_SelSupport = New DataSet
            'supTbls = ds_SelSupport.Tables
            da = New OleDbDataAdapter(idSql, myConnToAccess)
            da.Fill(ds_SelSupport, "Master_Lookup")
            myConnToAccess.Close()
            Dim supID As String
            Dim supLine As String
            If ds_SelSupport.Tables(0).Rows.Count > 0 Then
                supID = ds_SelSupport.Tables(0)(ds_SelSupport.Tables(0).Rows.Count - 1)("Notify_ID").ToString
                supLine = ds_SelSupport.Tables(0)(ds_SelSupport.Tables(0).Rows.Count - 1)("Eqp_Line").ToString
                'MsgBox(supID)
                If supID = "" Then
                    MsgBox("No Support ID found. Please contact ME for more info. Notification email cannot be sent out.")
                    Exit Sub
                End If
            End If
            Dim supSubj As String = "Equipment Breakdown Notification"
            Dim supBody As String = "Equipment: " & lblNewRec.Text & Environment.NewLine &
                "Breakdown Category: " & cmbFaultCat.SelectedValue.ToString & Environment.NewLine &
                "Breakdown Description: " & newBDtxt


            Call setEmailSend(supSubj, supBody, supID, "", "", "")


            txtAddNewBD.Text = ""
            System.Windows.Forms.Cursor.Current = Cursors.Default

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & " Error at add New breakdown Event." & Environment.NewLine &
                   " Please capture screen shot And contact Aik Koon. Thank You.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default
            txtAddNewBD.Text = ""
            Me.Close()
        End Try
    End Sub

    Private Sub frmAddNewBD_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dim ds As DataSet
        Dim dA As OleDbDataAdapter
        Dim Tables As DataTableCollection
        Dim bkDwnID = Form3.lblBrkDwnID.Text
        Dim getBDSQL = "Select ID,Fault,Start_Date,Eq_Status,Svc_Cost FROM [Fault] WHERE Eqp_ID='" & bkDwnID & "'"

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
            MsgBox(ex.Message & Environment.NewLine & "Error at Add new breakdown record event." & Environment.NewLine &
                "Please capture screen shot And notify Aik Koon. Thanks.")
        End Try
    End Sub

    Private Sub frmAddNewBD_HelpButtonClicked(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.HelpButtonClicked
        helpAddBDRec.Show()
    End Sub

    Private Sub setEmailSend(sSubject As String, sBody As String,
                             sTo As String, sCC As String,
                             sFilename As String, sDisplayname As String)
        Dim oApp As Interop.Outlook._Application
        oApp = New Interop.Outlook.Application

        Dim oMsg As Interop.Outlook._MailItem
        oMsg = oApp.CreateItem(Interop.Outlook.OlItemType.olMailItem)

        oMsg.Subject = sSubject
        oMsg.Body = sBody

        oMsg.To = sTo
        oMsg.CC = sCC


        Dim strS As String = sFilename
        Dim strN As String = sDisplayname
        If sFilename <> "" Then
            Dim sBodyLen As Integer = Int(sBody.Length)
            Dim oAttachs As Interop.Outlook.Attachments = oMsg.Attachments
            Dim oAttach As Interop.Outlook.Attachment

            oAttach = oAttachs.Add(strS, , sBodyLen, strN)

        End If

        oMsg.Send()
        MessageBox.Show("Email Send", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

    End Sub


End Class