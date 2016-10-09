
Imports ADOX
Imports System.Data.OleDb
Imports System.IO

Public Class Form4
    Dim pmData As String = My.Settings.pm_File
    Dim pmMasterList As String = My.Settings.pm_Masterlist
    Dim pmpath As String = My.Settings.pm_Path
    Public bkDwnID As String
    Public myConnToAccess As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & pmPath & pmData)

    Private Sub chkExpire(sArea As String)
        Dim dgRow As Integer
        Dim dueDate As Date
        Dim datePM As String
        Dim dateCult As String = Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
        Dim Mpos As Integer = InStr(dateCult, "M")
        Dim pmStat As String
        Dim pmFreq As String
        Dim totCount As Integer = Form1.DataGridViewPM.Rows.Count + 1
        Dim totDue As Integer = 0
        Dim totBkDwn As Integer = 0

        For dgRow = 0 To Form1.DataGridViewPM.Rows.Count - 1
            pmStat = Form1.DataGridViewPM.Rows(dgRow).Cells(8).Value.ToString
            pmFreq = Form1.DataGridViewPM.Rows(dgRow).Cells(5).Value.ToString
            If pmStat = "B-DOWN" Then
                Form1.DataGridViewPM.Rows(dgRow).Cells(8).Style.BackColor = Color.Red
                totBkDwn = totBkDwn + 1
            Else
                Form1.DataGridViewPM.Rows(dgRow).Cells(8).Style.BackColor = Color.Green
            End If
            If Not IsDBNull(Form1.DataGridViewPM.Rows(dgRow).Cells(7).Value) Then
                datePM = Form1.DataGridViewPM.Rows(dgRow).Cells(7).Value.ToString
                'Rev6 change***
                If Mpos = 1 Then
                    datePM = Mid(datePM, InStr(datePM, "/") + 1, Len(datePM) - 4 - (InStr(datePM, "/") + 1)) & "/" &
                        Mid(datePM, 1, InStr(datePM, "/") - 1) & "/" & Mid(datePM, Len(datePM) - 3, 4)
                End If
                '****************
                dueDate = DateTime.ParseExact(datePM, dateCult, Globalization.CultureInfo.InvariantCulture)
                If Now > dueDate.AddDays(-7) And Now < dueDate And pmStat <> "B-DOWN" And pmFreq <> "CONDITIONAL" Then
                    Form1.DataGridViewPM.Rows(dgRow).Cells(8).Value = "PM-DUE"
                    Form1.DataGridViewPM.Rows(dgRow).Cells(8).Style.BackColor = Color.Yellow
                ElseIf dueDate < Now And pmFreq <> "CONDITIONAL" Then
                    Form1.DataGridViewPM.Rows(dgRow).Cells(8).Value = "OVERDUED"
                    Form1.DataGridViewPM.Rows(dgRow).Cells(8).Style.BackColor = Color.Red
                    Form1.DataGridViewPM.Rows(dgRow).Cells(7).Style.BackColor = Color.Red
                    totDue = totDue + 1
                ElseIf pmFreq = "CONDITIONAL" Then
                    Form1.DataGridViewPM.Rows(dgRow).Cells(8).Value = "NA"
                End If
            Else
                Form1.DataGridViewPM.Rows(dgRow).Cells(7).Style.BackColor = Color.Red
            End If
        Next

        With Form1
            If .chkPortEqp.Checked = True Then
                .lblTotPMEqm.Text = "Total Portable Eqp in " & sArea & ": " & totCount
                .lblTotBrkDwn.Text = "Total due Portable Eqp in " & sArea & ": " & totDue
                .lblTotDue.Text = ""
            ElseIf .chkPortEqp.Checked = False And sArea <> "WIRELINE" Then
                .lblTotPMEqm.Text = "Total Eqm with PM in " & sArea & ": " & totCount
                .lblTotBrkDwn.Text = "Total Breakdown Eqm in " & sArea & ": " & totBkDwn
                '.lblTotDue.Text = "Total Eqm no PM required in " & sArea & ": " & totInactive
            ElseIf .chkPortEqp.Checked = False And sArea = "WIRELINE" Then
                .lblTotPMEqm.Text = "Total Eqm with PM in " & sArea & ": " & totCount
                .lblTotBrkDwn.Text = "Total Breakdown Eqm in " & sArea & ": " & totBkDwn
                '.lblTotDue.Text = "Total Inactive Eqm in " & sArea & ": " & totInactive
            End If
        End With
    End Sub

    Private Sub butPMDate_Click(sender As Object, e As EventArgs) Handles butPMDate.Click
        Dim selID As String = Label2.Text
        Dim lastPMdate As Date = MonthCalendar1.SelectionStart
        'MsgBox(selDate)
        Try
            'Get the PM frequency info
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'Dim gF As DataSet = New DataSet
            'Dim gFSQL As String = "SELECT Eqp_Freq FROM [GageMasterEntry] WHERE Eqp_ID='" & selID & "'"
            'Dim gFQuery As New OleDbDataAdapter(gFSQL, myConnToAccess)
            'gFQuery.Fill(gF, "GageMasterEntry")
            'myConnToAccess.Close()
            'dtDue.Format = DateTimePickerFormat.Custom
            'dtDue.CustomFormat = "d/M/yyyy"
            'Dim pmFreq As String = Trim(gF.Tables(0)(gF.Tables(0).Rows.Count - 1)("Eqp_Freq")
            Dim pmfreq As String = lblFreq.Text
            Dim nxDue As Date

            Select Case pmFreq
                Case "DAILY"
                    MsgBox("No PM Due date will be tracked")
                    Exit Sub
                Case "WEEKLY"
                    nxDue = lastPMdate.AddDays(7)
                Case "MONTHLY"
                    nxDue = lastPMdate.AddMonths(1)
                Case "QUARTERLY"
                    nxDue = lastPMdate.AddMonths(3)
                Case "HALF-YEARLY"
                    nxDue = lastPMdate.AddMonths(6)
                Case "ANNUALLY"
                    nxDue = lastPMdate.AddYears(1)
                Case "EVERY JOB"
                    MsgBox("No PM Due date will be tracked")
                    Exit Sub
                Case "CONDITIONAL"
                    MsgBox("No PM Due date will be tracked")
                    Exit Sub
            End Select


            'Rev 6 change
            'Dim pmSQL As String = "UPDATE [GageMasterEntry] SET Eqp_Due='" & nxDue.toshortdatestring & "' WHERE Eqp_ID='" & Label2.Text & "'"
            Dim pmSQL As String = "UPDATE [GageMasterEntry] SET Eqp_Due='" & nxDue.ToString("d/M/yyyy") & "' WHERE Eqp_ID='" & Label2.Text & "'"
            myConnToAccess.Open()
            Dim newPM As New OleDbCommand(pmSQL, myConnToAccess)
            newPM.ExecuteNonQuery()
            myConnToAccess.Close()
            MsgBox("PM Due Date updated.")

            'Update PM summary

            Dim selArea As String = Label5.Text
            Dim gD As DataSet = New DataSet
            Dim getPMSQL As String = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
            "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Owner,[GageMasterEntry].Eqp_Due,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] INNER JOIN [Master_Lookup] ON " &
        "[GageMasterEntry].Eqp_Line=[Master_Lookup].Eqp_Line WHERE [Master_Lookup].Prod_Area='" & selArea & "' AND ([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
        " [GageMasterEntry].Eqp_Status='ACTIVE') AND [GageMasterEntry].Eqp_Type='PORTABLE-EQP'"
            Dim gA As New OleDbDataAdapter(getPMSQL, myConnToAccess)
            gA.Fill(gD, "GageMasterEntry")
            Form1.DataGridViewPM.DataSource = gD
            Form1.DataGridViewPM.DataMember = "GageMasterEntry"

            Call chkExpire(selArea)
            System.Windows.Forms.Cursor.Current = Cursors.Default

            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error getting PM summary data." & Environment.NewLine &
                    "Please capture screenshot and email Aik Koon. Thank You.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default

            Me.Close()

        End Try
    End Sub
End Class