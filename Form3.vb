Imports ADOX
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class Form3
    Dim pmPath As String = My.Settings.pm_Path 'Application settings
    Dim pmData As String = My.Settings.pm_File
    Dim uID As Boolean = False
    'Public myConnToAccess As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & pmPath & pmData & ";Extended Properties=""Excel 12.0;HDR=YES""")
    Public myConnToAccess As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & pmPath & pmData)


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridBrkDwn.CellClick
        Dim ds_Update As DataSet
        Dim Tbls As DataTableCollection
        Dim da As New OleDbDataAdapter

        If DataGridBrkDwn.SelectedRows(0).Cells(0).Value.ToString = "" Then
            Exit Sub
        End If

        If Me.DataGridBrkDwn.SelectedRows.Count > 1 Then
            MsgBox("Please select only 1 equipment.")
            Exit Sub
        End If

        txtFaultSymptom.Text = ""
        txtSoln.Text = ""

        Dim brkDwnID As String = lblBrkDwnID.Text
        Dim recID As Integer = DataGridBrkDwn.SelectedRows(0).Cells(0).Value
        Dim brkDwnDate As String = DataGridBrkDwn.SelectedRows(0).Cells(2).Value
        Dim dateCult As String = Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
        Dim MPos As Integer = InStr(dateCult, "M")
        If MPos = 1 Then 'US MM/dd/yyyy format
            brkDwnDate = Mid(brkDwnDate, InStr(brkDwnDate, "/") + 1, Len(brkDwnDate) - 4 - (InStr(brkDwnDate, "/") + 1)) & "/" &
                        Mid(brkDwnDate, 1, InStr(brkDwnDate, "/") - 1) & "/" & Mid(brkDwnDate, Len(brkDwnDate) - 3, 4)
        End If
        Dim formatDate As String = brkDwnDate.ToString
        Dim bdSQL As String = "SELECT Breakdown, Solutions FROM [Fault] WHERE Eqp_ID='" & brkDwnID & "' AND ID=" & recID & ""
        ds_Update = New DataSet
        Tbls = ds_Update.Tables
        da = New OleDbDataAdapter(bdSQL, myConnToAccess)
        da.Fill(ds_Update, "Fault")
        Dim brkDwnInfo As String = Tbls(0).Rows(0).Item("Breakdown").ToString
        Dim brkDwnSoln As String = Tbls(0).Rows(0).Item("Solutions").ToString
        txtFaultSymptom.Text = brkDwnInfo
        txtSoln.Text = brkDwnSoln

    End Sub

    Private Sub butAddBrkDwn_Click(sender As Object, e As EventArgs) Handles butAddBrkDwn.Click
        Try
            Dim dTr As Integer
            Dim recRows As Integer = DataGridBrkDwn.Rows.Count
            If recRows > 0 Then
                For dTr = 0 To recRows - 1
                    If DataGridBrkDwn.Rows(dTr).Cells(3).Value.ToString = "" Then
                        Exit Sub
                    End If
                    If DataGridBrkDwn.Rows(dTr).Cells(3).Value.ToString = "B-DOWN" Then
                        MsgBox(lblBrkDwnID.Text & " has an opened breakdown record. Please close it before adding new records. Thanks!")
                        DataGridBrkDwn.Rows(dTr).Selected = True
                        Exit Sub
                    End If
                Next
            End If
            frmAddNewBD.lblNewRec.Text = Me.lblBrkDwnID.Text
            Dim ds_Type As DataSet = New DataSet
            Dim dT As OleDbDataAdapter
            Dim SQLquery = "SELECT * from [Fault-Type]"
            Dim faultCount As Integer = frmAddNewBD.cmbFaultCat.Items.Count
            If faultCount = 0 Then
                dT = New OleDbDataAdapter(SQLquery, myConnToAccess)
                dT.Fill(ds_Type, "Fault-Type")
                With frmAddNewBD.cmbFaultCat
                    .DataSource = ds_Type.Tables("Fault-Type")
                    .DisplayMember = "Fault_Type"
                    .ValueMember = "Fault_Type"
                    .SelectedIndex = 0
                    .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    .AutoCompleteSource = AutoCompleteSource.ListItems
                End With
                myConnToAccess.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message & ". Please capture screen shot of error and email Aik Koon. Thank you.")
            myConnToAccess.Close()
            Exit Sub
        End Try
        frmAddNewBD.cmbFaultCat.SelectedIndex = -1
        frmAddNewBD.ShowDialog()
    End Sub

    Private Sub butUpdateBrkDwn_Click(sender As Object, e As EventArgs) Handles butUpdateBrkDwn.Click
        Dim recStat As String
        Dim ds_Update As DataSet
        Dim Tbls As DataTableCollection
        Dim da As New OleDbDataAdapter

        If DataGridBrkDwn.RowCount = 0 Then
            MsgBox("There are no breakdown records for this equipment.")
            Exit Sub
            'ElseIf IsNothing(Me.DataGridBrkDwn.SelectedRows) Then
            '    recStat = Me.DataGridBrkDwn.Rows(0).Cells(2).Value.ToString
        Else
            recStat = Me.DataGridBrkDwn.SelectedRows(0).Cells(3).Value.ToString
        End If
        If recStat <> "B-DOWN" Then
            MsgBox("This record is already closed!")
            Exit Sub
        Else
            txtFaultSymptom.Text = ""
            txtSoln.Text = ""
            Try
                Dim brkDwnID As String = lblBrkDwnID.Text
                Dim recID As Integer = DataGridBrkDwn.SelectedRows(0).Cells(0).Value
                Dim brkDwnDate As String = DataGridBrkDwn.SelectedRows(0).Cells(2).Value
                Dim formatDate As String = brkDwnDate.ToString
                Dim bdSQL As String = "SELECT Breakdown, Solutions FROM [Fault] WHERE Eqp_ID='" & brkDwnID & "' AND ID=" & recID & ""
                ds_Update = New DataSet
                Tbls = ds_Update.Tables
                da = New OleDbDataAdapter(bdSQL, myConnToAccess)
                da.Fill(ds_Update, "Fault")
                Dim brkDwnInfo As String = Tbls(0).Rows(0).Item("Breakdown").ToString
                Dim brkDwnSoln As String = Tbls(0).Rows(0).Item("Solutions").ToString
                myConnToAccess.Close()
                txtFaultSymptom.Text = brkDwnInfo
                txtSoln.Text = brkDwnSoln
                bdUpdate.lblUpdateID.Text = Me.lblBrkDwnID.Text
                bdUpdate.lblStartDate.Text = Me.DataGridBrkDwn.SelectedRows(0).Cells(2).Value.ToString
                bdUpdate.lblRecID.Text = Me.DataGridBrkDwn.SelectedRows(0).Cells(0).Value.ToString
                bdUpdate.txtupdateFaultSymptom.Text = Me.txtFaultSymptom.Text
                bdUpdate.txtUpdateSoln.Text = Nothing
                bdUpdate.ShowDialog()
            Catch ex As Exception
                MsgBox(ex.Message & ". Please capture screenshot of error message and email Aik Koon.")
                myConnToAccess.Close()
            End Try
        End If
    End Sub

    Private Sub txtnewSupID_TextChanged(sender As Object, e As EventArgs) Handles txtnewSupID.TextChanged
        Dim str As String = txtnewSupID.Text
        str = str.ToUpper
        txtnewSupID.Text = str
        txtnewSupID.SelectionStart = str.Length
    End Sub

    Private Sub txtnewSupID_Leave(sender As Object, e As EventArgs) Handles txtnewSupID.Leave
        If txtnewSupID.Text <> "" Then
            Dim chkID As Boolean = Trim(txtnewSupID.Text) Like "H*"
            If chkID = False Or Len(Trim(txtnewSupID.Text)) <> 7 Then
                Label32.ForeColor = Color.Red
                MsgBox("The user ID may be invalid. Please correct it to a valid Halliburton User ID.")
                Label32.ForeColor = Color.Black
                uID = False
                Exit Sub
            Else
                uID = True
            End If
        End If
    End Sub

    Private Sub butUpdateNotify_Click(sender As Object, e As EventArgs) Handles butUpdateNotify.Click

        If uID = False Then
            Label32.ForeColor = Color.Red
            MsgBox("The user ID may be invalid. Please correct it to a valid Halliburton User ID.")
            Label32.ForeColor = Color.Black
            Exit Sub
        End If
        Try
            Dim ID As DataSet = New DataSet
            Dim getIDSQL As String = "SELECT Notify_ID FROM [Master_Lookup] WHERE Eqp_Line='" & lblProdLine.Text & "'"
            Dim getID As OleDbDataAdapter = New OleDbDataAdapter(getIDSQL, myConnToAccess)
            getID.Fill(ID, "Master_Lookup")
            Dim nID As String = ID.Tables(0).Rows(0).Item("Notify_ID").ToString
            myConnToAccess.Close()
            Dim concatID As String = nID & ";" & Trim(txtnewSupID.Text)
            Dim addIDSQL As String = "UPDATE [Master_Lookup] SET Notify_ID='" & concatID & "' WHERE Eqp_Line='" & lblProdLine.Text & "'"
            myConnToAccess.Open()
            Dim addID As OleDbCommand = New OleDbCommand(addIDSQL, myConnToAccess)
            addID.ExecuteNonQuery()
            myConnToAccess.Close()

            getIDSQL = "SELECT [Master_Lookup].Notify_ID FROM [Master_Lookup] LEFT JOIN [GageMasterEntry] ON" &
                    " [GageMasterEntry].Eqp_Line=[Master_Lookup].Eqp_Line WHERE [GageMasterEntry].Eqp_ID='" & lblBrkDwnID.Text & "'"
            myConnToAccess.Open()
            Dim dN As DataSet = New DataSet
            Dim dNote As OleDbDataAdapter = New OleDbDataAdapter(getIDSQL, myConnToAccess)
            dNote.Fill(dN, "Master_Lookup")
            DataGridNotify.DataSource = dN
            DataGridNotify.DataMember = "Master_Lookup"
            myConnToAccess.Close()
            MsgBox("User ID: " & Trim(txtnewSupID.Text) & " added to notification.", MsgBoxStyle.Information)
            txtnewSupID.Text = ""

        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at update support ID event." & Environment.NewLine &
                    "Please capture screenshot of error message and email Aik Koon. Thank You.")
            myConnToAccess.Close()
        End Try
    End Sub

    Private Sub Form3_HelpButtonClicked(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.HelpButtonClicked
        HelpForm3.Show()
    End Sub

    Private Sub Form3_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form1.cmbEqp_Name.SelectedIndex = -1
        Form1.cmbProd_Line.SelectedIndex = -1
        Form1.DataGridView1.DataSource = Nothing
    End Sub
End Class