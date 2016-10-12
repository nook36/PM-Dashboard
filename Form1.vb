'**************************************************************************************
'** PM-Dashboard VB.Net version                                                      **
'** This is converted from previous Excel VBA version                                **
'** Commenced conversion development: 30th Dec 2015                                  **
'** Completed:                                                                       **
'** ---------------------------------------------------------------------------------**
'** | Rev | Changes                                                     |Chg Date   |**
'** ---------------------------------------------------------------------------------**
'** | 0   | Initial                                                     |           |**
'** | 5   | Correct alphanumeric sorting problem in Eqp_ID to get the   | 19-Feb-16 |**
'** |     | correct last used ID No. for the next new ID to use.        |           |**
'** | 6   | i.Form4b line 93. Chg Toshortdatestring -> Tostring         | 03-Mar-16 |**
'** |     | ii.bdAddNew.vb line 27 same chg as above                    | 04-Mar-16 |**
'** |     | iii. Add line 1743-1746 for US date format                  | 06-Mar-16 |**
'** | 7   | Correct CurrentUICulture to CurrentCulture date format      | 08-Mar-16 |**
'** | 8   | Do nothing for DataGridViewPM click event. Stop PM due date | 08-Mar-16 |**
'** |     | Remove PM Due function in PM Summary and add Inactive Status|           |**
'** |     | Chg PM tab to PSL View tab                                  |           |**
'** |     | Log version info for user Ln74                              |           |**
'** |     | Added export for PSL View summary and 25 char limit for desc| 09-Mar-16 |**
'** | 9   | Line 1028 chg default eqp status to "ACTIVE" for new add eqp| 16-Mar-16 |**
'** |     | Add Cal equip only summary in PSL View tab                  |           |**
'** | 10  | Correct clr form in update/search tab where mfg txtbox is   | 16-Apr-16 |**
'** |     | cleared after update. Also added in 'PM_Need' field to      |           |**
'** |     | differentiate between cal only equip and PM needed equip    |           |**
'** |     | for PM masterlist printing purpose. Allow cal eqm have eqp_ |           |**
'** |     | line info which previously identfys PM equipment.           |           |**
'** | 11  | Corrected column numbering for checklist printout           | 26-Apr-16 |**
'** |     | Corrected error in query for equipment type combo           |           |**
'** | 13  | Added in portable equipment checks                          | 11-May-16 |**
'** |     | ##** Need to correct the equip id for those that have chmbr |           |**
'** | 14  | Corrected error in freq in datagridview for torqemasters.   | 07-Sep-16 |**
'** |     | Corrected master list freq error for torquemasters.         |           |**
'** |     | Remove torquemaster update function.                        |           |**
'** |     | Corrected eqp info mis-align in daily PM-checklist sheet.   |           |**
'** |     | Remove the selective cmb sub. no need to clr after each sel |           |**
'** | 15  | Corrected problem with updating notification ID in form3    | 03-Oct-16 |**
'** |     |                                                             |           |**
'** --------------------------------------------------------------------------------|**
'**************************************************************************************
'Path: \\corp.halliburton.com\ap\SIN\SG2\Public\SING2 MANUFACTURING ENGINEERING\PM\VB-PM\


Imports ADOX
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.ComponentModel
Imports System
Imports Microsoft.Office.Interop.Excel
Imports System.Globalization

Public Class Form1
    Dim SQLquery As String
    Dim da As New OleDbDataAdapter
    Dim dT As New OleDbDataAdapter
    Dim tables As DataTableCollection
    Dim dbName As String
    Dim dbPath As String
    Dim userID As String = GetUserName()

    Dim uPath As String = "C:\Users\" & userID & "\Desktop\"
    Dim pmPath As String = My.Settings.pm_Path 'Application settings
    Dim chklstPath As String = My.Settings.pm_Checklist
    'Dim pmData As String = "PM-Masterlist.xlsm"
    'Dim pmData As String = "PM-Masterlist - Copy.xlsm"
    Dim pmData As String = My.Settings.pm_File
    Dim pmMasterList As String = My.Settings.pm_Masterlist
    Public bkDwnID As String
    Public myConnToAccess As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & pmPath & pmData)


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.Text = "PM_Dashboard Ver." & Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            My.Settings.pm_DateFormat = System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern().ToString
            'label if this is currently running in dev mode with db on local hdd. if not, check if network db is presence and if not throw file missing exception.

            '************************************************************************************************************************************************************************************
            '*****************For access 'accdb' file access*************************************************************************************************************************************
            '*dbPath = "C:\Users\hb90747\Desktop\MDB\"
            '*dbName = "PM-Masterlist.accdb"
            '*
            '*SQLquery = "SELECT DISTINCT Eqp_Line from GageMasterEntry WHERE Eqp_Line IS NOT NULL AND (Eqp_Status<>'B-DOWN' OR Eqp_Status<>'REMOVE')"
            '*
            '*Call SQL_Query(SQLquery, "cmbProd_Line", "GageMasterEntry", 0)
            '*myConnToAccess = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & dbName)
            '************************************************************************************************************************************************************************************

            'myConnToAccess = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=P:\SING2 MANUFACTURING ENGINEERING\PM\PM-Masterlist.xlsm;Extended Properties=""Excel 12.0;HDR=YES""")
            'myConnToAccess = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=C:\Users\hb90747\Desktop\PM\PM-Masterlist.xlsm;Extended Properties=""Excel 12.0;HDR=YES""")
            myConnToAccess.Open()
            Dim ldSQL = "INSERT INTO [UserLog] (Users_Log,Date_Stamp,Ver_No) VALUES ('" & Environment.UserName & "','" & Now.ToShortDateString & " " & Now.ToShortTimeString & "'" &
                ",'" & Me.Text & "')"
            Dim ldQuery As OleDbCommand = New OleDbCommand(ldSQL, myConnToAccess)
            ldQuery.ExecuteNonQuery()
            myConnToAccess.Close()

            Call tab0_Init()
            System.Windows.Forms.Cursor.Current = Cursors.Default

        Catch ex As IO.FileNotFoundException
            MessageBox.Show("You may be offline. Please check connection And try again.")
            System.Windows.Forms.Cursor.Current = Cursors.Default

            myConnToAccess.Close()
        Catch ex As Exception
            If lblNetStatus.Text = "Offline" Then
                MsgBox("You're not connected to network.")
            End If
            MsgBox(ex.Message & Environment.NewLine & " Please contact Aik Koon If error repeats after several tries.")
            myConnToAccess.Close()
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 4 Then
            Call clr_Tab0()
            Call clrTab1()
            Call clr_Tab2()
            'clr tab 3
            Call clr_Tab3()
            Form2.tabIdx.Text = 4
            Form2.ShowDialog()
        ElseIf TabControl1.SelectedIndex = 1 Then
            Call clr_Tab0()
            Call clr_Tab2()
            'clr tab 3
            Call clr_Tab3()
            Call clr_Tab4()
            Form2.tabIdx.Text = 1
            Form2.ShowDialog()
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If cmbNewEqpLine.SelectedIndex = -1 Then
                If cmbNewEqpLine.Items.Count = 0 Then
                    Call tab1_Init("Eqp_Line1", cmbNewEqpLine)
                End If
                cmbNewEqpLine.SelectedValue = -1
            End If
            If cmbNewEqpType.SelectedIndex = -1 Then
                If cmbNewEqpType.Items.Count = 0 Then
                    Call tab1_Init("Eqp_Type", cmbNewEqpType)
                End If
                cmbNewEqpType.SelectedValue = -1
            End If
            Call addFreq(cmbNewFreq1)
            Call addFreq(cmbNewFreq2)
            Call addFreq(cmbNewFreq3)
            Call addFreq(cmbNewFreq4)
            Call addResp(cmbNewResp1)
            Call addResp(cmbNewResp2)
            Call addResp(cmbNewResp3)
            Call addResp(cmbNewResp4)
            System.Windows.Forms.Cursor.Current = Cursors.Default

        ElseIf TabControl1.SelectedIndex = 2 Then
            'Call clr_Tab0()
            'Call clrTab1()
            'Call clr_Tab3()
            'Call clr_Tab4()
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If cmbUSEqpDesc.Items.Count = 0 Then
                Call tab1_Init("Eqp_Type", cmbUSEqpDesc)
                cmbUSEqpDesc.SelectedValue = -1
            End If
            If cmbUSEqp_Line.Items.Count = 0 Then
                Call tab1_Init("Eqp_Line", cmbUSEqp_Line)
                cmbUSEqp_Line.SelectedValue = -1
            End If
            If cmbUSEqpStat.Items.Count = 0 Then
                Call tab1_Init("Stat_Type", cmbUSEqpStat)
                cmbUSEqpStat.SelectedValue = -1
            End If
            If cmbUSFreq1.Items.Count = 0 Then
                Call addFreq(cmbUSFreq1)
                cmbUSFreq1.SelectedItem = -1
            End If
            If cmbUSResp1.Items.Count = 0 Then
                Call addResp(cmbUSResp1)
                cmbUSResp1.SelectedItem = -1
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default

        ElseIf TabControl1.SelectedIndex = 3 Then
            Call clr_Tab0()
            Call clrTab1()
            Call clr_Tab2()
            Call clr_Tab4()
            chkCal.Checked = False
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If cmbEqpArea.Items.Count = 0 Then
                Dim prdAreaSQL As String = "SELECT DISTINCT Prod_Area FROM [Master_Lookup]"
                Dim pA As DataSet = New DataSet
                Dim pAQuery As OleDbDataAdapter = New OleDbDataAdapter(prdAreaSQL, myConnToAccess)
                pAQuery.Fill(pA, "Master_Lookup")

                With cmbEqpArea
                    '.Items.Insert(0, String.Empty) 'Insert empty row to 1st line
                    .DataSource = pA.Tables("Master_Lookup")
                    .DisplayMember = "Prod_Area"
                    .ValueMember = "Prod_Area"
                    .SelectedIndex = -1
                    .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    .AutoCompleteSource = AutoCompleteSource.ListItems
                End With
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default

        ElseIf TabControl1.SelectedIndex = 0 Then
            Call clrTab1()
            Call clr_Tab2()
            Call clr_Tab3()
            Call clr_Tab4()
        End If

    End Sub

    Private Sub clr_Tab4()
        txtAddNewFault.Text = ""
        txtaddNewStat.Text = ""
        txtBkUpFolder.Text = ""
        DataGridViewdbInfo.DataSource = Nothing
    End Sub
    Private Sub clr_Tab2()
        txtUSEqpID.Text = ""
        txtUSEqpID.ReadOnly = False
        cmbUSEqpStat.SelectedIndex = -1
        cmbUSEqpDesc.SelectedIndex = -1
        cmbUSEqp_Line.SelectedIndex = -1
        txtUSMfgSN.Text = ""
        txtUSModel.Text = ""
        txtUSmfg.Text = ""
        txtUSOldID.Text = ""
        txtEqpOwner.Text = ""
        cmbUSFreq1.SelectedIndex = -1
        cmbUSResp1.SelectedIndex = -1
        If DataGridPM.Rows.Count > 0 Then
            DataGridPM.DataSource = Nothing
        End If
        If chkWithBkDwn.Checked = True Then chkWithBkDwn.Checked = False
    End Sub
    Private Sub clr_Tab3()
        If DataGridViewPM.Rows.Count > 0 Then
            DataGridViewPM.DataSource = Nothing
            cmbEqpArea.SelectedIndex = -1
            chkPortEqp.Checked = False
            chkCal.Checked = False
        End If
    End Sub
    Private Sub addFreq(cmbNewFreq As ComboBox)
        'add interval items to new freq
        Dim cIdx As Integer = cmbNewFreq.Items.Count
        If cIdx = 0 Then
            cmbNewFreq.Items.Add("DAILY")
            cmbNewFreq.Items.Add("WEEKLY")
            cmbNewFreq.Items.Add("MONTHLY")
            cmbNewFreq.Items.Add("QUARTERLY")
            cmbNewFreq.Items.Add("HALF-YEARLY")
            cmbNewFreq.Items.Add("ANNUALLY")
            cmbNewFreq.Items.Add("CONDITIONAL")
            cmbNewFreq.Items.Add("EVERY JOB")
        End If

    End Sub
    Private Sub addResp(cmbNewResp As ComboBox)
        'add interval items to new Resp
        Dim cIdx As Integer = cmbNewResp.Items.Count
        If cIdx = 0 Then
            cmbNewResp.Items.Add("INTERNAL")
            cmbNewResp.Items.Add("EXTERNAL")
        End If
    End Sub

    Private Sub cmbProd_Line_Selectionchangecommitted(sender As Object, e As EventArgs) Handles cmbProd_Line.SelectionChangeCommitted
        Dim selLine As Object
        Dim selEqp As Object
        Dim prodLine As String
        Dim ds As DataSet

        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            selLine = cmbProd_Line.SelectedValue
            prodLine = selLine.ToString
            If cmbEqp_Name.SelectedIndex <> -1 Then
                selEqp = cmbEqp_Name.SelectedValue
                SQLquery = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Status, Old_ID,Eqp_Resp,Eqp_Owner,PM_Need from [GageMasterEntry] WHERE Eqp_Line='" & selLine.ToString() &
                "' AND (Eqp_Status='B-DOWN' OR Eqp_Status='ACTIVE') AND Eqp_Type='" & selEqp.ToString() & "'"
            Else
                SQLquery = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Status, Old_ID,Eqp_Resp,Eqp_Owner,PM_Need from [GageMasterEntry] WHERE Eqp_Line='" & selLine.ToString() &
                "' AND (Eqp_Status='B-DOWN' OR Eqp_Status='ACTIVE')"
            End If

            ds = New DataSet
            tables = ds.Tables
            da = New OleDbDataAdapter(SQLquery, myConnToAccess)
            da.Fill(ds, "GageMasterEntry")
            myConnToAccess.Close()
            DataGridView1.DataSource = ds
            DataGridView1.DataMember = "GageMasterEntry"
            DataGridView1.Focus()
            releaseObject(selLine)
            releaseObject(selEqp)
            If prodLine = "SPERRY/ALD-CTN-COLLAR" Or prodLine = "SPERRY/TRIAGE" Or prodLine = "SPERRY/GEOPILOT" Then
                Call cycDTGrid()
            End If
            'Call tab0_cmbSelec("Eqp_Type", cmbEqp_Name, "Eqp_Line", selLine)
            System.Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at Prod_Line selection change." & Environment.NewLine &
                   "Please capture screen shot and notify Aik Koon. Thanks.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub cycDTGrid()
        'sub to cycle through datagrid to search for torquemaster
        Dim tmFreq As String
        Dim hyphenPos As Integer
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells.Item("Eqp_Type").Value Like "TORQUEMASTER-*" Then
                tmFreq = row.Cells(1).Value.ToString
                hyphenPos = InStr(1, tmFreq, "-", 1)
                tmFreq = Microsoft.VisualBasic.Right(tmFreq, Len(tmFreq) - hyphenPos)
                'MsgBox(tmFreq)
                If tmFreq = "HALFYEAR" Then
                    row.Cells.Item("Eqp_Freq").Value = "HALF-YEARLY"
                ElseIf tmFreq Like "*-MONTHLY*" Then
                    row.Cells.Item("Eqp_Freq").Value = "MONTHLY"
                Else
                    row.Cells.Item("Eqp_Freq").Value = tmFreq
                End If
                'rowindex = row.Index.ToString()
                '                Dim actie As String = row.Cells("PRICE").Value.ToString()
                'Exit For
            End If
        Next
    End Sub
    Private Sub cmbEqp_Name_Selectionchangecommitted(sender As Object, e As EventArgs) Handles cmbEqp_Name.SelectionChangeCommitted
        Dim selType As Object
        Dim selLine As Object
        Dim prodLine As String
        Dim ds As DataSet

        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            selType = cmbEqp_Name.SelectedValue
            If cmbProd_Line.SelectedIndex <> -1 Then
                selLine = cmbProd_Line.SelectedValue
                prodline = selLine.ToString
                SQLquery = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Status,Old_ID,Eqp_Resp,Eqp_Owner, PM_Need from [GageMasterEntry] WHERE Eqp_Type='" & selType.ToString() &
                "' AND (Eqp_Status='ACTIVE' OR Eqp_Status='B-DOWN') AND Eqp_Line='" & selLine.ToString() & "'"
            Else
                SQLquery = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Status,Old_ID,Eqp_Resp,Eqp_Owner,PM_Need from [GageMasterEntry] WHERE Eqp_Type='" & selType.ToString() &
                "' AND (Eqp_Status='ACTIVE' OR Eqp_Status='B-DOWN')"
            End If
            ds = New DataSet
            tables = ds.Tables
            da = New OleDbDataAdapter(SQLquery, myConnToAccess)
            da.Fill(ds, "GageMasterEntry")
            myConnToAccess.Close()
            DataGridView1.DataSource = ds
            DataGridView1.DataMember = "GageMasterEntry"
            DataGridView1.Focus()
            releaseObject(selType)
            releaseObject(selLine)
            If prodLine <> "" And (prodLine = "SPERRY/ALD-CTN-COLLAR" Or prodLine = "SPERRY/TRIAGE" Or prodLine = "SPERRY/GEOPILOT") Then
                Call cycDTGrid()
            End If
            'Call tab0_cmbSelec("Eqp_Line", cmbProd_Line, "Eqp_Type", selType)
            System.Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at Eqp_Name Selection Change." & Environment.NewLine &
                   "Please capture screen shot and notify Aik Koon. Thanks.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub butMasterList_Click(sender As Object, e As EventArgs) Handles butMasterList.Click
        Dim xlApp As Excel.Application
        Dim xlWB As Excel.Workbook
        Dim xlWS As Excel.Worksheet

        If cmbProd_Line.SelectedIndex = -1 Then
            Label2.ForeColor = Color.Red
            MsgBox("Please select a production line.")
            Label2.ForeColor = Color.Black
            Exit Sub
        End If
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("No Records retrieved for selected production line. Please try again.")
            Exit Sub
        End If
        Try
            'get masterlist sheet name
            Dim gM As DataSet = New DataSet
            Dim selLine As String = cmbProd_Line.SelectedValue.ToString
            Dim gMSQL = "SELECT PM_Sht FROM [Master_Lookup] WHERE Eqp_Line='" & selLine & "'"
            Dim gMQuery As New OleDbDataAdapter(gMSQL, myConnToAccess)
            gMQuery.Fill(gM, "Master_Lookup")
            Dim pmSht As String = gM.Tables(0)(gM.Tables(0).Rows.Count - 1)("PM_Sht")

            If Not (File.Exists(uPath & pmMasterList)) Then
                My.Computer.FileSystem.CopyFile(pmPath & pmMasterList, uPath & pmMasterList)
            End If

            xlApp = New Excel.Application
            xlWB = xlApp.Workbooks.Open(pmPath & pmMasterList)
            xlApp.Visible = True
            xlWS = xlWB.Sheets(1)
            xlWS.Name = pmSht
            Dim cellRow As Integer = 4

            If DataGridView1.RowCount > 0 Then
                DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)
                Dim rW As Integer
                Dim chkPM As String
                Dim chkSht As String
                For rW = 0 To DataGridView1.RowCount - 1
                    chkPM = DataGridView1.Rows(rW).Cells(11).Value.ToString
                    chkSht = DataGridView1.Rows(rW).Cells(1).Value.ToString
                    'MessageBox.Show(rW.Cells(0).Value.ToString)
                    If chkPM = "YES" Then
                        cellRow = cellRow + 1
                        xlWS.Cells(cellRow, 1) = DataGridView1.Rows(rW).Cells(1).Value.ToString 'Eqp_Type
                        xlWS.Cells(cellRow, 2) = DataGridView1.Rows(rW).Cells(0).Value.ToString 'Eqp_ID
                        xlWS.Cells(cellRow, 3) = DataGridView1.Rows(rW).Cells(2).Value.ToString 'Eqp_Sn
                        xlWS.Cells(cellRow, 4) = DataGridView1.Rows(rW).Cells(4).Value.ToString 'Eqp_Mod
                        xlWS.Cells(cellRow, 5) = DataGridView1.Rows(rW).Cells(5).Value.ToString 'Eqp_Line
                        If chkSht Like "*-DAILY" Then
                            xlWS.Cells(cellRow, 6) = "DAILY" 'Eqp_Freq
                        ElseIf chkSht Like "*-WEEKLY" Then
                            xlWS.Cells(cellRow, 6) = "WEEKLY" 'Eqp_Freq
                        ElseIf chkSht Like "*-QUARTERLY" Then
                            xlWS.Cells(cellRow, 6) = "QUARTERLY" 'Eqp_Freq
                        ElseIf chkSht Like "*-MONTHLY" Then
                            xlWS.Cells(cellRow, 6) = "MONTHLY" 'Eqp_Freq
                        ElseIf chkSht Like "*-MONTHLY*" Then
                            xlWS.Cells(cellRow, 6) = "MONTHLY" 'Eqp_Freq
                        ElseIf chksht Like "*-HALFYEAR" Then
                            xlWS.Cells(cellRow, 6) = "HALF-YEARLY" 'Eqp_Freq
                        ElseIf chkSht Like "*-ANNUALLY" Then
                            xlWS.Cells(cellRow, 6) = "ANNUALLY" 'Eqp_Freq
                        ElseIf chkSht Like "*-CONDITIONAL" Then
                            xlWS.Cells(cellRow, 6) = "CONDITIONAL" 'Eqp_Freq
                        ElseIf chkSht Like "*-EVERYJOB" Then
                            xlWS.Cells(cellRow, 6) = "EVERY JOB" 'Eqp_Freq
                        Else
                            xlWS.Cells(cellRow, 6) = DataGridView1.Rows(rW).Cells(6).Value.ToString 'Eqp_Freq
                        End If
                        xlWS.Cells(cellRow, 7) = DataGridView1.Rows(rW).Cells(9).Value.ToString 'Eqp_Resp
                            xlWS.Cells(cellRow, 8) = DataGridView1.Rows(rW).Cells(7).Value.ToString 'Eqp_Status
                            xlWS.Cells(cellRow, 9) = DataGridView1.Rows(rW).Cells(8).Value.ToString 'Ol_ID
                        End If
                Next

                If xlWS.Columns("I").hidden = False Then
                    xlWS.Columns("I").hidden = True
                End If
            End If

            releaseObject(xlWS)
            releaseObject(xlWB)
            releaseObject(xlApp)
        Catch ex As Exception
            releaseObject(xlWS)
            releaseObject(xlWB)
            releaseObject(xlApp)
            MsgBox(ex.Message & Environment.NewLine & "Error at Masterlist view Event." & Environment.NewLine &
                   "Please capture screen shot And notify Aik Koon. Thanks.")

        End Try
    End Sub

    Private Sub butPrint_Click(sender As Object, e As EventArgs) Handles butPrint.Click
        Dim xlApp As Excel.Application
        Dim xlWB As Excel.Workbook
        Dim xlCLwb As Excel.Workbook
        Dim xlCLwbName As String
        Dim xlWS As Excel.Worksheet
        Dim xlWBName As String
        Dim misValue As Object = System.Reflection.Missing.Value

        Try
            xlApp = New Excel.Application
            'Check if Excel is installed
            If xlApp Is Nothing Then
                MessageBox.Show("Excel Is Not properly installed!!")
                Exit Sub
            End If

            'Can use Try to catch an exception in opening a file for checking file existance. this way, not only
            'file exist is check, but also the user rights to folders is identified.
            'Only problem is if exception happens the code got to repeat in the Catch section of Try.

            xlCLwb = xlApp.Workbooks.Add(misValue)
            xlCLwbName = xlCLwb.Name
            xlApp.Visible = True

            For Each rW As DataGridViewRow In DataGridView1.SelectedRows
                'MessageBox.Show(rW.Cells(0).Value.ToString)
                xlWBName = rW.Cells(1).Value.ToString
                If Not (File.Exists(pmPath & chklstPath & xlWBName & ".xlsx")) Then
                    MsgBox(xlWBName & " Do Not have a checklist. Please work With Me To generate 1 If equipment requires PM.")

                Else
                    xlWB = xlApp.Workbooks.Open(pmPath & chklstPath & xlWBName & ".xlsx")
                    xlWS = xlWB.Sheets(xlWBName)
                    xlWS.Cells(1, 8) = rW.Cells(1).Value.ToString 'Eqp_Type
                    If xlWBName Like "*-DAILY" Then
                        xlWS.Cells(2, 8) = "DAILY"
                    ElseIf xlWBName Like "*-WEEKLY" Then
                        xlWS.Cells(2, 8) = "WEEKLY"
                    ElseIf xlWBName Like "*-MONTHLY" Then
                        xlWS.Cells(2, 8) = "MONTHLY"
                    ElseIf xlWBName Like "*-MONTHLY*" Then
                        xlWS.Cells(2, 8) = "MONTHLY"
                    ElseIf xlWBName Like "*-QUARTERLY" Then
                        xlWS.Cells(2, 8) = "QUARTERLY"
                    ElseIf xlWBName Like "*-HALFYEAR" Then
                        xlWS.Cells(2, 8) = "HALF-YEARLY"
                    ElseIf xlWBName Like "*-ANNUALLY" Then
                        xlWS.Cells(2, 8) = "ANNUALLY"
                    ElseIf xlWBName Like "*-CONDITIONAL" Then
                        xlWS.Cells(2, 8) = "CONDITIONAL"
                    ElseIf xlWBName Like "*-EVERYJOB" Then
                        xlWS.Cells(2, 8) = "EVERYJOB"
                    Else
                        xlWS.Cells(2, 8) = rW.Cells(6).Value.ToString 'Eqp_Freq
                    End If
                    xlWS.Cells(3, 8) = rW.Cells(5).Value.ToString 'Eqp_Line
                    If xlWBName = "TORQUEMASTER-DAILY" Or xlWBName = "RUSKA-7615" Then
                        xlWS.Cells(1, 22) = rW.Cells(0).Value.ToString 'Eqp_ID
                        xlWS.Cells(2, 22) = rW.Cells(4).Value.ToString 'Eqp_Mod
                        xlWS.Cells(3, 22) = rW.Cells(2).Value.ToString 'Eqp_SN
                    ElseIf rW.cells(6).Value.ToString = "DAILY"
                        xlWS.Cells(1, 22) = rW.Cells(0).Value.ToString 'Eqp_ID
                        xlWS.Cells(2, 22) = rW.Cells(4).Value.ToString 'Eqp_Mod
                        xlWS.Cells(3, 22) = rW.Cells(2).Value.ToString 'Eqp_SN
                    Else
                        xlWS.Cells(1, 20) = rW.Cells(0).Value.ToString 'Eqp_ID
                            xlWS.Cells(2, 20) = rW.Cells(4).Value.ToString 'Eqp_Mod
                            xlWS.Cells(3, 20) = rW.Cells(2).Value.ToString 'Eqp_SN
                        End If
                        xlWS.Copy(After:=xlCLwb.Sheets(1))
                        xlWB.Close(False)
                    End If
            Next
            'xlCLwb.SaveAs(clPath) 'can't save file and in any case no need to
            'xlApp.Quit() 'keep checklist open for printing
            releaseObject(xlWS)
            releaseObject(xlWB)
            releaseObject(xlCLwb)
            releaseObject(xlApp)
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at printing checklist." & Environment.NewLine &
                   "Please capture screen shot And notify Aik Koon. Thanks.")
            'xlApp.Quit()
            releaseObject(xlWS)
            releaseObject(xlWB)
            releaseObject(xlCLwb)
            releaseObject(xlApp)
        End Try
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        'sub to release Excel COM objects and clean Excel interop objects after opening the checklist for updates and masterlist
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Function GetUserName() As String
        If TypeOf My.User.CurrentPrincipal Is
          Security.Principal.WindowsPrincipal Then
            ' The application is using Windows authentication.
            ' The name format is DOMAIN\USERNAME.
            Dim parts() As String = Split(My.User.Name, "\")
            Dim username As String = parts(1)
            Return username
        Else
            ' The application is using custom authentication.
            Return My.User.Name
        End If
    End Function



    Private Sub tab1_Init(ByVal cmbCtl As String, popCmb As ComboBox)
        'Sub to initiate tabindex 1 form load to populate combo boxes
        Try
            Dim ds_Line As DataSet
            Dim ds_Type As DataSet
            Dim chkPath As Boolean 'TRUE=online, FALSE=offline
            Dim dTbl As String

            '************************************************************************************************************************************************************************************
            '*****************For access 'accdb' file access*************************************************************************************************************************************
            '*dbPath = "C:\Users\hb90747\Desktop\MDB\"
            '*dbName = "PM-Masterlist.accdb"
            '*
            '*SQLquery = "SELECT DISTINCT Eqp_Line from GageMasterEntry WHERE Eqp_Line IS NOT NULL AND (Eqp_Status<>'B-DOWN' OR Eqp_Status<>'REMOVE')"
            '*
            '*Call SQL_Query(SQLquery, "cmbProd_Line", "GageMasterEntry", 0)
            '*myConnToAccess = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & dbName)
            '************************************************************************************************************************************************************************************

            'myConnToAccess = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=P:\SING2 MANUFACTURING ENGINEERING\PM\PM-Masterlist.xlsm;Extended Properties=""Excel 12.0;HDR=YES""")
            'myConnToAccess = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=C:\Users\hb90747\Desktop\PM\PM-Masterlist.xlsm;Extended Properties=""Excel 12.0;HDR=YES""")

            If cmbCtl = "Eqp_Line" Then
                SQLquery = "SELECT DISTINCT " & cmbCtl & " from [GageMasterEntry] WHERE Eqp_Line IS NOT NULL"
                dTbl = "GageMasterEntry"
            ElseIf cmbCtl = "Eqp_Type" Then
                SQLquery = "SELECT DISTINCT " & cmbCtl & " FROM [GageMasterEntry]"
                dTbl = "GageMasterEntry"
            ElseIf cmbCtl = "Stat_Type" Then
                SQLquery = "SELECT * FROM [Stat_Type]"
                dTbl = "Stat_Type"
            ElseIf cmbCtl = "Eqp_Line1" Then
                SQLquery = "SELECT DISTINCT Eqp_Line FROM [Master_Lookup]"
                dTbl = "Master_Lookup"
                cmbCtl = "Eqp_Line"
            End If

            myConnToAccess.Open()
            ds_Line = New DataSet
            tables = ds_Line.Tables
            da = New OleDbDataAdapter(SQLquery, myConnToAccess)
            da.Fill(ds_Line, dTbl)
            'Dim view1 As New DataView(tables(0))
            With popCmb
                '.Items.Insert(0, String.Empty) 'Insert empty row to 1st line
                .DataSource = ds_Line.Tables(dTbl)
                .DisplayMember = cmbCtl
                .ValueMember = cmbCtl
                .SelectedIndex = 0
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
            myConnToAccess.Close()
        Catch ex As IO.FileNotFoundException
            MessageBox.Show("You may be offline. Please check connection and try again.")
            myConnToAccess.Close()
        Catch ex As Exception
            If lblNetStatus.Text = "Offline" Then
                MsgBox("You're not connected to network.")
            End If
            MsgBox(ex.Message & Environment.NewLine & "Error at tab1_init." & Environment.NewLine & " Please contact Aik Koon If error repeats after several tries.")
            myConnToAccess.Close()
        End Try
    End Sub

    Private Sub cmbNewEqpType_TextChanged(sender As Object, e As EventArgs) Handles cmbNewEqpType.TextChanged
        If cmbNewEqpType.SelectedIndex < 0 Then 'Prevent this event from firing if the changes are selected from dropdown
            'Restrict to 25 char to prevent excel spreadsheet name exceeed allowable char.
            Dim str As String = cmbNewEqpType.Text
            Dim sLen As Integer = str.Length
            If sLen > 25 Then
                MsgBox("Exceeded 25 characters! Please shorten the descriiption!", MsgBoxStyle.Exclamation)
                cmbNewEqpType.Text = Mid(str, 1, 25)
                Exit Sub
            End If

            str = str.ToUpper
            cmbNewEqpType.Text = str
            cmbNewEqpType.SelectionStart = str.Length


        End If
    End Sub

    Private Sub cmbNewEqpmLine_TextChanged(sender As Object, e As EventArgs) Handles cmbNewEqpLine.TextChanged
        If cmbNewEqpLine.SelectedIndex < 0 Then 'Prevent this event from firing if the changes are selected from dropdown
            Dim str As String = Me.cmbNewEqpLine.Text
            str = str.ToUpper
            cmbNewEqpLine.Text = str
            cmbNewEqpLine.SelectionStart = str.Length
        End If
    End Sub
    Private Sub tab0_cmbSelec(ByVal cmbCtl As String, popCmb As ComboBox, criteriaField As String, reqFieldData As String)
        'Sub to narrow datasource of either cmbProd_Line or cmbEqp_Name based on selection from other combobox
        Try
            Dim ds_Line As DataSet
            SQLquery = "SELECT DISTINCT " & cmbCtl & " from [GageMasterEntry] WHERE Eqp_Line IS NOT NULL AND " & criteriaField & "='" & reqFieldData & "'"
            myConnToAccess.Open()
            ds_Line = New DataSet
            tables = ds_Line.Tables
            da = New OleDbDataAdapter(SQLquery, myConnToAccess)
            da.Fill(ds_Line, "GageMasterEntry")
            'Dim view1 As New DataView(tables(0))
            With popCmb
                '.Items.Insert(0, String.Empty) 'Insert empty row to 1st line
                .DataSource = ds_Line.Tables("GageMasterEntry")
                .DisplayMember = cmbCtl
                .ValueMember = cmbCtl
                .SelectedIndex = 0
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
            myConnToAccess.Close()
        Catch ex As IO.FileNotFoundException
            MessageBox.Show("You may be offline. Please check connection and try again.")
            myConnToAccess.Close()
        Catch ex As Exception
            If lblNetStatus.Text = "Offline" Then
                MsgBox("You're not connected to network.")
            End If
            MsgBox(ex.Message & Environment.NewLine & "Error at tab0_cmbSel." & Environment.NewLine &
                   " Please contact Aik Koon If error repeats after several tries.")
            myConnToAccess.Close()
        End Try
    End Sub

    Private Sub txtnewmfgsn_TextChanged(sender As Object, e As EventArgs) Handles txtNewMfgSn.TextChanged
        Dim str As String = txtNewMfgSn.Text
        str = str.ToUpper
        txtNewMfgSn.Text = str
        txtNewMfgSn.SelectionStart = str.Length
    End Sub

    Private Sub txtnewmod_TextChanged(sender As Object, e As EventArgs) Handles txtNewMod.TextChanged
        Dim str As String = txtNewMod.Text
        str = str.ToUpper
        txtNewMod.Text = str
        txtNewMod.SelectionStart = str.Length
    End Sub

    Private Sub txtNewPrefix_TextChanged(sender As Object, e As EventArgs) Handles txtNewPrefix.TextChanged
        Dim str As String = txtNewPrefix.Text
        str = str.ToUpper
        txtNewPrefix.Text = str
        txtNewPrefix.SelectionStart = str.Length
    End Sub

    Private Sub txtNewMfg_TextChanged(sender As Object, e As EventArgs) Handles txtNewMfg.TextChanged
        Dim str As String = txtNewMfg.Text
        str = str.ToUpper
        txtNewMfg.Text = str
        txtNewMfg.SelectionStart = str.Length
    End Sub

    Private Sub butClrNew_Click(sender As Object, e As EventArgs) Handles butClrNew.Click
        With Me
            .cmbNewEqpLine.SelectedIndex = -1
            .cmbNewEqpType.SelectedIndex = -1
            .txtNewMfgSn.Text = ""
            .txtNewMod.Text = ""
            .txtNewPrefix.Text = ""
            .chkNewCal.Checked = False
            .chkNewPM.Checked = False
            .cmbNewFreq1.SelectedIndex = -1
            .cmbNewFreq2.SelectedIndex = -1
            .cmbNewFreq3.SelectedIndex = -1
            .cmbNewFreq4.SelectedIndex = -1
            .cmbNewResp1.SelectedIndex = -1
            .cmbNewResp2.SelectedIndex = -1
            .cmbNewResp3.SelectedIndex = -1
            .cmbNewResp4.SelectedIndex = -1
            .DataGridView2.DataSource = Nothing
        End With

    End Sub

    Private Sub butSubNew_Click(sender As Object, e As EventArgs) Handles butSubNew.Click
        'Submit new equipment info ito the system to generate ID
        Try
            Dim sQl As String
            Dim ds_Update As DataSet
            Dim iLastID As Object
            'Dim idValue As Object
            Dim lID As String
            Dim Msht As String
            Dim newEqpLine As String
            Dim newEqpType As String
            Dim Use_By As String

            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'new eqp type verification
            If cmbNewEqpType.SelectedIndex = -1 And cmbNewEqpType.Text = "" Then
                cmbNewEqpType.BackColor = Color.Red
                MsgBox("Equipment type information required!", MsgBoxStyle.Exclamation)
                cmbNewEqpType.BackColor = Color.White
                Exit Sub
            ElseIf cmbNewEqpType.SelectedIndex = -1 Then
                'Need to trim off leading / trailing spaces to prevent user entry error
                newEqpType = Trim(cmbNewEqpType.Text).ToUpper 'combobox.text is used if the index selected is not correct due to manual entry
            Else
                newEqpType = Trim(cmbNewEqpType.SelectedValue.ToString).ToUpper
            End If

            'new production line verification and prefix check for QC use only
            If cmbNewEqpLine.SelectedIndex = -1 And cmbNewEqpLine.Text = "" Then
                cmbNewEqpLine.BackColor = Color.Red
                MsgBox("Production line information required!", MsgBoxStyle.Exclamation)
                cmbNewEqpLine.BackColor = Color.White
                Exit Sub
            ElseIf cmbNewEqpLine.SelectedIndex = -1 Then
                newEqpLine = Trim(cmbNewEqpLine.Text)
            ElseIf cmbNewEqpLine.SelectedIndex > -1 Then
                newEqpLine = Trim(cmbNewEqpLine.SelectedValue.ToString).ToUpper
                If txtNewPrefix.Text <> "" And Mid(newEqpLine, 1, 2) <> "QA" Then
                    MsgBox("Prefix Entry is for QA used only!", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If
            End If

            If txtNewMfgSn.Text = "" Then
                txtNewMfgSn.BackColor = Color.Red
                MsgBox("Please enter the serial number or NA if no serial number.")
                txtNewMfgSn.BackColor = Color.White
                Exit Sub
            End If
            Dim newMfgSn As String = Trim(txtNewMfgSn.Text)
            Dim newMod As String = Trim(txtNewMod.Text)

            'check if sn is an existing one.
            sQl = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Eqp_Line FROM [GageMasterEntry] WHERE Eqp_SN LIKE '%" & newMfgSn &
            "' OR [GageMasterEntry].Eqp_SN LIKE '" & newMfgSn & "%'"

            Dim ds_SN = New DataSet
            tables = ds_SN.Tables
            da = New OleDbDataAdapter(sQl, myConnToAccess)
            da.Fill(ds_SN, "GageMasterEntry")
            If ds_SN.Tables(0).Rows.Count > 0 Then
                DataGridView2.DataSource = ds_SN
                DataGridView2.DataMember = "GageMasterEntry"
                myConnToAccess.Close()
                If MsgBox("Similar Serial Number already exist. Please see table below. Do you want to continue with new equipment updates?", MsgBoxStyle.YesNo) = vbNo Then
                    Exit Sub
                End If
            End If

            'Check entries
            If txtNewMod.Text = "" Then
                txtNewMod.BackColor = Color.Red
                MsgBox("Please enter the model info.")
                txtNewMod.BackColor = Color.White
                Exit Sub
            End If
            Dim newMfg As String = Trim(txtNewMfg.Text)

            'Ensure that a prefix is added for below 3 line setup
            If txtNewPrefix.Text = "" And (newEqpLine = "QA/CMM-LAB" Or newEqpLine = "QA/COLLAR-AREA") Then 'These 2 QC lab needs prefix
                txtNewPrefix.BackColor = Color.Red
                MsgBox("Please enter a prefix For New equipment setup used In " & newEqpLine & "!", MsgBoxStyle.Exclamation)
                txtNewPrefix.BackColor = Color.Transparent
                Exit Sub
            ElseIf txtNewPrefix.Text <> "" And newEqpLine <> "QA/CMM-LAB" And newEqpLine <> "QA/COLLAR-AREA" And newEqpLine <> "QA/MECH-LAB" Then
                If MsgBox("Prefix is used for CMM-LAB, COLLAR-AREA and MECH-LAB in QAQC." & Environment.NewLine &
                           "Prefix info will be ignored. Proceed with new equipment entry?", MsgBoxStyle.YesNo) = vbYes Then
                    txtNewPrefix.Text = ""
                Else
                    Exit Sub
                End If
            End If
            If chkNewCal.Checked = False And chkNewPM.Checked = False Then
                chkNewCal.BackColor = Color.Red
                chkNewPM.BackColor = Color.Red
                MsgBox("Please indicate If New equipment will need PM/CAL?", MsgBoxStyle.Exclamation)
                chkNewCal.BackColor = Color.Transparent
                chkNewPM.BackColor = Color.Transparent
                Exit Sub
            End If
            'Need to add checks on cal to ensure that no freq and resp is indicated. Below looping method looks more pro but may take longer time 
            'cos as the combobox count in the form increased, the time to loop through each check takes longer.
            'If chkNewPM.Checked = False Then
            ' Dim cbs = Controls.OfType(Of ComboBox)()
            ' For Each cb In cbs
            ' If (cb.Name Like "cmbNewFreq*" And cb.SelectedIndex > -1) Or (cb.Name Like "cmbNewResp*" And cb.SelectedIndex > -1) Then
            ' MsgBox("You have indicated PM Freq And Respsonibility For a Calibration only equipment. This info will be ignored.")
            ' End If
            ' cb.SelectedIndex = -1
            ' Next
            'End If
            'Use below instead
            If chkNewPM.Checked = False And chkNewCal.Checked = True And (cmbNewFreq1.SelectedIndex > -1 Or cmbNewFreq2.SelectedIndex > -1 Or cmbNewFreq3.SelectedIndex > -1 Or
                    cmbNewFreq4.SelectedIndex > -1 Or cmbNewResp1.SelectedIndex > -1 Or cmbNewResp2.SelectedIndex > -1 Or cmbNewResp3.SelectedIndex > -1 Or
                    cmbNewResp4.SelectedIndex > -1) Then
                MsgBox("You have indicated PM Freq And Respsonibility For a Calibration only equipment." & Environment.NewLine &
                           "The frequency And respsonibility info will be ignored.")
            End If

            If chkNewPM.Checked = True Then 'making sure pm freq is indicated
                If (cmbNewFreq1.SelectedIndex < 0 And cmbNewResp1.SelectedIndex < 0) Or
                    (cmbNewFreq1.SelectedIndex > -1 And cmbNewResp1.SelectedIndex < 0) Or
                    (cmbNewFreq1.SelectedIndex < 0 And cmbNewResp1.SelectedIndex > -1) Then
                    Label8.ForeColor = Color.Red
                    Label12.ForeColor = Color.Red
                    MsgBox("Please indicate a PM Freq And responsibility.")
                    Label8.ForeColor = Color.Black
                    Label12.ForeColor = Color.Black
                    Exit Sub
                End If
                If (cmbNewFreq2.SelectedIndex > -1 And cmbNewResp2.SelectedIndex < 0) Or
                    (cmbNewFreq2.SelectedIndex < 0 And cmbNewResp2.SelectedIndex > -1) Then
                    Label9.ForeColor = Color.Red
                    Label13.ForeColor = Color.Red
                    MsgBox("Please indicate a PM Freq And responsibility.")
                    Label9.ForeColor = Color.Black
                    Label13.ForeColor = Color.Black
                    Exit Sub
                End If
                If (cmbNewFreq3.SelectedIndex > -1 And cmbNewResp3.SelectedIndex < 0) Or
                    (cmbNewFreq3.SelectedIndex < 0 And cmbNewResp3.SelectedIndex > -1) Then
                    Label10.ForeColor = Color.Red
                    Label14.ForeColor = Color.Red
                    MsgBox("Please indicate a PM Freq And responsibility.")
                    Label10.ForeColor = Color.Black
                    Label14.ForeColor = Color.Black
                    Exit Sub
                End If
                If (cmbNewFreq4.SelectedIndex > -1 And cmbNewResp4.SelectedIndex < 0) Or
                    (cmbNewFreq4.SelectedIndex < 0 And cmbNewResp4.SelectedIndex > -1) Then
                    Label11.ForeColor = Color.Red
                    Label15.ForeColor = Color.Red
                    MsgBox("Please indicate a PM Freq And responsibility.")
                    Label11.ForeColor = Color.Black
                    Label15.ForeColor = Color.Black
                    Exit Sub
                End If
            End If

            'Get the last ID number for production line
            If txtNewPrefix.Text <> "" And (newEqpLine = "QA/CMM-LAB" Or newEqpLine = "QA/COLLAR-AREA" Or newEqpLine = "QA/MECH-LAB") Then
                Dim newPrefix As String = Trim(txtNewPrefix.Text)
                If newPrefix.Contains("-") Then newPrefix = newPrefix.Replace("-", "")
                sQl = "Select Eqp_ID FROM [GageMasterEntry] WHERE Eqp_ID Like 'HS2-" & newPrefix & "[0-9]%' " &
                    "ORDER BY Eqp_ID ASC"
            ElseIf txtNewPrefix.Text = "" And newEqpLine = "QA/MECH-LAB" Then
                Dim newPrefix As String = Trim(txtNewPrefix.Text)
                If newPrefix.Contains("-") Then newPrefix = newPrefix.Replace("-", "")
                sQl = "Select Eqp_ID FROM [GageMasterEntry] WHERE Eqp_ID Like 'HS2-[0-9]%' " &
                        "ORDER BY Eqp_ID ASC"
            Else
                'sQl = "SELECT [GageMasterEntry].Eqp_ID FROM [GageMasterEntry] INNER JOIN [Master_Lookup] ON [GageMasterEntry].Eqp_ID" &
                '" Like [Master_Lookup].ID_Format WHERE [Master_Lookup].Eqp_Line='" & newEqpLine & "' ORDER BY LENGTH([GageMasterEntry].Eqp_ID),[GageMasterEntry].Eqp_ID"
                sQl = "SELECT Eqp_ID FROM [GageMasterEntry] WHERE Eqp_Line='" & newEqpLine & "' ORDER BY LEN(Eqp_ID),Eqp_ID"
            End If
            'Dim myCommand As New OleDbCommand(sQl, myConnToAccess)
            'myCommand.Connection.Open()
            ds_Update = New DataSet
            tables = ds_Update.Tables
            da = New OleDbDataAdapter(sQl, myConnToAccess)
            da.Fill(ds_Update, "GageMasterEntry")
            If ds_Update.Tables(0).Rows.Count > 0 Then
                lID = ds_Update.Tables(0)(ds_Update.Tables(0).Rows.Count - 1)("Eqp_ID").ToString
                'idValue = myCommand.ExecuteScalar
            ElseIf ds_Update.Tables(0).Rows.Count = 0 And (txtNewPrefix.Text <> "" Or cmbUSEqp_Line.SelectedValue.ToString = "QA/MECH-LAB") Then
                lID = "HS2-" & Trim(txtNewPrefix.Text) & "00"
            End If
            'MsgBox(lID)
            myConnToAccess.Close()
            'Exit Sub

            'Get the Master Sheet info
            sQl = "SELECT PM_Sht,Used_By from [Master_Lookup] WHERE Eqp_Line='" & newEqpLine & "'"
            'Dim myCommand As New OleDbCommand(sQl, myConnToAccess)
            'myCommand.Connection.Open()
            ds_Update = New DataSet
            tables = ds_Update.Tables
            da = New OleDbDataAdapter(sQl, myConnToAccess)
            da.Fill(ds_Update, "Master_Lookup")
            If ds_Update.Tables(0).Rows.Count > 0 Then
                Msht = ds_Update.Tables(0)(ds_Update.Tables(0).Rows.Count - 1)("PM_Sht").ToString
                Use_By = ds_Update.Tables(0)(ds_Update.Tables(0).Rows.Count - 1)("Used_By").ToString
                'idValue = myCommand.ExecuteScalar
                'Msht = iLastID.ToString
                'MsgBox(Msht)
            End If
            myConnToAccess.Close()

            Dim nxID As Integer
            If txtNewPrefix.Text <> "" And (newEqpLine = "QA/CMM-LAB" Or newEqpLine = "QA/COLLAR-AREA" Or newEqpLine = "QA/MECH-LAB") Then
                nxID = Mid(lID, Len(Trim(txtNewPrefix.Text)) + 5, Len(lID) - Len(Trim(txtNewPrefix.Text)) + 3) + 1 'Lenght 4=HS2-%+1
            Else
                nxID = Mid(lID, (InStrRev(lID, "-") * 1) + 1) + 1
            End If

            Dim newEqp_ID As String
            If txtNewPrefix.Text = "" Then
                If nxID < 10 Then
                    newEqp_ID = Mid(lID, 1, (InStrRev(lID, "-") * 1) - 1) & "-0" & nxID 'Add 0 for <10 int
                Else
                    newEqp_ID = Mid(lID, 1, (InStrRev(lID, "-") * 1) - 1) & "-" & nxID
                End If
            Else
                If nxID < 10 Then
                    newEqp_ID = Mid(lID, 1, (InStrRev(lID, "-") * 1) - 1) & "-" & Trim(txtNewPrefix.Text) & "00" & nxID 'Add 0 for <10 int
                Else
                    newEqp_ID = Mid(lID, 1, (InStrRev(lID, "-") * 1) - 1) & "-" & Trim(txtNewPrefix.Text) & nxID
                End If
            End If
            Dim newEqpType1 As String
            Dim newEqpType2 As String
            Dim newEqpType3 As String
            Dim newEqpType4 As String
            Dim Type0 As Boolean = False 'CAL only equipment
            Dim Type1 As Boolean = False 'Multiple PM Frequency. This is True after checking if Freq2 is set.
            Dim Type11 As Boolean = False 'Equipment only has 1 PM freq
            Dim Type2 As Boolean = False
            Dim Type3 As Boolean = False
            Dim Type4 As Boolean = False

            If chkNewCal.Checked = True Then Type0 = True

            If chkNewPM.Checked = True And cmbNewFreq1.SelectedIndex > -1 And cmbNewResp1.SelectedIndex > -1 Then
                If cmbNewFreq2.SelectedIndex > -1 Then
                    Type1 = True
                ElseIf cmbNewFreq2.SelectedIndex = -1 Then
                    Type11 = True 'Equipment has multiple PM freq.
                End If
            End If

            If chkNewPM.Checked = True And cmbNewFreq2.SelectedIndex > -1 And cmbNewResp2.SelectedIndex > -1 Then
                If cmbNewFreq1.SelectedIndex = -1 And cmbNewResp1.SelectedIndex = -1 Then
                    Label8.ForeColor = Color.Red
                    Label2.ForeColor = Color.Red
                    MsgBox("Please indicate the frequency and responsibility in Freq1 and Resp1!", MsgBoxStyle.Exclamation)
                    Label8.ForeColor = Color.Black
                    Label2.ForeColor = Color.Black
                    Exit Sub
                End If
                Type2 = True
            End If

            If chkNewPM.Checked = True And cmbNewFreq3.SelectedIndex > -1 And cmbNewResp3.SelectedIndex > -1 Then
                If cmbNewFreq1.SelectedIndex = -1 And cmbNewResp1.SelectedIndex = -1 Then
                    Label8.ForeColor = Color.Red
                    Label12.ForeColor = Color.Red
                    MsgBox("Please indicate the frequency and responsibility in Freq1 and Resp1!", MsgBoxStyle.Exclamation)
                    Label8.ForeColor = Color.Black
                    Label12.ForeColor = Color.Black
                    Exit Sub
                ElseIf cmbNewFreq2.SelectedIndex = -1 And cmbNewResp2.SelectedIndex = -1 Then
                    Label9.ForeColor = Color.Red
                    Label13.ForeColor = Color.Red
                    MsgBox("Please indicate the frequency and responsibility in Freq1 and Resp1!", MsgBoxStyle.Exclamation)
                    Label9.ForeColor = Color.Black
                    Label13.ForeColor = Color.Black
                    Exit Sub
                End If
                Type3 = True
            End If

            If chkNewPM.Checked = True And cmbNewFreq4.SelectedIndex > -1 And cmbNewResp4.SelectedIndex > -1 Then
                If cmbNewFreq1.SelectedIndex = -1 And cmbNewResp1.SelectedIndex = -1 Then
                    Label8.ForeColor = Color.Red
                    Label12.ForeColor = Color.Red
                    MsgBox("Please indicate the frequency and responsibility in Freq1 and Resp1!", MsgBoxStyle.Exclamation)
                    Label8.ForeColor = Color.Black
                    Label12.ForeColor = Color.Black
                    Exit Sub
                ElseIf cmbNewFreq2.SelectedIndex = -1 And cmbNewResp2.SelectedIndex = -1 Then
                    Label9.ForeColor = Color.Red
                    Label13.ForeColor = Color.Red
                    MsgBox("Please indicate the frequency and responsibility in Freq1 and Resp1!", MsgBoxStyle.Exclamation)
                    Label9.ForeColor = Color.Black
                    Label13.ForeColor = Color.Black
                    Exit Sub
                ElseIf cmbNewFreq3.SelectedIndex = -1 And cmbNewResp3.SelectedIndex = -1 Then
                    Label10.ForeColor = Color.Red
                    Label14.ForeColor = Color.Red
                    MsgBox("Please indicate the frequency and responsibility in Freq1 and Resp1!", MsgBoxStyle.Exclamation)
                    Label10.ForeColor = Color.Black
                    Label14.ForeColor = Color.Black
                    Exit Sub
                End If
                Type4 = True
            End If

            If Type0 = True And Type11 = False Then
                Call insNewEqp("", "", newEqpType, newEqp_ID, newMfgSn, newMfg, newMod, newEqpLine, "", "CAL-ONLY", Use_By, "NO")
            End If

            If Type1 = True Then
                newEqpType1 = newEqpType & "-" & cmbNewFreq1.SelectedItem.ToString 'Multiple PM freq equipment. Need to add freq suffix into the equipment description
                Call insNewEqp(cmbNewFreq1.SelectedItem.ToString, cmbNewResp1.SelectedItem.ToString, newEqpType1, newEqp_ID, newMfgSn, newMfg, newMod, newEqpLine, "ACTIVE", Msht, Use_By, "YES")
                If (File.Exists(pmPath & chklstPath & newEqpType1 & ".xlsx")) = False And cmbNewResp1.SelectedItem.ToString = "INTERNAL" Then
                    MsgBox("PM Cheklist for " & newEqpType1 & " is not generated. Please liaise with Mfg Eng to generate the checklist.")
                End If
            ElseIf Type11 = True Then
                Call insNewEqp(cmbNewFreq1.SelectedItem.ToString, cmbNewResp1.SelectedItem.ToString, newEqpType, newEqp_ID, newMfgSn, newMfg, newMod, newEqpLine, "ACTIVE", Msht, Use_By, "YES")
                If (File.Exists(pmPath & chklstPath & newEqpType & ".xlsx")) = False And cmbNewResp1.SelectedItem.ToString = "INTERNAL" Then
                    MsgBox("PM Checklist for " & newEqpType & " is not generated. Please liaise with Mfg Eng to generate the checklist.")
                End If
            End If

            If Type2 = True Then
                newEqpType2 = newEqpType & "-" & cmbNewFreq2.SelectedItem.ToString
                Call insNewEqp(cmbNewFreq2.SelectedItem.ToString, cmbNewResp2.SelectedItem.ToString, newEqpType2, newEqp_ID, newMfgSn, newMfg, newMod, newEqpLine, "ACTIVE", Msht, Use_By, "YES")
                If (File.Exists(pmPath & chklstPath & newEqpType2 & ".xlsx")) = False And cmbNewResp2.SelectedItem.ToString = "INTERNAL" Then
                    MsgBox("PM Checklist for " & newEqpType2 & " is not generated. Please liaise with Mfg Eng to generate the checklist.")
                End If
            End If

            If Type3 = True Then
                newEqpType3 = newEqpType & "-" & cmbNewFreq3.SelectedItem.ToString
                Call insNewEqp(cmbNewFreq3.SelectedItem.ToString, cmbNewResp3.SelectedItem.ToString, newEqpType3, newEqp_ID, newMfgSn, newMfg, newMod, newEqpLine, "ACTIVE", Msht, Use_By, "YES")
                If (File.Exists(pmPath & chklstPath & newEqpType3 & ".xlsx")) = False And cmbNewResp3.SelectedItem.ToString = "INTERNAL" Then
                    MsgBox("PM Checklist for " & newEqpType3 & " is not generated. Please liaise with Mfg Eng to generate the checklist.")
                End If
            End If

            If Type4 = True Then
                newEqpType4 = newEqpType & "-" & cmbNewFreq4.SelectedItem.ToString
                Call insNewEqp(cmbNewFreq4.SelectedItem.ToString, cmbNewResp4.SelectedItem.ToString, newEqpType4, newEqp_ID, newMfgSn, newMfg, newMod, newEqpLine, "ACTIVE", Msht, Use_By, "YES")
                If (File.Exists(pmPath & chklstPath & newEqpType4 & ".xlsx")) = False And cmbNewResp4.SelectedItem.ToString = "INTERNAL" Then
                    MsgBox("PM Checklist for " & newEqpType4 & " is not generated. Please liaise with Mfg Eng to generate the checklist.")
                End If
            End If

            MsgBox("New Eqp ID: " & newEqp_ID & vbCrLf & "Eqp_Type: " & newEqpType & vbCrLf & "Prod_Line: " & newEqpLine)
            Call clrTab1()
            System.Windows.Forms.Cursor.Current = Cursors.Default

        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at butsub_New." & Environment.NewLine &
                   " Please capture screenshot and email Aik Koon. Thank you for assisting.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default

        End Try
    End Sub

    Private Sub insNewEqp(cFreq As String, cResp As String, EqpType As String, Eqp_ID As String, MfgSn As String,
                          nMfg As String, nMod As String, EqLine As String, nStat As String, nMsht As String, uSer As String, pM As String)
        Try
            Dim insSQL As String
            myConnToAccess.Open()
            'myCommand.Connection = myConnToAccess
            'insert new equipment into GageMasterEntry.

            insSQL = "INSERT INTO [GageMasterEntry] (Used_By,Eqp_Type,Eqp_ID,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,PM_Sht,PM_Need)" &
                    " VALUES('" & uSer & "','" & EqpType & "','" & Eqp_ID & "','" & MfgSn & "','" & nMfg & "','" & nMod & "','" & EqLine &
                    "','" & cFreq & "','" & cResp & "','ACTIVE','" & nMsht & "','" & pM & "')"
            Dim myCommand As New OleDbCommand(insSQL, myConnToAccess)
            'myCommand.Connection.Open()
            'myCommand.CommandText = sql
            myCommand.ExecuteNonQuery()
            myConnToAccess.Close()

        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at insNewEqp." & Environment.NewLine &
                   " Please capture screenshot and email Aik Koon. Thank you for assisting.")
        End Try
    End Sub


    Private Sub clrTab1()
        If cmbNewEqpLine.SelectedIndex = -1 Then
            cmbNewEqpLine.Text = ""
        Else
            cmbNewEqpLine.SelectedIndex = -1
        End If
        If cmbNewEqpType.SelectedIndex = -1 Then
            cmbNewEqpType.Text = ""
        Else
            cmbNewEqpType.SelectedIndex = -1
        End If
        txtNewMfg.Text = ""
        txtNewMfgSn.Text = ""
        txtNewMod.Text = ""
        txtNewPrefix.Text = ""
        chkNewCal.Checked = False
        chkNewPM.Checked = False
        cmbNewFreq1.SelectedItem = Nothing
        cmbNewFreq2.SelectedItem = Nothing
        cmbNewFreq3.SelectedItem = Nothing
        cmbNewFreq4.SelectedItem = Nothing
        cmbNewResp1.SelectedItem = Nothing
        cmbNewResp2.SelectedItem = Nothing
        cmbNewResp3.SelectedItem = Nothing
        cmbNewResp4.SelectedItem = Nothing
        DataGridView2.DataSource = Nothing
    End Sub

    Private Sub butUpdateEqp_Click(sender As Object, e As EventArgs) Handles butUpdateEqp.Click
        'Update button from tabpage0
        Dim ds_Update As DataSet
        Dim selID As String
        Dim dgID As String

        Try
            If DataGridView1.SelectedRows.Count = 0 Then
                MsgBox("Please select 1 equipment for updating!")
                Exit Sub
            End If
            If DataGridView1.SelectedRows.Count > 1 Then
                MsgBox("Select only 1 equipment to update.")
                DataGridView1.ClearSelection()
                Exit Sub
            End If
            'MsgBox(DataGridView1.SelectedRows(0).Cells(0).Value.ToString)
            selID = DataGridView1.SelectedRows(0).Cells(0).Value.ToString

            'Call tab3_Init()
            'SELECT Eqp_ID,Eqp_Type,Eqp_SN,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Status,Old_ID,Eqp_Resp from [GageMasterEntry]
            Me.TabControl1.SelectedTab = TabPage3

            txtUSEqpID.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString
            txtUSEqpID.ReadOnly = True
            If cmbUSEqpStat.Items.Count = 0 Then
                cmbUSEqpStat.Text = DataGridView1.SelectedRows(0).Cells(7).Value.ToString
            Else
                cmbUSEqpStat.SelectedValue = DataGridView1.SelectedRows(0).Cells(7).Value.ToString
            End If
            cmbUSEqp_Line.SelectedValue = DataGridView1.SelectedRows(0).Cells(5).Value.ToString
            cmbUSEqpDesc.SelectedValue = DataGridView1.SelectedRows(0).Cells(1).Value.ToString
            txtUSMfgSN.Text = DataGridView1.SelectedRows(0).Cells(2).Value.ToString
            txtUSModel.Text = DataGridView1.SelectedRows(0).Cells(4).Value.ToString
            txtUSmfg.Text = DataGridView1.SelectedRows(0).Cells(3).Value.ToString
            txtUSOldID.Text = DataGridView1.SelectedRows(0).Cells(8).Value.ToString
            txtEqpOwner.Text = DataGridView1.SelectedRows(0).Cells(10).Value.ToString
            cmbUSFreq1.SelectedItem = Trim(DataGridView1.SelectedRows(0).Cells(6).Value.ToString)
            cmbUSResp1.SelectedItem = DataGridView1.SelectedRows(0).Cells(9).Value.ToString
            cmbEqp_Name.SelectedIndex = -1
            cmbProd_Line.SelectedIndex = -1
            DataGridView1.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at butUpdateEqp." & Environment.NewLine &
                    "Please capture screenshot and email Aik Koon. Thank You for assisting.")
        End Try
    End Sub

    Private Sub tab3_Init()
        Dim SQLupdate As String
        Dim ds_updateEqpLine As DataSet
        Dim ds_updateEqpType As DataSet

        Try
            SQLupdate = "SELECT DISTINCT Eqp_Line from [GageMasterEntry] WHERE Eqp_Line IS NOT NULL"
            myConnToAccess.Open()
            ds_updateEqpLine = New DataSet
            tables = ds_updateEqpLine.Tables
            da = New OleDbDataAdapter(SQLupdate, myConnToAccess)
            da.Fill(ds_updateEqpLine, "GageMasterEntry")
            'Dim view1 As New DataView(tables(0))
            With cmbUSEqp_Line
                '.Items.Insert(0, String.Empty) 'Insert empty row to 1st line
                .DataSource = ds_updateEqpLine.Tables("GageMasterEntry")
                .DisplayMember = "Eqp_Line"
                .ValueMember = "Eqp_Line"
                .SelectedIndex = 0
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With

            SQLupdate = "SELECT DISTINCT Eqp_Type from [GageMasterEntry]"
            ds_updateEqpType = New DataSet
            dT = New OleDbDataAdapter(SQLupdate, myConnToAccess)
            dT.Fill(ds_updateEqpType, "GageMasterEntry")
            With cmbUSEqpDesc
                .DataSource = ds_updateEqpType.Tables("GageMasterEntry")
                .DisplayMember = "Eqp_Type"
                .ValueMember = "Eqp_Type"
                .SelectedIndex = 0
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
            myConnToAccess.Close()

            Call addFreq(cmbUSFreq1)
            'Call addFreq(cmbUSFreq2)
            'Call addFreq(cmbUSFreq3)
            'Call addFreq(cmbUSFreq4)
            Call addResp(cmbUSResp1)
            'Call addResp(cmbUSResp2)
            'Call addResp(cmbUSResp3)
            'Call addResp(cmbUSResp4)

            cmbUSEqpStat.Items.Add("ACTIVE")
            cmbUSEqpStat.Items.Add("B-DOWN")
            cmbUSEqpStat.Items.Add("DECOM")
            cmbUSEqpStat.Items.Add("REMOVE")

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Error at tab3_Init. Please capture screen shot and notify Aik Koon. Thank You.")
        End Try
    End Sub

    Private Sub butclrUSentries_Click(sender As Object, e As EventArgs) Handles butclrUSentries.Click
        Call clr_Tab2()
    End Sub

    Private Sub butUpdate_Click(sender As Object, e As EventArgs) Handles butUpdate.Click
        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim updateSQL As String
            If cmbUSEqpStat.SelectedIndex = -1 Then
                Label18.ForeColor = Color.Red
                MsgBox("Please indicate the equipment status.")
                Label18.ForeColor = Color.Black
                Exit Sub
            End If
            Dim updateEqpStat As String = Trim(cmbUSEqpStat.SelectedValue.ToString)

            If cmbUSEqp_Line.SelectedIndex = -1 Then
                Label20.ForeColor = Color.Red
                MsgBox("Please select a productline that the equipment is used " & Environment.NewLine & "NA if not applicable.")
                Label20.ForeColor = Color.Red
                Exit Sub
            End If
            Dim updateEqp_Line As String = Trim(cmbUSEqp_Line.SelectedValue.ToString)
            If cmbUSEqpStat.SelectedValue.ToString <> "DELETE" Then
                'myConnToAccess.Open()
                If cmbUSEqpDesc.SelectedIndex = -1 Then
                    Label19.ForeColor = Color.Red
                    MsgBox("Please select an equipment description")
                    Label19.ForeColor = Color.Black
                    Exit Sub
                End If
                Dim updateEqpDesc As String = Trim(cmbUSEqpDesc.SelectedValue.ToString)
                If updateEqpDesc Like "TORQUEMASTER-*" Then
                    MsgBox("TorqueMater has multiple PM frequency." & vbCr & "Please contact Aik Koon to update PM changes.", MsgBoxStyle.Information)
                    Exit Sub
                End If

                If txtUSMfgSN.Text = "" Then
                    Label21.ForeColor = Color.Red
                    MsgBox("Please enter manufacturer serial number or NA if no serial number.")
                    Label21.ForeColor = Color.Black
                    Exit Sub
                End If
                Dim updateMfgSN As String = Trim(txtUSMfgSN.Text)


                If txtUSModel.Text = "" Then
                    Label22.ForeColor = Color.Red
                    MsgBox("Please enter the model ")
                    Label22.ForeColor = Color.Black
                    Exit Sub
                End If
                Dim updateModel As String = Trim(txtUSModel.Text)

                Dim updateMfg As String = Trim(txtUSmfg.Text)

                If cmbUSFreq1.SelectedIndex = -1 Then
                    Label31.ForeColor = Color.Red
                    MsgBox("Please indicate the PM frequency or select 'DELETE' if this equipment PM is no longer needed")
                    Label31.ForeColor = Color.Black
                    Exit Sub
                ElseIf cmbUSFreq1.SelectedIndex > -1 Then
                    Dim updateFreq As String = Trim(cmbUSFreq1.SelectedItem.ToString)
                    If cmbUSResp1.SelectedIndex = -1 Then
                        Label27.ForeColor = Color.Red
                        MsgBox("Please indicate PM responsibility")
                        Label27.ForeColor = Color.Black
                        Exit Sub
                    Else
                        Dim updateResp As String = Trim(cmbUSResp1.SelectedItem.ToString)
                        'SELECT Eqp_ID,Eqp_Type,Eqp_SN,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Status,Old_ID,Eqp_Resp from [GageMasterEntry]
                        updateSQL = "UPDATE [GageMasterEntry] SET Eqp_Type='" & updateEqpDesc & "',Eqp_SN='" & updateMfgSN & "',Manufacturer='" & txtUSmfg.Text & "',Eqp_Mod='" & updateModel & "',Eqp_Line='" &
                            updateEqp_Line & "',Eqp_Freq='" & updateFreq & "',Eqp_Status='" & updateEqpStat & "',Old_ID='" & txtUSOldID.Text & "',Eqp_Resp='" & updateResp &
                            "',Eqp_Owner='" & txtEqpOwner.Text & "' WHERE Eqp_ID='" & txtUSEqpID.Text & "'"
                    End If
                End If
            ElseIf cmbUSEqpStat.SelectedValue.ToString = "DELETE" Then
                Dim delPM As Boolean = True
                Dim confirmDel As Integer
                confirmDel = MsgBox("You have selected to delete PM. Proceed to delete?", MsgBoxStyle.YesNo)
                If confirmDel = 7 Then
                    Exit Sub
                ElseIf confirmDel = 6 Then
                    'Mark PM record as deleted. no other eqm info is necessary to be updated if delete
                    updateSQL = "UPDATE [GageMasterEntry] SET Eqp_Status='DELETE' WHERE Eqp_ID='" & txtUSEqpID.Text & "'"
                End If
            End If

            'Update into GageMasterEntry table
            myConnToAccess.Open()
            Dim myCommand As New OleDbCommand(updateSQL, myConnToAccess)
            myCommand.ExecuteNonQuery()
            myConnToAccess.Close()
            MsgBox(txtUSEqpID.Text & " equipment information is updated. Thank you!")
            'updateSQL = "INSERT INTO [GageMasterEntry] (Eqp_ID,Eqp_Type,Eqp_SN,Eqp_Type,Manufacturer)" &
            '" VALUES('" & Eqp_ID & "','" & EqpType & "','" & MfgSn & "','" & nMod & "','" & nMfg & "')"
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & " Please contact Aik Koon If error repeats after several tries.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default

        End Try
    End Sub

    Private Sub butClrT1_Click(sender As Object, e As EventArgs) Handles butClrT1.Click
        cmbNewFreq1.SelectedIndex = -1
        cmbNewResp1.SelectedIndex = -1
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles butClrT2.Click
        cmbNewFreq2.SelectedIndex = -1
        cmbNewResp2.SelectedIndex = -1
    End Sub

    Private Sub butClrT3_Click(sender As Object, e As EventArgs) Handles butClrT3.Click
        cmbNewFreq3.SelectedIndex = -1
        cmbNewResp3.SelectedIndex = -1
    End Sub

    Private Sub butClrT4_Click(sender As Object, e As EventArgs) Handles butClrT4.Click
        cmbNewResp4.SelectedIndex = -1
        cmbNewResp4.SelectedIndex = -1
    End Sub

    Private Sub butBreakDown_Click(sender As Object, e As EventArgs) Handles butBreakDown.Click
        Dim selRowCount As Integer = DataGridView1.SelectedRows.Count
        Dim ds As DataSet

        If selRowCount > 1 Then
            MsgBox("Please select only 1 equipment.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim bkDwnID As String = DataGridView1.SelectedRows(0).Cells(0).Value.ToString
        Dim bkDwnType As String = DataGridView1.SelectedRows(0).Cells(1).Value.ToString
        Dim prodLine As String = DataGridView1.SelectedRows(0).Cells(5).Value.ToString
        Form3.lblBrkDwnID.Text = bkDwnID
        Form3.lblBrkDwnType.Text = bkDwnType
        Form3.lblProdLine.Text = prodLine
        Dim getBDSQL = "SELECT ID,Fault,Start_Date,Eq_Status,Svc_Cost FROM [Fault] WHERE Eqp_ID='" & bkDwnID & "'"

        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            ds = New DataSet
            da = New OleDbDataAdapter(getBDSQL, myConnToAccess)
            da.Fill(ds, "Fault")
            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox("No Breakdown records found for " & bkDwnID & ".")
            End If
            myConnToAccess.Close()
            Form3.DataGridBrkDwn.DataSource = ds
            Form3.DataGridBrkDwn.DataMember = "Fault"
            Form3.txtFaultSymptom.Text = ""
            Form3.txtSoln.Text = ""

            getBDSQL = "SELECT [Master_Lookup].Notify_ID FROM [Master_Lookup] LEFT JOIN [GageMasterEntry] ON" &
                " [GageMasterEntry].Eqp_Line=[Master_Lookup].Eqp_Line WHERE [GageMasterEntry].Eqp_ID='" & bkDwnID & "'"
            myConnToAccess.Open()
            Dim dN As DataSet = New DataSet
            Dim dNote As OleDbDataAdapter = New OleDbDataAdapter(getBDSQL, myConnToAccess)
            dNote.Fill(dN, "Master_Lookup")
            myConnToAccess.Close()
            Form3.DataGridNotify.DataSource = dN
            Form3.DataGridNotify.DataMember = "Master_Lookup"
            System.Windows.Forms.Cursor.Current = Cursors.Default

            Form3.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at breakdown button click event." & Environment.NewLine &
                   "Please capture screen shot and notify Aik Koon. Thanks.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default

        End Try

    End Sub



    Private Sub butSearch_Click(sender As Object, e As EventArgs) Handles butSearch.Click
        Dim searchID As String = txtUSEqpID.Text
        If searchID.StartsWith("*") Or searchID.EndsWith("*") Then
            searchID = searchID.Replace("*", "%")
        End If
        Dim searchSN As String = txtUSMfgSN.Text
        If searchSN.StartsWith("*") Or searchSN.EndsWith("*") Then
            searchSN = searchSN.Replace("*", "%")
        End If
        Dim searchStat As String
        Dim searchOwner As String = txtEqpOwner.Text
        If cmbUSEqpStat.SelectedIndex = -1 Then
            searchStat = ""
        Else
            searchStat = cmbUSEqpStat.SelectedValue.ToString
        End If

        If searchOwner <> "" And (Len(searchOwner) <> 7 Or Not (searchOwner Like "H*")) And chkWithBkDwn.Checked = False Then
            Label29.ForeColor = Color.Red
            MsgBox("Please enter a valid User ID H??????")
            Label29.ForeColor = Color.Black
            Exit Sub
        End If



        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        If searchID = "" And searchSN = "" And searchStat = "" And searchOwner = "" And chkWithBkDwn.Checked = False Then
            MsgBox("Please enter search criteria for Eqp_ID or Mfg_SN or Eqp_Status or Eqpm_Ownerto search!")
            Exit Sub
        End If
        Dim searchSQL As String
        If chkWithBkDwn.Checked = False Then

            If searchID <> "" And searchSN = "" And searchStat = "" And searchOwner = "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_ID LIKE '" & searchID & "'"

            ElseIf searchID = "" And searchSN <> "" And searchStat = "" And searchOwner = "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_SN LIKE '" & searchSN & "'"

            ElseIf searchID = "" And searchSN = "" And searchStat <> "" And searchOwner = "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_Status='" & searchStat & "'"
            ElseIf searchID = "" And searchSN = "" And searchStat = "" And searchOwner <> "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_Owner='" & searchOwner & "'"

            ElseIf searchID <> "" And searchSN <> "" And searchStat = "" And searchOwner = "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_SN LIKE '" & searchSN &
            "' AND Eqp_ID LIKE '" & searchID & "'"

            ElseIf searchID = "" And searchSN <> "" And searchStat <> "" And searchOwner = "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_SN LIKE '" & searchSN &
            "'  AND Eqp_Status='" & searchStat & "'"

            ElseIf searchID <> "" And searchSN = "" And searchStat <> "" And searchOwner = "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_ID LIKE '" & searchID &
            "' AND Eqp_Status='" & searchStat & "'"

            ElseIf searchID <> "" And searchSN <> "" And searchStat <> "" And searchOwner = "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_SN LIKE '" & searchSN &
            "' AND Eqp_ID LIKE '" & searchID & "' AND Eqp_Status='" & searchStat & "'"

            ElseIf searchID <> "" And searchSN <> "" And searchStat = "" And searchOwner <> "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_SN LIKE '" & searchSN &
            "' AND Eqp_ID LIKE '" & searchID & "' AND Eqp_Owner='" & searchOwner & "'"

            ElseIf searchID = "" And searchSN <> "" And searchStat <> "" And searchOwner <> "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_SN LIKE '" & searchSN &
            "' AND Eqp_Status='" & searchStat & "' AND Eqp_Owner='" & searchOwner & "'"

            ElseIf searchID <> "" And searchSN = "" And searchStat <> "" And searchOwner <> "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_ID LIKE '" & searchID &
            "' AND Eqp_Status='" & searchStat & "' AND Eqp_Owner='" & searchOwner & "'"

            ElseIf searchID <> "" And searchSN <> "" And searchStat <> "" And searchOwner <> "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_SN LIKE '" & searchSN &
            "' AND Eqp_ID LIKE '" & searchID & "' AND Eqp_Status='" & searchStat & "' AND Eqp_Owner='" & searchOwner & "'"

            ElseIf searchID = "" And searchSN <> "" And searchStat = "" And searchOwner <> "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE Eqp_SN LIKE '" & searchSN &
            "' AND Eqp_Owner='" & searchOwner & "'"

            ElseIf searchID = "" And searchSN = "" And searchStat <> "" And searchOwner <> "" Then
                searchSQL = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Manufacturer,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE  Eqp_Status='" & searchStat &
                "' AND Eqp_Owner='" & searchOwner & "'"

            End If
        ElseIf chkWithBkDwn.Checked = True Then
            'If searchID <> "" And searchSN = "" And searchStat = "" And searchOwner = "" Then
            searchSQL = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,Manufacturer,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
                    "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Resp,[GageMasterEntry].Eqp_Status,[GageMasterEntry].Old_ID,[GageMasterEntry].Eqp_Owner " &
                    "FROM [Fault] LEFT JOIN [GageMasterEntry] on [GageMasterEntry].Eqp_ID=[Fault].Eqp_ID" 'ORDER BY COUNT([Fault].Eqp_ID)"
            'End If
        End If
        Dim ds As DataSet
        Try
            ds = New DataSet
            tables = ds.Tables
            da = New OleDbDataAdapter(searchSQL, myConnToAccess)
            da.Fill(ds, "GageMasterEntry")
            myConnToAccess.Close()
            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox("No equipment found from search criteria for:" & Environment.NewLine & "Eqp_ID: " & searchID & Environment.NewLine &
                       "Mfg_SN: " & searchSN & Environment.NewLine & "Eqp_Status: " & searchStat & Environment.NewLine &
                       "Eqp_Owner: " & txtEqpOwner.Text & Environment.NewLine & "Reduce the search criteria to find wider match.")
                Exit Sub
            End If
            DataGridPM.DataSource = ds
            DataGridPM.DataMember = "GageMasterEntry"
            If chkWithBkDwn.Checked = True Then chkWithBkDwn.Checked = False
            DataGridPM.Focus()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at butSearch." & Environment.NewLine &
                   "Please capture screen shot And notify Aik Koon. Thanks.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function searchDB(sID1 As String, sID2 As String, sID3 As String, sSN1 As String, sSN2 As String, sSN3 As String, sOwner As String, sStat As String) As String
        If sID3 = "" Then sID3 = "%"
        If sSN3 = "" Then sSN3 = "%"

        searchDB = "SELECT Eqp_ID,Eqp_Type,Eqp_SN,Eqp_Mod,Eqp_Line,Eqp_Freq,Eqp_Resp,Eqp_Status,Old_ID,Eqp_Owner FROM [GageMasterEntry] WHERE (Eqp_SN LIKE '" & sSN1 &
        "' OR Eqp_SN LIKE '" & sSN2 & "') AND (Eqp_ID LIKE '" & sID1 & "' OR Eqp_ID LIKE '" & sID2 &
         "') AND Eqp_Status LIKE '" & sStat & "' AND Eqp_Owner LIKE '" & sOwner & "'"

        Return searchDB
    End Function
    Private Sub txtUSMfgSN_TextChanged(sender As Object, e As EventArgs) Handles txtUSMfgSN.TextChanged
        Dim str As String = txtUSMfgSN.Text
        str = str.ToUpper
        txtUSMfgSN.Text = str
        txtUSMfgSN.SelectionStart = str.Length
    End Sub

    Private Sub txtUSOldID_TextChanged(sender As Object, e As EventArgs) Handles txtUSOldID.TextChanged
        Dim str As String = txtUSOldID.Text
        str = str.ToUpper
        txtUSOldID.Text = str
        txtUSOldID.SelectionStart = str.Length
    End Sub

    Private Sub txtUSModel_TextChanged(sender As Object, e As EventArgs) Handles txtUSModel.TextChanged
        Dim str As String = txtUSModel.Text
        str = str.ToUpper
        txtUSModel.Text = str
        txtUSModel.SelectionStart = str.Length
    End Sub

    Private Sub txtUSmfg_TextChanged(sender As Object, e As EventArgs) Handles txtUSmfg.TextChanged
        Dim str As String = txtUSmfg.Text
        str = str.ToUpper
        txtUSmfg.Text = str
        txtUSmfg.SelectionStart = str.Length
    End Sub

    Private Sub txtUSEqpID_TextChanged(sender As Object, e As EventArgs) Handles txtUSEqpID.TextChanged
        Dim str As String = txtUSEqpID.Text
        str = str.ToUpper
        txtUSEqpID.Text = str
        txtUSEqpID.SelectionStart = str.Length
    End Sub
    Private Sub txtEqpOwner_TextChanged(sender As Object, e As EventArgs) Handles txtEqpOwner.TextChanged
        Dim str As String = txtEqpOwner.Text
        str = str.ToUpper
        txtEqpOwner.Text = str
        txtEqpOwner.SelectionStart = str.Length
    End Sub

    Private Sub butViewBrkDwn_Click(sender As Object, e As EventArgs) Handles butViewBrkDwn.Click
        If DataGridPM.RowCount = 0 Then
            MsgBox("No data selected!")
            Exit Sub
        End If
        Dim viewBDType As String = DataGridPM.SelectedRows(0).Cells(1).Value.ToString
        Dim viewBDID As String = DataGridPM.SelectedRows(0).Cells(0).Value.ToString
        Dim dS As DataSet
        Form3.lblBrkDwnID.Text = viewBDID
        Form3.lblBrkDwnType.Text = viewBDType
        Dim getBDSQL = "Select ID,Fault, Start_Date, Eq_Status, Svc_Cost FROM [Fault] WHERE Eqp_ID='" & viewBDID & "'"

        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            dS = New DataSet
            tables = dS.Tables
            da = New OleDbDataAdapter(getBDSQL, myConnToAccess)
            da.Fill(dS, "Fault")
            myConnToAccess.Close()

            If dS.Tables(0).Rows.Count = 0 Then
                MsgBox("No Breakdown records found for " & viewBDID & ".")
                Exit Sub
            End If

            Form3.DataGridBrkDwn.DataSource = dS
            Form3.DataGridBrkDwn.DataMember = "Fault"
            Form3.txtFaultSymptom.Text = ""
            Form3.txtSoln.Text = ""

            getBDSQL = "SELECT [Master_Lookup].Notify_ID FROM [Master_Lookup] LEFT JOIN [GageMasterEntry] ON" &
                " [GageMasterEntry].Eqp_Line=[Master_Lookup].Eqp_Line WHERE [GageMasterEntry].Eqp_ID='" & viewBDID & "'"
            myConnToAccess.Open()
            Dim dN As DataSet = New DataSet
            Dim dNote As OleDbDataAdapter = New OleDbDataAdapter(getBDSQL, myConnToAccess)
            dNote.Fill(dN, "Master_Lookup")
            myConnToAccess.Close()
            Form3.DataGridNotify.DataSource = dN
            Form3.DataGridNotify.DataMember = "Master_Lookup"
            Form3.ShowDialog()
            System.Windows.Forms.Cursor.Current = Cursors.Default

        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at view breakdown event from search tab." & Environment.NewLine &
                   "Please capture screen shot and notify Aik Koon. Thanks.")
            System.Windows.Forms.Cursor.Current = Cursors.Default
            myConnToAccess.Close()
        End Try
    End Sub

    Private Sub txtEqpOwner_Leave(sender As Object, e As EventArgs) Handles txtEqpOwner.Leave
        If txtEqpOwner.Text <> "" Then
            txtEqpOwner.Text = Trim(txtEqpOwner.Text.ToUpper)
            Dim chkUserID As Boolean = txtEqpOwner.Text Like "H*"
            If chkUserID = False Or Len(txtEqpOwner.Text) <> 7 Then
                Label29.ForeColor = Color.Red
                MsgBox("The owner user ID has to be the network user ID. Please correct before proceeding. Thanks.")
                Label29.ForeColor = Color.Black
            End If
        End If
    End Sub

    Private Sub txtnewEqpOwner_Leave(sender As Object, e As EventArgs) Handles txtnewEqpOwner.Leave
        If txtnewEqpOwner.Text <> "" Then
            Dim chknewOwner As Boolean = Trim(txtnewEqpOwner.Text.ToUpper) Like "H*"
            If chknewOwner = False Or Len(txtnewEqpOwner.Text) <> 7 Then
                Label32.ForeColor = Color.Red
                MsgBox("Please enter a valid network user ID before proceeding. Thanks.")
                Label32.ForeColor = Color.Black
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles butSelFolder.Click
        Dim folderDlg As New FolderBrowserDialog
        folderDlg.ShowNewFolderButton = True
        If (folderDlg.ShowDialog() = DialogResult.OK) Then
            txtBkUpFolder.Text = folderDlg.SelectedPath
            Dim root As Environment.SpecialFolder = folderDlg.RootFolder
        End If

    End Sub


    Private Sub MoveAllItems(ByVal fromPath As String, ByVal toPath As String)
        ''Create the target directory if necessary
        Dim toPathInfo = New DirectoryInfo(toPath)
        If (Not toPathInfo.Exists) Then
            toPathInfo.Create()
        End If
        Dim fromPathInfo = New DirectoryInfo(fromPath)
        ''move all files
        For Each file As FileInfo In fromPathInfo.GetFiles()
            file.MoveTo(Path.Combine(toPath, file.Name))
        Next
        ''move all folders
        For Each dir As DirectoryInfo In fromPathInfo.GetDirectories()
            dir.MoveTo(Path.Combine(toPath, dir.Name))
        Next
    End Sub

    Private Sub butBUdb_Click(sender As Object, e As EventArgs) Handles butBUdb.Click
        If txtBkUpFolder.Text = "" Then
            Label34.ForeColor = Color.Red
            MsgBox("Please select a destination folder to backup database.")
            Label34.ForeColor = Color.Black
            Exit Sub
        End If
        Dim curFolder = My.Settings.pm_Path.ToString
        curFolder = Mid(curFolder, 1, Len(curFolder) - 1)
        Dim tgtFolder = txtBkUpFolder.Text
        My.Computer.FileSystem.CopyDirectory(curFolder, tgtFolder)
        MsgBox("Backup to " & tgtFolder & " done!")
        txtBkUpFolder.Text = ""
    End Sub

    Private Sub butUpdatePMDue_Click(sender As Object, e As EventArgs) Handles butUpdatePMDue.Click
        If txtUSEqpID.Text = "" Then
            Label17.ForeColor = Color.Red
            MsgBox("No eqpm ID selected for PM Due update.")
            Label17.ForeColor = Color.Black
            Exit Sub
        End If
        If cmbUSFreq1.SelectedIndex = -1 Then
            Label31.ForeColor = Color.Red
            MsgBox("Please indicate the PM Frequency!")
            Label31.ForeColor = Color.Black
            Exit Sub
        End If
        Try
            'Get the PM frequency info

            dtDue.Format = DateTimePickerFormat.Custom
            dtDue.CustomFormat = "d/M/yyyy"
            Dim pmFreq As String = cmbUSFreq1.SelectedItem.ToString
            Dim lastPMdate As Date = dtDue.Value
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


            Dim pmSQL As String = "UPDATE [GageMasterEntry] SET Eqp_Due='" & nxDue.ToShortDateString & "' WHERE Eqp_ID='" & txtUSEqpID.Text & "'"
            myConnToAccess.Open()
            Dim newPM As New OleDbCommand(pmSQL, myConnToAccess)
            newPM.ExecuteNonQuery()
            myConnToAccess.Close()
            MsgBox("PM Due Date updated.")
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at updating new PM_Due." & Environment.NewLine &
                    "Please capture screenshot of this error and email Aik Koon. Thanks.")
            myConnToAccess.Close()
        End Try
    End Sub

    Private Sub dtDue_ValueChanged(sender As Object, e As EventArgs) Handles dtDue.ValueChanged
        Dim nowDate As Date = Now
        If dtDue.Value > nowDate Then
            Label36.ForeColor = Color.Red
            MsgBox("Last done date cannot be later than current date!")
            Label36.ForeColor = Color.Black
            Exit Sub
        End If
    End Sub

    Private Sub butgetPMdata_Click(sender As Object, e As EventArgs) Handles butgetPMdata.Click
        If cmbEqpArea.SelectedIndex = -1 Then
            Label40.ForeColor = Color.Red
            MsgBox("Please select the production area to retrieve summary.")
            Label40.ForeColor = Color.Black
            Exit Sub
        End If
        Try
            'Dim myCursor As String = pmPath & "RollingWait2.ani"
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'Me.Cursor = New Cursor(myCursor)
            Dim selArea As String = cmbEqpArea.SelectedValue.ToString
            Dim gD As DataSet = New DataSet
            'Dim getPMSQL As String = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
            '"[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Owner,[GageMasterEntry].Eqp_Due,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] INNER JOIN [Master_Lookup] ON " &
            '"[GageMasterEntry].Eqp_Line=[Master_Lookup].Eqp_Line WHERE [Master_Lookup].Prod_Area='" & selArea & "' AND ([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
            '" [GageMasterEntry].Eqp_Status='ACTIVE')"
            Dim getPMSQL As String
            If chkCal.Checked = False And chkPortEqp.Checked = False Then
                getPMSQL = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
            "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] INNER JOIN [Master_Lookup] ON " &
        "[GageMasterEntry].Eqp_Line=[Master_Lookup].Eqp_Line WHERE [Master_Lookup].Prod_Area='" & selArea & "' AND ([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
        " [GageMasterEntry].Eqp_Status='ACTIVE' OR [GageMasterEntry].Eqp_Status='REMOVE' OR [GageMasterEntry].Eqp_Status='INACTIVE')"
            ElseIf chkCal.Checked = True And selArea = "SPERRY" Then
                getPMSQL = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
            "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] WHERE (Eqp_ID LIKE 'HS2-SP-%' OR Eqp_ID LIKE 'HS2-TEST-%') AND" &
            " Eqp_ID NOT LIKE 'HS2-SP-TECH%' AND ([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
        " [GageMasterEntry].Eqp_Status='ACTIVE' OR [GageMasterEntry].Eqp_Status='REMOVE' OR [GageMasterEntry].Eqp_Status='INACTIVE')"
            ElseIf chkCal.Checked = True And selArea = "WIRELINE" Then
                getPMSQL = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
            "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] WHERE Eqp_ID LIKE 'HS2-WL-%' AND Eqp_ID NOT LIKE 'HS2-WL-TECH%' AND " &
                "([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
        " [GageMasterEntry].Eqp_Status='ACTIVE' OR [GageMasterEntry].Eqp_Status='REMOVE' OR [GageMasterEntry].Eqp_Status='INACTIVE')"
            ElseIf chkCal.Checked = True And selArea = "SPERRY-TECH" Then
                getPMSQL = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
            "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] WHERE Eqp_ID LIKE 'HS2-SP-TECH-%' AND ([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
        " [GageMasterEntry].Eqp_Status='ACTIVE' OR [GageMasterEntry].Eqp_Status='REMOVE' OR [GageMasterEntry].Eqp_Status='INACTIVE')"
            ElseIf chkCal.Checked = True And selArea = "WIRELINE-TECH" Then
                getPMSQL = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
            "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] WHERE Eqp_ID LIKE 'HS2-WL-TECH-%' AND ([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
        " [GageMasterEntry].Eqp_Status='ACTIVE' OR [GageMasterEntry].Eqp_Status='REMOVE' OR [GageMasterEntry].Eqp_Status='INACTIVE')"
            ElseIf chkCal.Checked = True And selArea = "QA" Then
                getPMSQL = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
            "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] WHERE Used_By='QAQC' AND ([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
        " [GageMasterEntry].Eqp_Status='ACTIVE' OR [GageMasterEntry].Eqp_Status='REMOVE' OR [GageMasterEntry].Eqp_Status='INACTIVE')"
            ElseIf chkPortEqp.Checked = True
                getPMSQL = "SELECT [GageMasterEntry].Eqp_ID,[GageMasterEntry].Eqp_Type,[GageMasterEntry].Eqp_SN,[GageMasterEntry].Eqp_Mod,[GageMasterEntry].Eqp_Line," &
                "[GageMasterEntry].Eqp_Freq,[GageMasterEntry].Eqp_Owner,[GageMasterEntry].Eqp_Due,[GageMasterEntry].Eqp_Status FROM [GageMasterEntry] INNER JOIN [Master_Lookup] ON " &
                "[GageMasterEntry].Eqp_Line=[Master_Lookup].Eqp_Line WHERE [Master_Lookup].Prod_Area='" & selArea & "' AND ([GageMasterEntry].Eqp_Status='B-DOWN' OR" &
                " [GageMasterEntry].Eqp_Status='ACTIVE') AND [GageMasterEntry].Eqp_Type='PORTABLE-EQP'"
            End If
            Dim gA As New OleDbDataAdapter(getPMSQL, myConnToAccess)
            gA.Fill(gD, "GageMasterEntry")
            DataGridViewPM.DataSource = gD
            DataGridViewPM.DataMember = "GageMasterEntry"
            If DataGridViewPM.Rows.Count = 0 Then
                MsgBox("No Records Found.")
                Exit Sub
            End If
            Call chkExpire(selArea)
            System.Windows.Forms.Cursor.Current = Cursors.Default

        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error getting PM summary data." & Environment.NewLine &
                    "Please capture screenshot and email Aik Koon. Thank You.")
            myConnToAccess.Close()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub chkExpire(sArea As String)
        Dim dgRow As Integer
        'Dim dueDate As Date
        Dim pmStat As String
        Dim pmFreq As String
        Dim datePM As String
        Dim dateCult As String = Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
        Dim Mpos As Integer = InStr(dateCult, "M")
        Dim totCount As Integer = DataGridViewPM.Rows.Count + 1
        Dim totDue As Integer = 0
        Dim totBkDwn As Integer = 0
        Dim totInactive As Integer = 0
        Dim sCol As Integer
        If chkPortEqp.Checked = True Then
            sCol = 8
        Else
            sCol = 6
        End If

        For dgRow = 0 To DataGridViewPM.Rows.Count - 1
            pmStat = DataGridViewPM.Rows(dgRow).Cells(sCol).Value.ToString
            pmFreq = DataGridViewPM.Rows(dgRow).Cells(5).Value.ToString

            If pmStat = "B-DOWN" Then
                DataGridViewPM.Rows(dgRow).Cells(sCol).Style.BackColor = Color.Red
                totBkDwn = totBkDwn + 1
            ElseIf pmStat = "ACTIVE" Then
                DataGridViewPM.Rows(dgRow).Cells(sCol).Style.BackColor = Color.Green
            ElseIf sArea = "WIRELINE" And (pmStat = "REMOVE" Or pmStat = "INACTIVE") Then
                DataGridViewPM.Rows(dgRow).Cells(sCol).Value = "INACTIVE"
                DataGridViewPM.Rows(dgRow).Cells(sCol).Style.BackColor = Color.LightYellow
                totInactive = totInactive + 1
            ElseIf pmStat = "REMOVE" Or pmStat = "INACTIVE" Or pmStat = "NO PM" Then
                DataGridViewPM.Rows(dgRow).Cells(sCol).Value = "INACTIVE"
                DataGridViewPM.Rows(dgRow).Cells(sCol).Style.BackColor = Color.LightYellow
                totInactive = totInactive + 1

            End If
            'Rev8 change - bypass codes for due date check
            'GoTo Rev8Chg
            'Rev 12 change - reinstate the checking of due dates for portable equipment checks.
            If chkPortEqp.Checked = True Then
                If Not IsDBNull(DataGridViewPM.Rows(dgRow).Cells(7).Value) Then

                    datePM = DataGridViewPM.Rows(dgRow).Cells(7).Value.ToString
                    'Rev6 change***
                    If Mpos = 1 Then
                        datePM = Mid(datePM, InStr(datePM, "/") + 1, Len(datePM) - 4 - (InStr(datePM, "/") + 1)) & "/" &
                            Mid(datePM, 1, InStr(datePM, "/") - 1) & "/" & Mid(datePM, Len(datePM) - 3, 4)
                    End If
                    '****************
                    Dim dueDate As Date = DateTime.ParseExact(datePM, dateCult, CultureInfo.InvariantCulture)
                    If Now > dueDate.AddDays(-7) And Now < dueDate And pmStat <> "B-DOWN" And pmFreq <> "CONDITIONAL" Then
                        DataGridViewPM.Rows(dgRow).Cells(8).Value = "PM-DUE"
                        DataGridViewPM.Rows(dgRow).Cells(8).Style.BackColor = Color.Yellow
                    ElseIf dueDate < Now And pmFreq <> "CONDITIONAL" Then
                        DataGridViewPM.Rows(dgRow).Cells(8).Value = "OVERDUED"
                        DataGridViewPM.Rows(dgRow).Cells(8).Style.BackColor = Color.Red
                        DataGridViewPM.Rows(dgRow).Cells(7).Style.BackColor = Color.Red
                        totDue = totDue + 1
                    ElseIf pmFreq = "CONDITIONAL" Then
                        DataGridViewPM.Rows(dgRow).Cells(8).Value = "NA"
                    End If
                Else
                    DataGridViewPM.Rows(dgRow).Cells(7).Style.BackColor = Color.Red
                End If
            End If
Rev8Chg:
        Next
        If chkPortEqp.Checked = True Then
            lblTotPMEqm.Text = "Total Portable Eqp in " & sArea & ": " & totCount
            lblTotBrkDwn.Text = "Total due Portable Eqp in " & sArea & ": " & totDue
            lblTotDue.Text = ""
        ElseIf chkPortEqp.Checked = False And sArea <> "WIRELINE" Then
            lblTotPMEqm.Text = "Total Eqm with PM in " & sArea & ": " & totCount
            lblTotBrkDwn.Text = "Total Breakdown Eqm in " & sArea & ": " & totBkDwn
            lblTotDue.Text = "Total Eqm no PM required in " & sArea & ": " & totInactive
        ElseIf chkPortEqp.Checked = False And sArea = "WIRELINE" Then
            lblTotPMEqm.Text = "Total Eqm with PM in " & sArea & ": " & totCount
            lblTotBrkDwn.Text = "Total Breakdown Eqm in " & sArea & ": " & totBkDwn
            lblTotDue.Text = "Total Inactive Eqm in " & sArea & ": " & totInactive
        End If
    End Sub

    Private Sub DataGridViewPM_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewPM.ColumnHeaderMouseClick
        DataGridViewPM.ClearSelection()
        Call chkExpire(cmbEqpArea.SelectedValue.ToString)
    End Sub

    Private Sub txtAddNewFault_TextChanged(sender As Object, e As EventArgs) Handles txtAddNewFault.TextChanged
        Dim str As String = txtAddNewFault.Text
        str = str.ToUpper
        txtAddNewFault.Text = str
        txtAddNewFault.SelectionStart = str.Length
    End Sub

    Private Sub butAddFault_Click(sender As Object, e As EventArgs) Handles butAddFault.Click
        If txtAddNewFault.Text = "" Then
            Label41.ForeColor = Color.Red
            MsgBox("PLease enter new fault type.")
            Label41.ForeColor = Color.Black
            Exit Sub
        End If
        Try
            'check if new fault type exist
            DataGridViewdbInfo.DataSource = Nothing
            Dim newFault As String = txtAddNewFault.Text
            Dim aN As DataSet = New DataSet
            Dim aNSQL As String = "SELECT Fault_Type FROM [Fault-Type] WHERE Fault_Type='" & newFault & "'"
            Dim aNQuery As New OleDbDataAdapter(aNSQL, myConnToAccess)
            aNQuery.Fill(aN, "Fault-Type")
            myConnToAccess.Close()
            If aN.Tables(0).Rows.Count > 0 Then
                myConnToAccess.Open()
                aNSQL = "SELECT * FROM [Fault-Type]"
                Dim anQuery2 As New OleDbDataAdapter(aNSQL, myConnToAccess)
                anQuery2.Fill(aN, "Fault-Type")
                DataGridViewdbInfo.DataSource = aN
                DataGridViewdbInfo.DataMember = "Fault-Type"
                myConnToAccess.Close()
                MsgBox(newFault & " already exist.")
                Exit Sub
            End If

            aNSQL = "INSERT INTO [Fault-Type] (Fault_Type) VALUES ('" & newFault & "')"
            myConnToAccess.Open()
            Dim anCom As OleDbCommand = New OleDbCommand(aNSQL, myConnToAccess)
            anCom.ExecuteNonQuery()
            myConnToAccess.Close()
            MsgBox("New Fault-Type: " & newFault & " added into database.")
            txtAddNewFault.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at adding new fault type." & Environment.NewLine &
                    "Please capture screen shot of error message and email Aik Koon. Thank You.")
            myConnToAccess.Close()
        End Try

    End Sub

    Private Sub butAddStat_Click(sender As Object, e As EventArgs) Handles butAddStat.Click
        If txtaddNewStat.Text = "" Then
            Label42.ForeColor = Color.Red
            MsgBox("Please enter new status.")
            Label42.ForeColor = Color.Black
            Exit Sub
        End If

        Try
            DataGridViewdbInfo.DataSource = Nothing
            Dim newStat As String = txtaddNewStat.Text
            Dim aStat As DataSet = New DataSet
            Dim aSSQL As String = "SELECT Stat_Type FROM [Stat_Type] WHERE Stat_Type='" & newStat & "'"
            Dim aSQuery As OleDbDataAdapter = New OleDbDataAdapter(aSSQL, myConnToAccess)
            aSQuery.Fill(aStat, "Stat_Type")
            myConnToAccess.Close()
            If aStat.Tables(0).Rows.Count > 0 Then
                myConnToAccess.Open()
                aSSQL = "SELECT *  FROM [Stat_Type]"
                Dim asQuery1 As New OleDbDataAdapter(aSSQL, myConnToAccess)
                asQuery1.Fill(aStat, "Stat_Type")
                DataGridViewdbInfo.DataSource = aStat
                DataGridViewdbInfo.DataMember = "Stat_Type"
                myConnToAccess.Close()
                MsgBox(newStat & " already exist!")
                Exit Sub
            End If

            aSSQL = "INSERT INTO [Stat_Type] (Stat_Type) VALUES('" & newStat & "')"
            myConnToAccess.Open()
            Dim aSCom As New OleDbCommand(aSSQL, myConnToAccess)
            aSCom.ExecuteNonQuery()
            myConnToAccess.Close()
            MsgBox(newStat & " Added into database.")
        Catch ex As Exception
            MsgBox(ex.Message & Environment.NewLine & "Error at adding new status!" & Environment.NewLine &
                    "Please capture screen shot of error and email Aik Koon. Thank You")
        End Try
    End Sub

    Private Sub DataGridViewPM_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewPM.CellClick
        'Rev8 - 
        Dim selArea As String = cmbEqpArea.SelectedValue.ToString
        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        chkExpire(selArea)
        System.Windows.Forms.Cursor.Current = Cursors.Default
        'Exit Sub
        'Rev12 change reinstate check due 
        If chkPortEqp.Checked = True Then
            Dim cRow As Integer = e.RowIndex
            Dim cCol As Integer = e.ColumnIndex
            If cCol <> 7 Or cRow = -1 Then
                Exit Sub
            End If
            Dim eqFreq As String = DataGridViewPM.Rows(cRow).Cells(5).Value.ToString 'Eqp_Freq
            If eqFreq = "CONDITIONAL" Then
                MsgBox("No PM due date for this equipment.", MsgBoxStyle.Information)
                Exit Sub
            End If
            Form4.Label2.Text = DataGridViewPM.Rows(cRow).Cells(0).Value.ToString 'Eqp_ID
            Form4.Label3.Text = DataGridViewPM.Rows(cRow).Cells(4).Value.ToString & " in: " 'Eqp_Area
            Form4.Label5.Text = cmbEqpArea.SelectedValue.ToString
            Form4.Label4.Text = DataGridViewPM.Rows(cRow).Cells(1).Value.ToString 'Eqp_Line
            Form4.lblFreq.Text = DataGridViewPM.Rows(cRow).Cells(5).Value.ToString 'Eqp Freq
            Form4.ShowDialog()
        End If
    End Sub

    Private Sub chkWithBkDwn_CheckedChanged(sender As Object, e As EventArgs) Handles chkWithBkDwn.CheckedChanged
        If chkWithBkDwn.Checked = True Then
            MsgBox("All other search criteria will be ignored if 'Search all with breakdown history' is selected.")
        End If
    End Sub

    Private Sub DataGridPM_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridPM.CellClick
        If e.RowIndex < 0 Then
            Exit Sub
        End If
        Dim sRow As Integer = e.RowIndex
        txtUSEqpID.Text = DataGridPM.Rows(sRow).Cells(0).Value.ToString
        txtUSEqpID.ReadOnly = True
        cmbUSEqpDesc.SelectedValue = DataGridPM.Rows(sRow).Cells(1).Value.ToString
        txtUSMfgSN.Text = DataGridPM.Rows(sRow).Cells(2).Value.ToString
        txtUSmfg.Text = DataGridPM.Rows(sRow).Cells(3).Value.ToString
        txtUSModel.Text = DataGridPM.Rows(sRow).Cells(4).Value.ToString
        cmbUSEqp_Line.SelectedValue = DataGridPM.Rows(sRow).Cells(5).Value.ToString
        cmbUSFreq1.SelectedItem = DataGridPM.Rows(sRow).Cells(6).Value.ToString
        cmbUSResp1.SelectedItem = DataGridPM.Rows(sRow).Cells(7).Value.ToString
        cmbUSEqpStat.SelectedValue = DataGridPM.Rows(sRow).Cells(8).Value.ToString
        txtUSOldID.Text = DataGridPM.Rows(sRow).Cells(9).Value.ToString
        txtEqpOwner.Text = DataGridPM.Rows(sRow).Cells(10).Value.ToString
    End Sub

    Private Sub butClrtab1_Click(sender As Object, e As EventArgs) Handles butClrtab0.Click
        Call clr_Tab0()
        Call tab0_Init()
    End Sub
    Private Sub clr_Tab0()
        cmbEqp_Name.SelectedIndex = -1
        cmbProd_Line.SelectedIndex = -1
        DataGridView1.DataSource = Nothing

    End Sub

    Private Sub tab0_Init()
        Dim ds_Line As DataSet
        Dim ds_Type As DataSet

        SQLquery = "SELECT DISTINCT Eqp_Line from [GageMasterEntry] WHERE Eqp_Line IS NOT NULL"
        myConnToAccess.Open()
        ds_Line = New DataSet
        tables = ds_Line.Tables
        da = New OleDbDataAdapter(SQLquery, myConnToAccess)
        da.Fill(ds_Line, "GageMasterEntry")
        'Dim view1 As New DataView(tables(0))
        With cmbProd_Line
            '.Items.Insert(0, String.Empty) 'Insert empty row to 1st line
            .DataSource = ds_Line.Tables("GageMasterEntry")
            .DisplayMember = "Eqp_Line"
            .ValueMember = "Eqp_Line"
            .SelectedIndex = 0
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With

        SQLquery = "SELECT DISTINCT Eqp_Type from [GageMasterEntry] WHERE Eqp_Line<>''"
        ds_Type = New DataSet
        dT = New OleDbDataAdapter(SQLquery, myConnToAccess)
        dT.Fill(ds_Type, "GageMasterEntry")
        With cmbEqp_Name
            .DataSource = ds_Type.Tables("GageMasterEntry")
            .DisplayMember = "Eqp_Type"
            .ValueMember = "Eqp_Type"
            .SelectedIndex = 0
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
        myConnToAccess.Close()
        cmbProd_Line.SelectedIndex = -1
        cmbEqp_Name.SelectedIndex = -1
    End Sub

    Private Sub butViewGMnew_Click(sender As Object, e As EventArgs) Handles butViewGMnew.Click
        Dim xlApp As Excel.Application
        Dim xlWB As Excel.Workbook


        Try
            'check if gage master file exist locally. if not copy over.
            If File.Exists(uPath & "GageMasterEntry.xlsx") Then
                My.Computer.FileSystem.DeleteFile(uPath & "GageMasterEntry.xlsx")
            End If

            Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pmPath & pmData & ";")

            AccessConn.Open()

            'New sheet in Workbook
            Dim AccessCommand As New System.Data.OleDb.OleDbCommand("SELECT * INTO [Excel 12.0 Xml;DATABASE=" & uPath & "GageMasterEntry.xlsx;HDR=Yes;].[GageMasterEntry] from [GageMasterEntry]", AccessConn)

            AccessCommand.ExecuteNonQuery()
            AccessConn.Close()

            xlApp = New Excel.Application
            xlWB = xlApp.Workbooks.Open(pmPath & "GageMasterEntry.xlsx")
            xlApp.Visible = True

            releaseObject(xlWB)
            releaseObject(xlApp)
        Catch ex As Exception
            'releaseObject(xlWB)
            'releaseObject(xlApp)
            MsgBox(ex.Message & Environment.NewLine & "Error at Masterlist view event." & Environment.NewLine &
                   "Please capture screen shot and notify Aik Koon. Thanks.")

        End Try
    End Sub

    Private Sub butHelp_Click(sender As Object, e As EventArgs) Handles butHelp.Click
        SearchHelp.Show()
    End Sub

    Private Sub butHelpMain_Click(sender As Object, e As EventArgs) Handles butHelpMain.Click
        helpMain.Show()
    End Sub

    Private Sub butPMTrackHelp_Click(sender As Object, e As EventArgs) Handles butPMTrackHelp.Click
        pmTrackHelp.Show()
    End Sub

    Private Sub butExpExcel_Click(sender As Object, e As EventArgs) Handles butExpExcel.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim uName As String = GetUserName()

        Dim i As Integer, j As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("Sheet1")

        If DataGridViewPM.RowCount <= 0 Then
            Exit Sub
        End If

        For i = 0 To DataGridViewPM.RowCount - 2
            For j = 0 To DataGridViewPM.ColumnCount - 1
                xlWorkSheet.Cells(i + 1, j + 1) = DataGridViewPM(j, i).Value.ToString()
            Next
        Next

        xlWorkBook.SaveAs("C:\Users\" & uName & "\Downloads\PSL-View.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        MessageBox.Show("File save as PSL-View.xls in C:\Users\" & uName & "\Downloads\PSL-View.xls")
    End Sub

    Private Sub chkPortEqp_CheckedChanged(sender As Object, e As EventArgs) Handles chkPortEqp.CheckedChanged
        If chkPortEqp.Checked = True Then
            chkCal.Checked = False

        End If
    End Sub

    Private Sub chkCal_CheckedChanged(sender As Object, e As EventArgs) Handles chkCal.CheckedChanged
        If chkCal.Checked = True Then
            chkPortEqp.Checked = False
        End If
    End Sub
End Class
