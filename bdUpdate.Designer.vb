<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class bdUpdate
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(bdUpdate))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblUpdateID = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtupdateFaultSymptom = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtUpdateSoln = New System.Windows.Forms.TextBox()
        Me.txtSvcCost = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSvcPO = New System.Windows.Forms.TextBox()
        Me.butUpdateCloseRec = New System.Windows.Forms.Button()
        Me.dtCloseRec = New System.Windows.Forms.DateTimePicker()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblRecID = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(262, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Close breakdown record for: "
        '
        'lblUpdateID
        '
        Me.lblUpdateID.AutoSize = True
        Me.lblUpdateID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUpdateID.Location = New System.Drawing.Point(277, 9)
        Me.lblUpdateID.Name = "lblUpdateID"
        Me.lblUpdateID.Size = New System.Drawing.Size(23, 25)
        Me.lblUpdateID.TabIndex = 9
        Me.lblUpdateID.Text = "?"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(153, 17)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Breakdown Description"
        '
        'txtupdateFaultSymptom
        '
        Me.txtupdateFaultSymptom.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtupdateFaultSymptom.Location = New System.Drawing.Point(17, 86)
        Me.txtupdateFaultSymptom.Multiline = True
        Me.txtupdateFaultSymptom.Name = "txtupdateFaultSymptom"
        Me.txtupdateFaultSymptom.Size = New System.Drawing.Size(563, 147)
        Me.txtupdateFaultSymptom.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 250)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(192, 17)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Corrective Actions / Solutions"
        '
        'txtUpdateSoln
        '
        Me.txtUpdateSoln.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUpdateSoln.Location = New System.Drawing.Point(17, 270)
        Me.txtUpdateSoln.Multiline = True
        Me.txtUpdateSoln.Name = "txtUpdateSoln"
        Me.txtUpdateSoln.Size = New System.Drawing.Size(563, 147)
        Me.txtUpdateSoln.TabIndex = 12
        '
        'txtSvcCost
        '
        Me.txtSvcCost.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSvcCost.Location = New System.Drawing.Point(197, 437)
        Me.txtSvcCost.Name = "txtSvcCost"
        Me.txtSvcCost.Size = New System.Drawing.Size(107, 30)
        Me.txtSvcCost.TabIndex = 14
        Me.txtSvcCost.Text = "0.00"
        Me.txtSvcCost.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(13, 437)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(161, 25)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Servicing Cost: $"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 488)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(133, 25)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Servicing PO:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtSvcPO
        '
        Me.txtSvcPO.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSvcPO.Location = New System.Drawing.Point(168, 488)
        Me.txtSvcPO.Name = "txtSvcPO"
        Me.txtSvcPO.Size = New System.Drawing.Size(136, 30)
        Me.txtSvcPO.TabIndex = 17
        '
        'butUpdateCloseRec
        '
        Me.butUpdateCloseRec.Location = New System.Drawing.Point(18, 537)
        Me.butUpdateCloseRec.Name = "butUpdateCloseRec"
        Me.butUpdateCloseRec.Size = New System.Drawing.Size(104, 53)
        Me.butUpdateCloseRec.TabIndex = 18
        Me.butUpdateCloseRec.Text = "Close Record"
        Me.butUpdateCloseRec.UseVisualStyleBackColor = True
        '
        'dtCloseRec
        '
        Me.dtCloseRec.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtCloseRec.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtCloseRec.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtCloseRec.Location = New System.Drawing.Point(426, 435)
        Me.dtCloseRec.Name = "dtCloseRec"
        Me.dtCloseRec.Size = New System.Drawing.Size(154, 30)
        Me.dtCloseRec.TabIndex = 19
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(310, 437)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 25)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Close Date: "
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(13, 35)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(105, 25)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Start Date:"
        '
        'lblStartDate
        '
        Me.lblStartDate.AutoSize = True
        Me.lblStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartDate.Location = New System.Drawing.Point(121, 34)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(23, 25)
        Me.lblStartDate.TabIndex = 22
        Me.lblStartDate.Text = "?"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(250, 35)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(109, 25)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "Record ID: "
        '
        'lblRecID
        '
        Me.lblRecID.AutoSize = True
        Me.lblRecID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRecID.Location = New System.Drawing.Point(363, 35)
        Me.lblRecID.Name = "lblRecID"
        Me.lblRecID.Size = New System.Drawing.Size(23, 25)
        Me.lblRecID.TabIndex = 24
        Me.lblRecID.Text = "?"
        '
        'bdUpdate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(596, 609)
        Me.Controls.Add(Me.lblRecID)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lblStartDate)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dtCloseRec)
        Me.Controls.Add(Me.butUpdateCloseRec)
        Me.Controls.Add(Me.txtSvcPO)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSvcCost)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtUpdateSoln)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtupdateFaultSymptom)
        Me.Controls.Add(Me.lblUpdateID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label7)
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "bdUpdate"
        Me.Text = "Close Equipment Breakown Record"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents lblUpdateID As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtupdateFaultSymptom As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtUpdateSoln As TextBox
    Friend WithEvents txtSvcCost As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents txtSvcPO As TextBox
    Friend WithEvents butUpdateCloseRec As Button
    Friend WithEvents dtCloseRec As DateTimePicker
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents lblStartDate As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents lblRecID As Label
End Class
