<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form3
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.lbl = New System.Windows.Forms.Label()
        Me.DataGridBrkDwn = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFaultSymptom = New System.Windows.Forms.TextBox()
        Me.txtSoln = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.butAddBrkDwn = New System.Windows.Forms.Button()
        Me.lblBrkDwnID = New System.Windows.Forms.Label()
        Me.butUpdateBrkDwn = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblBrkDwnType = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtnewSupID = New System.Windows.Forms.TextBox()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.butUpdateNotify = New System.Windows.Forms.Button()
        Me.lblProdLine = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.DataGridNotify = New System.Windows.Forms.DataGridView()
        CType(Me.DataGridBrkDwn, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridNotify, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbl
        '
        Me.lbl.AutoSize = True
        Me.lbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl.Location = New System.Drawing.Point(24, 4)
        Me.lbl.Name = "lbl"
        Me.lbl.Size = New System.Drawing.Size(222, 25)
        Me.lbl.TabIndex = 0
        Me.lbl.Text = "Breakdown Record For: "
        '
        'DataGridBrkDwn
        '
        Me.DataGridBrkDwn.AllowUserToAddRows = False
        Me.DataGridBrkDwn.AllowUserToDeleteRows = False
        Me.DataGridBrkDwn.AllowUserToOrderColumns = True
        Me.DataGridBrkDwn.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridBrkDwn.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.DataGridBrkDwn.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridBrkDwn.Location = New System.Drawing.Point(29, 87)
        Me.DataGridBrkDwn.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DataGridBrkDwn.MultiSelect = False
        Me.DataGridBrkDwn.Name = "DataGridBrkDwn"
        Me.DataGridBrkDwn.ReadOnly = True
        Me.DataGridBrkDwn.RowTemplate.Height = 24
        Me.DataGridBrkDwn.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridBrkDwn.Size = New System.Drawing.Size(532, 345)
        Me.DataGridBrkDwn.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 66)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(381, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Select records below to view details or to add new / update."
        '
        'txtFaultSymptom
        '
        Me.txtFaultSymptom.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtFaultSymptom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFaultSymptom.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFaultSymptom.Location = New System.Drawing.Point(629, 87)
        Me.txtFaultSymptom.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtFaultSymptom.Multiline = True
        Me.txtFaultSymptom.Name = "txtFaultSymptom"
        Me.txtFaultSymptom.ReadOnly = True
        Me.txtFaultSymptom.Size = New System.Drawing.Size(563, 147)
        Me.txtFaultSymptom.TabIndex = 3
        '
        'txtSoln
        '
        Me.txtSoln.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtSoln.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSoln.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoln.Location = New System.Drawing.Point(629, 286)
        Me.txtSoln.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtSoln.Multiline = True
        Me.txtSoln.Name = "txtSoln"
        Me.txtSoln.ReadOnly = True
        Me.txtSoln.Size = New System.Drawing.Size(563, 147)
        Me.txtSoln.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(627, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(153, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Breakdown Description"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(627, 265)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(122, 17)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Corrective Actions"
        '
        'butAddBrkDwn
        '
        Me.butAddBrkDwn.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butAddBrkDwn.Location = New System.Drawing.Point(29, 452)
        Me.butAddBrkDwn.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.butAddBrkDwn.Name = "butAddBrkDwn"
        Me.butAddBrkDwn.Size = New System.Drawing.Size(104, 53)
        Me.butAddBrkDwn.TabIndex = 7
        Me.butAddBrkDwn.Text = "Add New Rec"
        Me.butAddBrkDwn.UseVisualStyleBackColor = True
        '
        'lblBrkDwnID
        '
        Me.lblBrkDwnID.AutoSize = True
        Me.lblBrkDwnID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBrkDwnID.Location = New System.Drawing.Point(259, 4)
        Me.lblBrkDwnID.Name = "lblBrkDwnID"
        Me.lblBrkDwnID.Size = New System.Drawing.Size(23, 25)
        Me.lblBrkDwnID.TabIndex = 8
        Me.lblBrkDwnID.Text = "?"
        '
        'butUpdateBrkDwn
        '
        Me.butUpdateBrkDwn.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butUpdateBrkDwn.Location = New System.Drawing.Point(160, 452)
        Me.butUpdateBrkDwn.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.butUpdateBrkDwn.Name = "butUpdateBrkDwn"
        Me.butUpdateBrkDwn.Size = New System.Drawing.Size(104, 53)
        Me.butUpdateBrkDwn.TabIndex = 9
        Me.butUpdateBrkDwn.Text = "Close Record"
        Me.butUpdateBrkDwn.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(24, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(111, 25)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Equipment:"
        '
        'lblBrkDwnType
        '
        Me.lblBrkDwnType.AutoSize = True
        Me.lblBrkDwnType.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBrkDwnType.Location = New System.Drawing.Point(149, 31)
        Me.lblBrkDwnType.Name = "lblBrkDwnType"
        Me.lblBrkDwnType.Size = New System.Drawing.Size(23, 25)
        Me.lblBrkDwnType.TabIndex = 11
        Me.lblBrkDwnType.Text = "?"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtnewSupID)
        Me.GroupBox1.Controls.Add(Me.Label32)
        Me.GroupBox1.Controls.Add(Me.butUpdateNotify)
        Me.GroupBox1.Controls.Add(Me.lblProdLine)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.DataGridNotify)
        Me.GroupBox1.Location = New System.Drawing.Point(28, 528)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(1172, 188)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "BreakDown Notification"
        '
        'txtnewSupID
        '
        Me.txtnewSupID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtnewSupID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnewSupID.Location = New System.Drawing.Point(115, 126)
        Me.txtnewSupID.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtnewSupID.Name = "txtnewSupID"
        Me.txtnewSupID.Size = New System.Drawing.Size(85, 30)
        Me.txtnewSupID.TabIndex = 46
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.Location = New System.Drawing.Point(21, 128)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(89, 25)
        Me.Label32.TabIndex = 47
        Me.Label32.Text = "User_ID:"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'butUpdateNotify
        '
        Me.butUpdateNotify.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butUpdateNotify.Location = New System.Drawing.Point(219, 117)
        Me.butUpdateNotify.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.butUpdateNotify.Name = "butUpdateNotify"
        Me.butUpdateNotify.Size = New System.Drawing.Size(104, 53)
        Me.butUpdateNotify.TabIndex = 10
        Me.butUpdateNotify.Text = "Add Notify"
        Me.butUpdateNotify.UseVisualStyleBackColor = True
        '
        'lblProdLine
        '
        Me.lblProdLine.AutoSize = True
        Me.lblProdLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProdLine.Location = New System.Drawing.Point(663, 27)
        Me.lblProdLine.Name = "lblProdLine"
        Me.lblProdLine.Size = New System.Drawing.Size(23, 25)
        Me.lblProdLine.TabIndex = 2
        Me.lblProdLine.Text = "?"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(21, 27)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(584, 25)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Below user IDs will be notified for this equipment in Production line:"
        '
        'DataGridNotify
        '
        Me.DataGridNotify.AllowUserToAddRows = False
        Me.DataGridNotify.AllowUserToDeleteRows = False
        Me.DataGridNotify.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridNotify.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.DataGridNotify.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridNotify.ColumnHeadersVisible = False
        Me.DataGridNotify.Location = New System.Drawing.Point(24, 60)
        Me.DataGridNotify.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DataGridNotify.Name = "DataGridNotify"
        Me.DataGridNotify.Size = New System.Drawing.Size(1127, 48)
        Me.DataGridNotify.TabIndex = 0
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.PowderBlue
        Me.ClientSize = New System.Drawing.Size(1221, 729)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblBrkDwnType)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.butUpdateBrkDwn)
        Me.Controls.Add(Me.lblBrkDwnID)
        Me.Controls.Add(Me.butAddBrkDwn)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSoln)
        Me.Controls.Add(Me.txtFaultSymptom)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridBrkDwn)
        Me.Controls.Add(Me.lbl)
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form3"
        Me.Text = "Equipment Breakdown Records"
        CType(Me.DataGridBrkDwn, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridNotify, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbl As Label
    Friend WithEvents DataGridBrkDwn As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents txtFaultSymptom As TextBox
    Friend WithEvents txtSoln As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents butAddBrkDwn As Button
    Friend WithEvents lblBrkDwnID As Label
    Friend WithEvents butUpdateBrkDwn As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents lblBrkDwnType As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents DataGridNotify As DataGridView
    Friend WithEvents butUpdateNotify As Button
    Friend WithEvents lblProdLine As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents txtnewSupID As TextBox
    Friend WithEvents Label32 As Label
End Class
