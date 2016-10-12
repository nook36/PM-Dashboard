<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmAddNewBD
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddNewBD))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblNewRec = New System.Windows.Forms.Label()
        Me.cmbFaultCat = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtStart = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtAddNewBD = New System.Windows.Forms.TextBox()
        Me.butAddBD = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 7)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(233, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Add new breakdown record for: "
        '
        'lblNewRec
        '
        Me.lblNewRec.AutoSize = True
        Me.lblNewRec.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNewRec.Location = New System.Drawing.Point(250, 7)
        Me.lblNewRec.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNewRec.Name = "lblNewRec"
        Me.lblNewRec.Size = New System.Drawing.Size(19, 20)
        Me.lblNewRec.TabIndex = 1
        Me.lblNewRec.Text = "?"
        '
        'cmbFaultCat
        '
        Me.cmbFaultCat.BackColor = System.Drawing.SystemColors.Window
        Me.cmbFaultCat.CausesValidation = False
        Me.cmbFaultCat.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbFaultCat.FormattingEnabled = True
        Me.cmbFaultCat.Location = New System.Drawing.Point(132, 37)
        Me.cmbFaultCat.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmbFaultCat.Name = "cmbFaultCat"
        Me.cmbFaultCat.Size = New System.Drawing.Size(246, 28)
        Me.cmbFaultCat.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(9, 40)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(122, 20)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Fault-Category: "
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dtStart
        '
        Me.dtStart.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtStart.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtStart.Location = New System.Drawing.Point(460, 37)
        Me.dtStart.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtStart.Name = "dtStart"
        Me.dtStart.Size = New System.Drawing.Size(116, 26)
        Me.dtStart.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(382, 41)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(91, 20)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Start Date: "
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtAddNewBD
        '
        Me.txtAddNewBD.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddNewBD.Location = New System.Drawing.Point(13, 86)
        Me.txtAddNewBD.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtAddNewBD.Multiline = True
        Me.txtAddNewBD.Name = "txtAddNewBD"
        Me.txtAddNewBD.Size = New System.Drawing.Size(564, 167)
        Me.txtAddNewBD.TabIndex = 7
        '
        'butAddBD
        '
        Me.butAddBD.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butAddBD.Location = New System.Drawing.Point(13, 268)
        Me.butAddBD.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.butAddBD.Name = "butAddBD"
        Me.butAddBD.Size = New System.Drawing.Size(78, 43)
        Me.butAddBD.TabIndex = 8
        Me.butAddBD.Text = "Add"
        Me.butAddBD.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 70)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(117, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Breakdown Description"
        '
        'frmAddNewBD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightBlue
        Me.ClientSize = New System.Drawing.Size(592, 332)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.butAddBD)
        Me.Controls.Add(Me.txtAddNewBD)
        Me.Controls.Add(Me.dtStart)
        Me.Controls.Add(Me.cmbFaultCat)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblNewRec)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddNewBD"
        Me.Text = "Add New Breakdown Record"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents lblNewRec As Label
    Friend WithEvents cmbFaultCat As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents dtStart As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents txtAddNewBD As TextBox
    Friend WithEvents butAddBD As Button
    Friend WithEvents Label4 As Label
End Class
