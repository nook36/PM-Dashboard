<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtConfigPasswrd = New System.Windows.Forms.TextBox()
        Me.butConfig = New System.Windows.Forms.Button()
        Me.tabIdx = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(365, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Enter Password to assess application configuration page"
        '
        'txtConfigPasswrd
        '
        Me.txtConfigPasswrd.Location = New System.Drawing.Point(28, 58)
        Me.txtConfigPasswrd.Name = "txtConfigPasswrd"
        Me.txtConfigPasswrd.Size = New System.Drawing.Size(377, 22)
        Me.txtConfigPasswrd.TabIndex = 1
        '
        'butConfig
        '
        Me.butConfig.Location = New System.Drawing.Point(174, 101)
        Me.butConfig.Name = "butConfig"
        Me.butConfig.Size = New System.Drawing.Size(75, 23)
        Me.butConfig.TabIndex = 2
        Me.butConfig.Text = "Enter"
        Me.butConfig.UseVisualStyleBackColor = True
        '
        'tabIdx
        '
        Me.tabIdx.AutoSize = True
        Me.tabIdx.Location = New System.Drawing.Point(375, 119)
        Me.tabIdx.Name = "tabIdx"
        Me.tabIdx.Size = New System.Drawing.Size(16, 17)
        Me.tabIdx.TabIndex = 3
        Me.tabIdx.Text = "?"
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(431, 146)
        Me.Controls.Add(Me.tabIdx)
        Me.Controls.Add(Me.butConfig)
        Me.Controls.Add(Me.txtConfigPasswrd)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form2"
        Me.Text = "Assess Application Configuration"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents txtConfigPasswrd As TextBox
    Friend WithEvents butConfig As Button
    Friend WithEvents tabIdx As Label
End Class
