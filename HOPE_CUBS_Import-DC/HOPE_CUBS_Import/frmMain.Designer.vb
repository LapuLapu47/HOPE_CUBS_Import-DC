<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.lblRunTime1 = New System.Windows.Forms.Label()
        Me.lblRunTime2 = New System.Windows.Forms.Label()
        Me.lblError = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(12, 13)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(272, 244)
        Me.CheckedListBox1.TabIndex = 1
        '
        'lblRunTime1
        '
        Me.lblRunTime1.AutoSize = True
        Me.lblRunTime1.Location = New System.Drawing.Point(12, 282)
        Me.lblRunTime1.Name = "lblRunTime1"
        Me.lblRunTime1.Size = New System.Drawing.Size(88, 13)
        Me.lblRunTime1.TabIndex = 2
        Me.lblRunTime1.Text = "Import Run Time:"
        '
        'lblRunTime2
        '
        Me.lblRunTime2.AutoSize = True
        Me.lblRunTime2.Location = New System.Drawing.Point(250, 282)
        Me.lblRunTime2.Name = "lblRunTime2"
        Me.lblRunTime2.Size = New System.Drawing.Size(34, 13)
        Me.lblRunTime2.TabIndex = 4
        Me.lblRunTime2.Text = "00:00"
        '
        'lblError
        '
        Me.lblError.AutoSize = True
        Me.lblError.Location = New System.Drawing.Point(12, 316)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(0, 13)
        Me.lblError.TabIndex = 6
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(314, 359)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.lblRunTime2)
        Me.Controls.Add(Me.lblRunTime1)
        Me.Controls.Add(Me.CheckedListBox1)
        Me.Name = "frmMain"
        Me.Text = "H.O.P.E. - DC - Import"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents CheckedListBox1 As CheckedListBox
    Friend WithEvents lblRunTime1 As Label
    Friend WithEvents lblRunTime2 As Label
    Friend WithEvents lblError As Label
    Friend WithEvents Timer1 As Timer
End Class
