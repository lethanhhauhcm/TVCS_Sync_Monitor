<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TVCS
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TVCS))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TxtFeedBack = New System.Windows.Forms.TextBox()
        Me.TrayIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 5600
        '
        'TxtFeedBack
        '
        Me.TxtFeedBack.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFeedBack.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtFeedBack.Location = New System.Drawing.Point(1, 0)
        Me.TxtFeedBack.Multiline = True
        Me.TxtFeedBack.Name = "TxtFeedBack"
        Me.TxtFeedBack.Size = New System.Drawing.Size(458, 232)
        Me.TxtFeedBack.TabIndex = 1
        '
        'TrayIcon
        '
        Me.TrayIcon.Icon = CType(resources.GetObject("TrayIcon.Icon"), System.Drawing.Icon)
        Me.TrayIcon.Text = "TVCS"
        Me.TrayIcon.Visible = True
        '
        'TVCS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(460, 231)
        Me.Controls.Add(Me.TxtFeedBack)
        Me.MaximizeBox = False
        Me.Name = "TVCS"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: TVCS"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents TxtFeedBack As System.Windows.Forms.TextBox
    Friend WithEvents TrayIcon As System.Windows.Forms.NotifyIcon

End Class
