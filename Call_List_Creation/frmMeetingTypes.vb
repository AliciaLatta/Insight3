Public Class frmMeetingTypes
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblText As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMeetingTypes))
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblText = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Meeting Types"
        '
        'lblText
        '
        Me.lblText.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblText.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblText.Location = New System.Drawing.Point(24, 64)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(472, 280)
        Me.lblText.TabIndex = 1
        '
        'frmMeetingTypes
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(520, 374)
        Me.Controls.Add(Me.lblText)
        Me.Controls.Add(Me.Label1)
        Me.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMeetingTypes"
        Me.Text = "Meeting Types - Help"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmMeetingTypes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblText.Text = "Any appointments with a type listed in the box will not be called.  The values in the box are compared to the value found in the 'TYPE' column in the report.  "
        lblText.Text += "Meeting (appointment) types may be added to "
        lblText.Text += "or removed from the box via the user interface.  Types are not case-sensitive.  'New' is the same as 'NEW'.  "
    End Sub
End Class
