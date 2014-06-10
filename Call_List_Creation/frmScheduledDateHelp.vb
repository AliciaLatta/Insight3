Public Class frmScheduledDateHelp
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
    Friend WithEvents lblHeading As System.Windows.Forms.Label
    Friend WithEvents lblText As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmScheduledDateHelp))
        Me.lblHeading = New System.Windows.Forms.Label
        Me.lblText = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblHeading
        '
        Me.lblHeading.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeading.Location = New System.Drawing.Point(16, 16)
        Me.lblHeading.Name = "lblHeading"
        Me.lblHeading.Size = New System.Drawing.Size(264, 23)
        Me.lblHeading.TabIndex = 0
        Me.lblHeading.Text = "Scheduled Date"
        '
        'lblText
        '
        Me.lblText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblText.Location = New System.Drawing.Point(16, 56)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(520, 312)
        Me.lblText.TabIndex = 1
        '
        'frmScheduledDateHelp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(552, 390)
        Me.Controls.Add(Me.lblText)
        Me.Controls.Add(Me.lblHeading)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmScheduledDateHelp"
        Me.Text = "Help - Scheduled Date"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmScheduledDateHelp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblText.Text = "Appointments that were scheduled (based on the scheduled date of the report) within the number "
        lblText.Text += "of days specified in the texbox will not be called.  If you would prefer that "
        lblText.Text += "the program not consider the scheduled date in determining which appointments will be added "
        lblText.Text += "to the call list, LEAVE THE TEXTBOX BLANK.  Acceptable values for the textbox are numbers greater than zero.  "
        lblText.Text += "Entering 1 will result in appointments scheduled on the day of the appointment not being called.  "
        lblText.Text += "Following the same logic, by entering 2 in the textbox, appointments scheduled the day prior to the appointment date "
        lblText.Text += "will not be called."
    End Sub
End Class
