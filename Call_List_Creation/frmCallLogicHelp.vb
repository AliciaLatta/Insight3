Public Class frmCallLogicHelp
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
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCallLogicHelp))
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblText = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Call Logic"
        '
        'lblText
        '
        Me.lblText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblText.Location = New System.Drawing.Point(16, 56)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(424, 280)
        Me.lblText.TabIndex = 1
        '
        'frmCallLogicHelp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(456, 350)
        Me.Controls.Add(Me.lblText)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCallLogicHelp"
        Me.Text = "Call Logic - Help"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmCallLogicHelp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblText.Text = "The Call Logic value determines the overall logic of the program.  It can be set to two values: 'AllBut' or 'NoneBut'.  "
        lblText.Text += "When set to 'AllBut', all home phone numbers in the report will be eligible to be placed on the call list unless the letters 'PRIV' are found somewhere in the "
        lblText.Text += "home phone number field for that record in the report.  On the other hand, when set to 'NoneBut', the only phone numbers eligible for the call list are ones ending in 'OK'.  "
        lblText.Text += "Please note that the program is NOT case-sensitive.  'NONEBUT' is the same as 'NoneBut' and 'priv' is the same as 'PRIV', etc."
    End Sub
End Class
