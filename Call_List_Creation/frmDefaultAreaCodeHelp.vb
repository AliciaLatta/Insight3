Public Class frmDefaultAreaCodeHelp
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
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDefaultAreaCodeHelp))
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblText = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(184, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Default Area Code"
        '
        'lblText
        '
        Me.lblText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblText.Location = New System.Drawing.Point(24, 72)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(432, 280)
        Me.lblText.TabIndex = 1
        '
        'frmDefaultAreaCodeHelp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(480, 382)
        Me.Controls.Add(Me.lblText)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDefaultAreaCodeHelp"
        Me.Text = "Default Area Code - Help"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmDefaultAreaCodeHelp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblText.Text = "The Default Area Code value will be inserted before any phone numbers in the report with only seven "
        lblText.Text += "numbers in order to make a ten digits.  If the Default Area Code is blank, any records with seven character phone numbers will be omitted from the call list.  "
        lblText.Text += "This value can only be changed in the config file by setting the value of 'DefaultAreaCode'."
    End Sub
End Class
