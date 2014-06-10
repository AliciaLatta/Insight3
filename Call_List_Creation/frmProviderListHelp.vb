Public Class frmProviderListHelp
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
    Friend WithEvents lblFormatText As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnLaunchExample As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmProviderListHelp))
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblText = New System.Windows.Forms.Label
        Me.lblFormatText = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnLaunchExample = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Provider List"
        '
        'lblText
        '
        Me.lblText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblText.Location = New System.Drawing.Point(24, 40)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(632, 168)
        Me.lblText.TabIndex = 1
        '
        'lblFormatText
        '
        Me.lblFormatText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormatText.Location = New System.Drawing.Point(24, 240)
        Me.lblFormatText.Name = "lblFormatText"
        Me.lblFormatText.Size = New System.Drawing.Size(632, 144)
        Me.lblFormatText.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(24, 216)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(208, 23)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Notes on CSV File Format"
        '
        'btnLaunchExample
        '
        Me.btnLaunchExample.Location = New System.Drawing.Point(496, 392)
        Me.btnLaunchExample.Name = "btnLaunchExample"
        Me.btnLaunchExample.Size = New System.Drawing.Size(152, 23)
        Me.btnLaunchExample.TabIndex = 4
        Me.btnLaunchExample.Text = "View CSV Example"
        '
        'frmProviderListHelp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(696, 470)
        Me.Controls.Add(Me.btnLaunchExample)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblFormatText)
        Me.Controls.Add(Me.lblText)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmProviderListHelp"
        Me.Text = "Provider List - Help"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmProviderListHelp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblText.Text = "The provider list tells the program how to match up the provider ids in the report with the provider ids in the IVR database.  "
        lblText.Text += "There are two options for where the program pulls the provider list.  "
        lblText.Text += "If the checkbox to use the CSV file is unchecked, the program will use the list of providers in the box on the user interface.  "
        lblText.Text += "Alternatively, if the checkbox indicating that the CSV file should be used "
        lblText.Text += "is checked, the program will look for a CSV file in the path supplied.  Using a CSV file allows an extra message to be added to the call.  "
        lblText.Text += "IVR can be set up to play a message with special instructions.  "
        lblText.Text += "For example, if the patient should fast 24 hours before the appointment, these instructions can be recorded in IVR.  "
        lblText.Text += "Each message has a unique id and that is what's held in the 'Message Tag ID' column in the CSV file.  "
        lblFormatText.Text += "The CSV file MUST be in the specified format and must be a '.csv' file.  (Note: .csv files may be viewed and updated in Excel.)  "
        lblFormatText.Text += "It is important that the first row of the .csv file contain the column headers, NOT ACTUAL PROVIDER DATA.  If your office will "
        lblFormatText.Text += "not be using special messages leave the 'Message Tag ID' column values blank BUT BE SURE TO KEEP THE COLUMN HEADER, 'Message Tag ID', AS THE LAST COLUMN HEADER OF THE FILE."
    End Sub

    Private Sub btnLaunchExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLaunchExample.Click
        Dim form As csvExample
        form = New csvExample
        form.Show()
        form.BringToFront()
    End Sub
End Class
