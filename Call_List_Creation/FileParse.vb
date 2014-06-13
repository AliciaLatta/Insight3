Imports System.Configuration
Public Class FileParse
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
    Friend WithEvents webBrowser As AxSHDocVw.AxWebBrowser
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FileParse))
        Me.webBrowser = New AxSHDocVw.AxWebBrowser
        CType(Me.webBrowser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'webBrowser
        '
        Me.webBrowser.Enabled = True
        Me.webBrowser.Location = New System.Drawing.Point(24, 16)
        Me.webBrowser.OcxState = CType(resources.GetObject("webBrowser.OcxState"), System.Windows.Forms.AxHost.State)
        Me.webBrowser.Size = New System.Drawing.Size(648, 440)
        Me.webBrowser.TabIndex = 0
        '
        'FileParse
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(688, 470)
        Me.Controls.Add(Me.webBrowser)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FileParse"
        Me.Text = "Web"
        CType(Me.webBrowser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FileParse_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        webBrowser.Navigate(ConfigurationManager.AppSettings("ParseURL").ToString)
    End Sub
End Class
