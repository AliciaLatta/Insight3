Imports System.Configuration
Imports System.IO
Imports System
Imports System.Net
Imports System.Data
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Resources
'Imports Rebex.Net
Imports System.Xml.Serialization
Imports System.Text.RegularExpressions
Imports Call_List_Creation.DataExport

Public Class HMM_IVR_Console
    Inherits System.Windows.Forms.Form
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
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
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents btnDeleteType As System.Windows.Forms.Button
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents lblTypeMsg As System.Windows.Forms.Label
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents txtNewType As System.Windows.Forms.TextBox
    Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents txtDaysPrior As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblMsg As System.Windows.Forms.Label
    Friend WithEvents btnFTP As System.Windows.Forms.Button
    Friend WithEvents btnCreateFile As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblCallLogic As System.Windows.Forms.Label
    Friend WithEvents ChooseCSV As System.Windows.Forms.Button
    Friend WithEvents lblCSVFile As System.Windows.Forms.Label
    Friend WithEvents chkUseCSV As System.Windows.Forms.CheckBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtAreaCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblInsightReport As System.Windows.Forms.Label
    Friend WithEvents ChooseInsightReport As System.Windows.Forms.Button
    Friend WithEvents txtCustID As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(HMM_IVR_Console))
        Me.Label15 = New System.Windows.Forms.Label()
        Me.btnDeleteType = New System.Windows.Forms.Button()
        Me.lblTypeMsg = New System.Windows.Forms.Label()
        Me.btnAddNew = New System.Windows.Forms.Button()
        Me.txtNewType = New System.Windows.Forms.TextBox()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.txtDaysPrior = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.btnFTP = New System.Windows.Forms.Button()
        Me.btnCreateFile = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lblCallLogic = New System.Windows.Forms.Label()
        Me.ChooseCSV = New System.Windows.Forms.Button()
        Me.chkUseCSV = New System.Windows.Forms.CheckBox()
        Me.lblCSVFile = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtCustID = New System.Windows.Forms.TextBox()
        Me.txtAreaCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblInsightReport = New System.Windows.Forms.Label()
        Me.ChooseInsightReport = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(207, 18)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(128, 23)
        Me.Label15.TabIndex = 36
        Me.Label15.Text = "Default Area Code:"
        '
        'btnDeleteType
        '
        Me.btnDeleteType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteType.Location = New System.Drawing.Point(235, 148)
        Me.btnDeleteType.Name = "btnDeleteType"
        Me.btnDeleteType.Size = New System.Drawing.Size(56, 23)
        Me.btnDeleteType.TabIndex = 57
        Me.btnDeleteType.Text = "Remove"
        '
        'lblTypeMsg
        '
        Me.lblTypeMsg.ForeColor = System.Drawing.Color.Red
        Me.lblTypeMsg.Location = New System.Drawing.Point(84, 153)
        Me.lblTypeMsg.Name = "lblTypeMsg"
        Me.lblTypeMsg.Size = New System.Drawing.Size(135, 34)
        Me.lblTypeMsg.TabIndex = 102
        Me.lblTypeMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnAddNew
        '
        Me.btnAddNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddNew.Location = New System.Drawing.Point(235, 116)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(56, 23)
        Me.btnAddNew.TabIndex = 101
        Me.btnAddNew.Text = ">"
        '
        'txtNewType
        '
        Me.txtNewType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewType.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.txtNewType.Location = New System.Drawing.Point(83, 116)
        Me.txtNewType.Name = "txtNewType"
        Me.txtNewType.Size = New System.Drawing.Size(136, 20)
        Me.txtNewType.TabIndex = 100
        '
        'ListBox2
        '
        Me.ListBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox2.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.ListBox2.Location = New System.Drawing.Point(82, 223)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(352, 95)
        Me.ListBox2.TabIndex = 99
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(79, 205)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 23)
        Me.Label7.TabIndex = 98
        Me.Label7.Text = "Provider List:"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(79, 92)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(279, 16)
        Me.Label5.TabIndex = 97
        Me.Label5.Text = "Appt/meeting types in the box will NOT be called:"
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.ListBox1.Location = New System.Drawing.Point(304, 116)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(160, 95)
        Me.ListBox1.TabIndex = 96
        '
        'txtDaysPrior
        '
        Me.txtDaysPrior.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDaysPrior.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.txtDaysPrior.Location = New System.Drawing.Point(307, 64)
        Me.txtDaysPrior.Name = "txtDaysPrior"
        Me.txtDaysPrior.Size = New System.Drawing.Size(24, 22)
        Me.txtDaysPrior.TabIndex = 93
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(79, 67)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(231, 23)
        Me.Label4.TabIndex = 92
        Me.Label4.Text = "No of days calls should be made prior to appt:"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(80, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 20)
        Me.Label2.TabIndex = 89
        Me.Label2.Text = "Customer ID:"
        '
        'lblMsg
        '
        Me.lblMsg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMsg.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblMsg.Location = New System.Drawing.Point(7, 453)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(656, 99)
        Me.lblMsg.TabIndex = 115
        Me.lblMsg.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnFTP
        '
        Me.btnFTP.BackColor = System.Drawing.Color.White
        Me.btnFTP.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFTP.Location = New System.Drawing.Point(303, 414)
        Me.btnFTP.Name = "btnFTP"
        Me.btnFTP.Size = New System.Drawing.Size(222, 27)
        Me.btnFTP.TabIndex = 114
        Me.btnFTP.Text = "Upload Call List to IVR Server"
        Me.btnFTP.UseVisualStyleBackColor = False
        '
        'btnCreateFile
        '
        Me.btnCreateFile.BackColor = System.Drawing.Color.White
        Me.btnCreateFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreateFile.Location = New System.Drawing.Point(85, 414)
        Me.btnCreateFile.Name = "btnCreateFile"
        Me.btnCreateFile.Size = New System.Drawing.Size(138, 27)
        Me.btnCreateFile.TabIndex = 113
        Me.btnCreateFile.Text = "Create Call List "
        Me.btnCreateFile.UseVisualStyleBackColor = False
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(80, 43)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 23)
        Me.Label11.TabIndex = 111
        Me.Label11.Text = "Call Logic:"
        '
        'lblCallLogic
        '
        Me.lblCallLogic.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCallLogic.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblCallLogic.Location = New System.Drawing.Point(143, 41)
        Me.lblCallLogic.Name = "lblCallLogic"
        Me.lblCallLogic.Size = New System.Drawing.Size(433, 23)
        Me.lblCallLogic.TabIndex = 110
        Me.lblCallLogic.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ChooseCSV
        '
        Me.ChooseCSV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChooseCSV.Location = New System.Drawing.Point(172, 327)
        Me.ChooseCSV.Name = "ChooseCSV"
        Me.ChooseCSV.Size = New System.Drawing.Size(34, 23)
        Me.ChooseCSV.TabIndex = 116
        Me.ChooseCSV.Text = "..."
        '
        'chkUseCSV
        '
        Me.chkUseCSV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseCSV.Location = New System.Drawing.Point(448, 229)
        Me.chkUseCSV.Name = "chkUseCSV"
        Me.chkUseCSV.Size = New System.Drawing.Size(176, 40)
        Me.chkUseCSV.TabIndex = 117
        Me.chkUseCSV.Text = "Use CSV File for Provider List"
        '
        'lblCSVFile
        '
        Me.lblCSVFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCSVFile.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblCSVFile.Location = New System.Drawing.Point(212, 327)
        Me.lblCSVFile.Name = "lblCSVFile"
        Me.lblCSVFile.Size = New System.Drawing.Size(402, 39)
        Me.lblCSVFile.TabIndex = 118
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(81, 332)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(85, 23)
        Me.Label17.TabIndex = 119
        Me.Label17.Text = "Provider List:"
        '
        'txtCustID
        '
        Me.txtCustID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustID.Location = New System.Drawing.Point(149, 16)
        Me.txtCustID.Name = "txtCustID"
        Me.txtCustID.Size = New System.Drawing.Size(52, 20)
        Me.txtCustID.TabIndex = 120
        '
        'txtAreaCode
        '
        Me.txtAreaCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAreaCode.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.txtAreaCode.Location = New System.Drawing.Point(303, 15)
        Me.txtAreaCode.Name = "txtAreaCode"
        Me.txtAreaCode.Size = New System.Drawing.Size(38, 22)
        Me.txtAreaCode.TabIndex = 121
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(82, 380)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 23)
        Me.Label1.TabIndex = 124
        Me.Label1.Text = "Insight Report:"
        '
        'lblInsightReport
        '
        Me.lblInsightReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInsightReport.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblInsightReport.Location = New System.Drawing.Point(213, 374)
        Me.lblInsightReport.Name = "lblInsightReport"
        Me.lblInsightReport.Size = New System.Drawing.Size(402, 37)
        Me.lblInsightReport.TabIndex = 123
        '
        'ChooseInsightReport
        '
        Me.ChooseInsightReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChooseInsightReport.Location = New System.Drawing.Point(173, 375)
        Me.ChooseInsightReport.Name = "ChooseInsightReport"
        Me.ChooseInsightReport.Size = New System.Drawing.Size(34, 23)
        Me.ChooseInsightReport.TabIndex = 122
        Me.ChooseInsightReport.Text = "..."
        '
        'HMM_IVR_Console
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(665, 605)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblInsightReport)
        Me.Controls.Add(Me.ChooseInsightReport)
        Me.Controls.Add(Me.txtAreaCode)
        Me.Controls.Add(Me.txtCustID)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.lblCSVFile)
        Me.Controls.Add(Me.chkUseCSV)
        Me.Controls.Add(Me.ChooseCSV)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.btnFTP)
        Me.Controls.Add(Me.btnCreateFile)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.lblCallLogic)
        Me.Controls.Add(Me.lblTypeMsg)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.txtNewType)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.txtDaysPrior)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnDeleteType)
        Me.Controls.Add(Me.Label15)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "HMM_IVR_Console"
        Me.Text = "Appointment Reminder Call List Creation Tool (Version 7.3.7)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    '  Dim configPath As String
    Dim cust As Customer
    Dim output As OutputResults
    Private Sub HMM_IVR_Console_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '   configPath = ConfigurationManager.AppSettings("AppConfigPath").ToString
        GetConfigValues()
        If ConfigurationManager.AppSettings("UseCSV").ToUpper = "TRUE" Then
            chkUseCSV.Checked = True
        End If
    End Sub
    Private Sub GetConfigValues()
        txtCustID.Text = Trim(ConfigurationManager.AppSettings("CustId").ToString)
        txtDaysPrior.Text = Trim(ConfigurationManager.AppSettings("DaysPrior").ToString)
        If Trim(ConfigurationManager.AppSettings("MeetingSkipTypeTotal").ToString) <> "" Then
            BuildSkipTypeListBox()
        Else
            lblMsg.ForeColor = System.Drawing.Color.Red
            lblMsg.Text = "The config file must have an accurate value for MeetingSkipTypeTotal."
        End If
       
        BuildProviderListBox()

        txtAreaCode.Text = Trim(ConfigurationManager.AppSettings("DefaultAreaCode").ToString).ToUpper
        lblCSVFile.Text = GetCurrentDirectory & "\" & Trim(ConfigurationManager.AppSettings("CSVFileName").ToString).ToUpper
        Me.lblInsightReport.Text = Trim(ConfigurationManager.AppSettings("ReportFilePath").ToString).ToUpper
        If Trim(ConfigurationManager.AppSettings("CallLogic").ToString.ToUpper) = "NONEBUT" Then
            lblCallLogic.Text = "Only numbers ending in OK will be called"
        ElseIf Trim(ConfigurationManager.AppSettings("CallLogic").ToString.ToUpper) = "ALLBUT" Then
            lblCallLogic.Text = "Patients will not receive a call if they have a PR in a patient class field."
        Else
            lblCallLogic.ForeColor = System.Drawing.Color.Red
            lblCallLogic.Text = "The value of CallLogic is not valid."
        End If
    End Sub
    Private Sub BuildSkipTypeListBox()
        Dim y As Integer
        Dim typeArray(CType(ConfigurationManager.AppSettings("MeetingSkipTypeTotal"), Integer)) As String
        Dim meeting As String
        y = 0
        Try
            Do Until y = CType(ConfigurationManager.AppSettings("MeetingSkipTypeTotal"), Integer)
                meeting = "MeetingSkipType" & (y + 1)
                If Trim(ConfigurationManager.AppSettings(meeting).ToString) <> "" Then
                    typeArray(y) = Trim(ConfigurationManager.AppSettings(meeting).ToString)
                Else
                    typeArray(y) = ""
                End If
                y += 1
            Loop

            Dim x As Integer
            Dim count As Integer
            x = 0
            count = 0
            Do Until x = CType(ConfigurationManager.AppSettings("MeetingSkipTypeTotal"), Integer)
                If typeArray(x).ToString() <> "" Then
                    ListBox1.Items.Add(typeArray(x))
                End If
                x += 1
            Loop
        Catch e As Exception
            lblMsg.ForeColor = System.Drawing.Color.Red
            lblMsg.Text = "There was an error populating the Meeting Skip Types.  Please review the config file for the problem."
        End Try
    End Sub
    Private Sub BuildProviderListBox()
        Dim providerIDs As Array
        Dim reader As StreamReader
        Dim line As String
        Try
            If File.Exists("..\ProviderList.txt") Then
                reader = New StreamReader("..\ProviderList.txt")
                Do Until reader.EndOfStream
                    line = reader.ReadLine
                    providerIDs = Split(line, ",")
                    ListBox2.Items.Add("Report Provider ID " & providerIDs(1))
                Loop
            End If
        Catch e As Exception
            lblMsg.ForeColor = System.Drawing.Color.Red
            lblMsg.Text = "There was an error populating the Providers.  Please review the config file for the problem."
        End Try
    End Sub

    Private Sub ReplaceConfigSettings(ByVal key As String, ByVal val As String)
        Dim xDoc As New Xml.XmlDataDocument
        Dim xNode As Xml.XmlNode
        Dim path As String() = Directory.GetFiles(GetCurrentDirectory, "*.config")
        xDoc.Load(path(0))
        'For Each xNode In xDoc 
        For Each xNode In xDoc("configuration")("appSettings")
            If (xNode.Name = "add") Then
                If (xNode.Attributes.GetNamedItem("key").Value = key) Then
                    xNode.Attributes.GetNamedItem("value").Value = val
                    Exit For
                End If
            End If
        Next
        xDoc.Save(path(0))
    End Sub

    Private Sub FTP()
        Dim host As String = ConfigurationManager.AppSettings("Server").ToString
        Dim username As String = ConfigurationManager.AppSettings("Username").ToString
        Dim password As String = ConfigurationManager.AppSettings("Password").ToString
        Dim callList As String = GetCurrentDirectory() & "\" & ConfigurationManager.AppSettings("CallListFileName").ToString
        Dim URI As String
        Dim ftp As System.Net.FtpWebRequest

        Try
            URI = host & "/" & My.Computer.FileSystem.GetFileInfo(callList).Name

            ftp = CType(FtpWebRequest.Create(URI), FtpWebRequest)

            'Make the connection secure
            'ftp.EnableSsl = True

            Dim clsRequest As System.Net.FtpWebRequest = _
                DirectCast(System.Net.WebRequest.Create(URI), System.Net.FtpWebRequest)

            clsRequest.Credentials = New System.Net.NetworkCredential(username, password)
            clsRequest.Method = System.Net.WebRequestMethods.Ftp.UploadFile

            ' read in file...
            Dim bFile() As Byte = System.IO.File.ReadAllBytes(callList)

            ' upload file...
            Dim clsStream As System.IO.Stream = _
                clsRequest.GetRequestStream()

            clsStream.Write(bFile, 0, bFile.Length)
            clsStream.Close()
            clsStream.Dispose()

            'Archive
            'Create archive folder if it doesn't already exist
            If Not My.Computer.FileSystem.DirectoryExists(GetCurrentDirectory() & "\Archive") Then
                My.Computer.FileSystem.CreateDirectory(GetCurrentDirectory() & "\Archive")
            End If
            Dim archive As String = GetCurrentDirectory() & "\Archive\" & My.Computer.FileSystem.GetFileInfo(callList).Name
            If My.Computer.FileSystem.FileExists(archive) Then
                My.Computer.FileSystem.DeleteFile(archive)
            End If

            If File.Exists(callList) Then File.Move(callList, archive)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
   
    Private Sub btnDeleteType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteType.Click
        If ListBox1.Items.Count <> 0 And ListBox1.SelectedIndex <> -1 Then
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            lblTypeMsg.ForeColor = System.Drawing.Color.Black
        End If
    End Sub
    Private Sub btnDefaultAreaCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmDefaultAreaCodeHelp
        form = New frmDefaultAreaCodeHelp
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub bntScheduledDateHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmScheduledDateHelp
        form = New frmScheduledDateHelp
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub btnCallLogicHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmCallLogicHelp
        form = New frmCallLogicHelp
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub btnMeetingTypesHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmMeetingTypes
        form = New frmMeetingTypes
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub btnProviderListHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmProviderListHelp
        form = New frmProviderListHelp
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub Save()
        ReplaceConfigSettings("CustId", txtCustID.Text)
        ReplaceConfigSettings("DefaultAreaCode", txtAreaCode.Text)
        ReplaceConfigSettings("UseCSV", chkUseCSV.Checked)
        ReplaceConfigSettings("CSVFileName", lblCSVFile.Text.Trim.Substring(lblCSVFile.Text.Trim.LastIndexOf("\") + 1))
        ReplaceConfigSettings("ReportFilePath", Me.lblInsightReport.Text.Trim)
        'Save ListBox of meeting types to skip
        Dim x As Integer
        Dim meetingType As String
        x = 0
        Do Until x = ListBox1.Items.Count
            meetingType = "MeetingSkipType" & (x + 1)
            ReplaceConfigSettings(meetingType, CType(ListBox1.Items.Item(x), String).ToString())
            x += 1
        Loop
        Do Until x = ListBox1.Items.Count
            meetingType = "MeetingSkipType" & (x + 1)
            ReplaceConfigSettings(meetingType, "")
            x += 1
        Loop
        cust.ID = txtCustID.Text.Trim
        cust.ClinicName = ConfigurationManager.AppSettings("ClinicName").Trim
        cust.DaysPrior = IIf(txtDaysPrior.Text.Trim.Length > 0, txtDaysPrior.Text.Trim, 0)
        cust.AreaCode = txtAreaCode.Text.Trim
        cust.UseCSV = chkUseCSV.Checked
        cust.ReportPath = lblInsightReport.Text.Trim
        cust.ArchivePath = GetCurrentDirectory() & "\Archive"
        cust.CSVPath = lblCSVFile.Text.Trim
        cust.Engine = ConfigurationManager.AppSettings("Engine").Trim
        cust.ErrorPath = GetCurrentDirectory() & "\Exception.txt"
        cust.CallLogic = ConfigurationManager.AppSettings("CallLogic").Trim
    End Sub
    Private Function NumberOfCallsWritten() As Integer
        Dim reportReader As StreamReader
        Dim callfile As String = GetCurrentDirectory & "\" & ConfigurationManager.AppSettings("CallListFileName")
        Try
            Dim line As String
            Dim writtenCounter As Integer = -1
            If File.Exists(callfile) Then
                'Count the number of lines in the Call List file
                reportReader = New StreamReader(callfile)
                line = reportReader.ReadLine
                Do While Not line Is Nothing
                    If line.Length > 0 Then
                        writtenCounter += 1
                    End If
                    line = reportReader.ReadLine
                Loop
                Return writtenCounter
            Else
                Return 0
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If reportReader IsNot Nothing Then
                reportReader.Close()
                reportReader.Dispose()
            End If
        End Try
    End Function

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        Dim max As Integer
        'Determine how many items can be added according to the config file
        max = CType(ConfigurationManager.AppSettings("MeetingSkipTypeTotal"), Integer)
        If ListBox1.Items.Count < max Then
            If Trim(txtNewType.Text) <> "" Then
                ListBox1.Items.Add(txtNewType.Text.ToUpper)
                lblTypeMsg.Text = ""
                txtNewType.Text = ""
            End If
        Else
            lblTypeMsg.Text = "The maximum of " & max & " types has been reached."
        End If
    End Sub
    Private Sub btnCallLogicHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmCallLogicHelp
        form = New frmCallLogicHelp
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub bntScheduledDateHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmScheduledDateHelp
        form = New frmScheduledDateHelp
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub btnMeetingTypesHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmMeetingTypes
        form = New frmMeetingTypes
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub btnProviderListHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim form As frmProviderListHelp
        form = New frmProviderListHelp
        form.Show()
        form.BringToFront()
    End Sub
    Private Sub ChooseCSV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChooseCSV.Click
        Dim fDialog As New OpenFileDialog

        fDialog.Title = "Get Provider List in CSV Format"
        '   fDialog.Filter = "CSV Files|*.csv"
        fDialog.InitialDirectory = Directory.GetCurrentDirectory

        If fDialog.ShowDialog() = DialogResult.OK Then
            lblCSVFile.Text = fDialog.FileName.ToString()
        End If
    End Sub
    Private Sub btnCreateFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateFile.Click
        Dim export As New DataExport
        Dim useCSV As Boolean
        Dim x As Integer
        cust = New Customer
        cursor.Current() = System.Windows.Forms.Cursors.WaitCursor
        lblMsg.ForeColor = System.Drawing.Color.Black
        lblMsg.Text = ""

        'The customer id must not be blank
        If Trim(txtCustID.Text) = "" Then
            lblMsg.ForeColor = System.Drawing.Color.Red
            lblMsg.Text = "Customer ID is missing."
            Exit Sub
        End If
       
        'If the Use CSV option is selected, a file must be chosen
        If chkUseCSV.Checked Then
            If Trim(lblCSVFile.Text) = "" Then
                lblMsg.ForeColor = System.Drawing.Color.Red
                lblMsg.Text = "Please select a CSV file."
                Exit Sub
            End If
            useCSV = True
        Else
            useCSV = False
        End If
    
        'Build array of values from ListBox
        x = 0
        Do Until x = ListBox1.Items.Count
            cust.meetingTypeArray.Add(ListBox1.Items.Item(x))
            x += 1
        Loop

        Try
            Save()
        Catch ex As Exception
            lblMsg.Text = ex.Message
            Exit Sub
        End Try
        output = New OutputResults
        output = export.Main(cust)

        lblMsg.Text = output.Msg

        If Not output.FatalError Then
            lblMsg.ForeColor = System.Drawing.Color.Black
            If output.CallsCount > 0 Then Shell("notepad " & output.CallListPath, AppWinStyle.NormalFocus)
            If output.ExceptionCount > 0 Then Shell("notepad " & cust.ErrorPath, AppWinStyle.NormalFocus)
        Else
            lblMsg.ForeColor = System.Drawing.Color.Red
        End If
    End Sub

    Private Sub btnFTP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTP.Click
        Try
            If Me.NumberOfCallsWritten <= 0 Then
                lblMsg.Text = "There are no rows in the call list file."
                lblMsg.ForeColor = Color.Red
                Exit Sub
            End If

            FTP()

            lblMsg.Text = "Call list successfully uploaded to server."
            lblMsg.ForeColor = Color.Blue
        Catch ex As Exception
            lblMsg.Text = ex.Message
            lblMsg.ForeColor = Color.Red
        End Try
    End Sub

    Private Sub ChooseInsightReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChooseInsightReport.Click
        Dim fDialog As New OpenFileDialog

        fDialog.Title = "Get Report in Tab Delimited Format"
        '     fDialog.Filter = "Tab Delimited Files|*.txt"
        ' fDialog.InitialDirectory = Directory.GetCurrentDirectory

        If fDialog.ShowDialog() = DialogResult.OK Then
            Me.lblInsightReport.Text = fDialog.FileName.ToString()
        End If
    End Sub

End Class
