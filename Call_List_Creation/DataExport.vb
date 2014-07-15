
Imports System.Configuration
Imports System.IO
Imports Microsoft.VisualBasic
Imports Rebex.Net
Imports Microsoft.VisualBasic.FileIO
Imports System.Collections.Generic
Imports Microsoft.VisualBasic.FileIO.TextFieldParser
Imports System.Runtime.CompilerServices
'=================================================================================================
'Class Name:	DataExport
'Description:	Holds functions used to parse a report and create a formatted data file.
'Property of Archipelago Systems, LLC.
'=================================================================================================
Public Class DataExport
#Region "Enums"
    Public Enum ProcessingStatus
        SuccessfulRow
        ErroredRow
        SkippedRow
        WrittenRow
        InProcessRow
    End Enum
   
    Private Enum LogicType
        AllBut
        NoneBut
    End Enum
#End Region
#Region "Variables"
    'Declare variable used as the return string for main
    Private Const _className As String = "DataExport"
    Private Const skipReasonMeetingType As String = "Meeting Type"
    Private Const skipReasonPrivPhone As String = "PRIV Keyword"
    Private Const skipReasonNoProvider As String = "Missing Provider ID"
    Private Const skipReasonNoOK As String = "OK Keyword Not Found"
    Private Const invalidCallTime As String = "The call date/time cannot be in the past."
    Private Const skipReasonScheduledDaysPrior As String = "The appointment was created within the timeframe of the 'no reminder necessary' rule."
    Private Const csvProviderProblem As String = "There was a problem with the Provider List xls file.  Be sure it is in the correct directory (as specified in the xlsProviderListFile key of the config file).  Also verify that it is formatted and named properly.  Also, be sure it is not open."
    Private Const xlsReportProblem As String = "There was a problem with the Insight Report xls file.  Be sure it is in the correct directory (as specified in the xlsProviderListFile key of the config file).  Also verify that it is formatted and named properly.  Also, be sure it is not open."
    Private Const OpenReportFile As String = "The Insight report is open.  Please close it and try again."
    Private Const providerIDNotFoundCSV As String = "A provider id/meeting type combination matching a combination in the Insight OPEN report was not found in the CSV file.  Please review the CSV and Insight report files to be sure all provider ids are accurate."
    Private Const locIDNotFoundCSV As String = "A location id/meeting type combination matching a combination in the Insight OPEN report was not found in the CSV file.  Please review the CSV and Insight report files to be sure all provider ids are accurate."
    Private Const providerIDNotFoundConfig As String = "A provider id/meeting type combination matching a combination in the Insight OPEN report was not found in the config file.  Please review the config file and Insight report files to be sure all provider ids are accurate."
    Private Const problemReadingProviderXLS As String = "There is a problem with the columns in the Provider List CSV file.  Please check that they all exist."
    Private Const problemReadingReportXLS As String = "There is a problem with the columns in the Insight Report tab delimited file.  Please check that they all exist."
    Dim ReportFile As FileInfo
    Dim outputReader As StreamReader
    'Dim exceptions As New List(Of ProcessedRow)
    Dim exceptionLog As New ArrayList
    Dim callRows As New ArrayList()
    Shared meetingType As String
    Shared msg As String = String.Empty
    Shared locID As Integer = 1
    Private _error As String 'Member level variable used to hold the error message
    Dim gCallDate As String
    Shared callDate As Date
    Dim output As OutputResults
    Dim cust As Customer
#End Region
   
    Public Function Main(ByVal icust As Customer) As OutputResults
        Dim exceptionWriter As StreamWriter
        cust = icust
        output = New OutputResults()
        Try
            ProcessTransactions(cust)
            exceptionWriter = New StreamWriter(cust.ErrorPath, False)
           
            For Each row As String In exceptionLog
                exceptionWriter.WriteLine(row)
            Next
            output.ExceptionCount = exceptionLog.Count
            Return output
        Catch ex As Exception
        Finally
            CloseReaders()
            If exceptionWriter IsNot Nothing Then
                exceptionWriter.Close()
                exceptionWriter.Dispose()
            End If
        End Try
    End Function
    '=================================================================================================
    'Method Name:	DataExport.ArchiveCallList
    'Description:	Writes and executes the FTP script
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Public Function ArchiveCallList(ByVal cust As Customer) As Boolean
        Try
            'Create the directories if they're not there already
            Dim archive As String()
            archive = cust.ArchivePath.Split(".")
            archive(0) = archive(0) & "_" & Date.Now.Ticks
            archive(0) = archive(0) & ".txt"
            CheckForMissingDirectory(cust.ArchivePath)
            'Move the file to the archive
            File.Move(output.CallListPath, archive(0).ToString())
        Catch ex As Exception
            Throw ex
        End Try
        Return True
    End Function
    '=================================================================================================
    'Method Name:	DataExport.UpdateResults
    'Description:	Updates a string that holds the results of the execution of the integration file creation code
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Sub UpdateResults(ByVal status As String, ByVal fatal As Boolean)
        output.FatalError = fatal
        output.Msg = output.Msg & status
    End Sub
   
    Private Sub CloseReaders()
        Try
            If Not ReportFile Is Nothing Then
                ReportFile = Nothing
            End If
            If Not outputReader Is Nothing Then
                outputReader.Close()
                outputReader = Nothing
            End If
        Catch e As Exception
            'Problem with the config file
            UpdateResults("Exception: " & e.Message, True)
        End Try
    End Sub

    '=================================================================================================
    'Method Name:	DataExport.ProcessTransactions
    'Description:	Processes all transactions in the Input file
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Sub ProcessTransactions(ByVal cust As Customer)
        'Dim skipCounter As Integer
        'Dim processedCounter As Integer
        Dim splitout As Array
        Dim x, y As Integer
        Dim exists As Boolean
        Dim xlsReader As StreamReader
        Dim dt As DataTable
        Try
            exceptionLog.Add(" " & vbCrLf & New String("*", 100) & vbCrLf & "* Run Date: " & Date.Now.ToString("f") & vbCrLf & New String("*", 100) & vbCrLf & " ")

            If File.Exists(cust.ReportPath) Then
                ReportFile = New FileInfo(cust.ReportPath)
                If ReportFile.LastWriteTime < Date.Now.AddHours(-8) Then
                    UpdateResults("The report was created more than 8 hours ago.  Please recreate it and try again.", True)
                    Exit Sub
                End If
                'Create the directories if they're not there already
                CheckForMissingDirectory(cust.ReportPath)
                Try
                    'Get report
                    xlsReader = New StreamReader(cust.ReportPath)
                    dt = BuildDataTable(xlsReader)
                Catch e As Exception
                    UpdateResults(OpenReportFile & ": " & e.Message, True)
                    If xlsReader IsNot Nothing Then
                        xlsReader.Close()
                        xlsReader = Nothing
                    End If
                    Exit Sub
                End Try
                Dim pr1 As ProcessedRow
                For Each row As DataRow In dt.Rows
                    cust.RowCount = (cust.RowCount + 1)
                    If cust.RowCount > 1 Then
                        pr1 = ProcessRow(row, cust)
                        Select Case pr1.Type
                            Case ProcessingStatus.SuccessfulRow
                                cust.ProcessedCount += 1
                            Case ProcessingStatus.SkippedRow
                                If pr1.Msg <> String.Empty Then exceptionLog.Add(pr1.Msg)
                                cust.SkippedCount += 1
                            Case ProcessingStatus.ErroredRow
                                UpdateResults(pr1.Msg, False)
                                output.FatalError = True
                                Exit Sub
                            Case ProcessingStatus.WrittenRow
                        End Select
                    End If
                Next

                Dim keepers As New ArrayList
                Dim allrows As New ArrayList
                Dim splitout2 As Array
                Dim dup As Boolean
                Dim time1, time2 As DateTime
                Dim phone1, phone2, name1, name2 As String

                allrows.Add("Phone,LocationID,DocID,ApptDateTime,SpecialMessage,PatientName,SMS,Extra")
                For Each row As String In callRows
                    'Find any duplicate phone number - name combos and only print the first appointment of the day
                    splitout = Split(row, ",")
                    phone1 = splitout(0).ToString.Trim
                    name1 = splitout(5).ToString.Trim
                    time1 = Convert.ToDateTime(splitout(3).ToString.Trim)
                    y = 0
                    For Each row2 As String In callRows
                        exists = False

                        splitout2 = Split(row2, ",")
                        phone2 = splitout2(0).ToString.Trim
                        name2 = splitout2(5).ToString.Trim
                        time2 = Convert.ToDateTime(splitout2(3).ToString.Trim)
                        Try
                            If phone1 = phone2 And name1 = name2 Then
                                If time1 > time2 Then 'Only leave it if it's the earlier time
                                    dup = True
                                    Exit For
                                End If
                            End If
                        Catch ex As Exception
                            'Handle argument out-of-range exception
                        End Try

                    Next
                    If Not dup Then allrows.Add(row)
                    dup = False
                Next

                callRows = allrows
                Dim callList As String
                Dim callsWriter As StreamWriter
                Try
                    'This is where we will open and write the call list
                    callList = GetCurrentDirectory() & "\" & "CallList" & cust.ClinicName & "_" & gCallDate & "-" & cust.ID & ".csv"
                    ReplaceConfigSettings("CallListFile", callList)
                    If My.Computer.FileSystem.FileExists(callList) Then
                        My.Computer.FileSystem.DeleteFile(callList)
                    End If
                    output.CallListPath = callList
                    callsWriter = New StreamWriter(callList)
                    callRows.Add("*EOF*")
                    For Each row As String In callRows
                        callsWriter.WriteLine(row)
                    Next
                    callsWriter.Close()
                    'If outputReader IsNot Nothing Then outputReader.Dispose()
                Catch ex As Exception
                    UpdateResults("Exception: " & ex.Message, True)
                    If callsWriter IsNot Nothing Then callsWriter.Close()
                    Exit Sub
                End Try
                output.CallsCount = callRows.Count - 2
                exceptionLog.Add(" ")
                exceptionLog.Add("Rows Written: " & output.CallsCount)
                'Need to subtract the last line (*EOF*) from the count
                'Write out the processing results to the screen
                UpdateResults(New String("-", 100), False)
                UpdateResults(vbCrLf & "Run Date: " & Date.Now.ToString("s"), False)
                UpdateResults(vbCrLf & "Processing Complete" & vbCrLf, False)
                'If the row count is less than 0, set it to 0
                UpdateResults(vbCrLf & "Rows Written: " & output.CallsCount & vbCrLf & vbCrLf, False)
                UpdateResults("Call list created in ", False)

                UpdateResults(callList & vbCrLf, False)
                UpdateResults(New String("-", 100), False)
            Else
                'Write a line to the log file to indicate that the input file did not exist
                exceptionLog.Add("Input File: " & cust.ReportPath & " does not exist.")
                ''Let user know that input file does not exist
                UpdateResults("Input File: " & cust.ReportPath & " does not exist.", True)
            End If
        Catch ex As Exception
            UpdateResults(_error & ": " & ex.Message, True)
        Finally
            CloseReaders()
            ' If Not outputWriter Is Nothing Then outputWriter.Close()
            If Not outputReader Is Nothing Then outputReader.Dispose()
        End Try
    End Sub

    Private Sub ReplaceConfigSettings(ByVal key As String, ByVal val As String)
        Dim xDoc As New Xml.XmlDataDocument
        Dim xNode As Xml.XmlNode
        Dim path As String() = Directory.GetFiles(GetCurrentDirectory, "*.config")
        xDoc.Load(path(0)) 
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
    Public Shared Function GetCurrentDirectory() As String
        Dim retval As String
        retval = Directory.GetCurrentDirectory
        If retval.Contains("bin") Then
            retval = retval.Substring(0, retval.Length - 4)
        End If
        Return retval
    End Function
    Private Function ProcessRow(ByVal row As DataRow, ByVal cust As Customer) As ProcessedRow
        Dim InsightProviderID, IVRProviderID As String
        Dim apptDate As Date
        Dim retval As New ProcessedRow()
        apptDate = Convert.ToDateTime(row(8).ToString.Trim)

        If row(10).ToString.Trim = "12:00n" Then
            apptDate = Convert.ToDateTime(apptDate & " " & "12:00pm")
        Else
            apptDate = Convert.ToDateTime(apptDate & " " & row(10).ToString.Trim)
        End If

        If cust.RowCount = 2 Then
            callDate = Convert.ToDateTime(apptDate).AddDays(-Convert.ToInt32(cust.DaysPrior))
            gCallDate = callDate.ToString("yyyyMMdd")

            If callDate < Today Then
                retval.Msg = invalidCallTime
                retval.Type = ProcessingStatus.ErroredRow
                Return retval
            End If
        End If

        meetingType = row(9).ToString.Trim

        If DetermineMeeting(meetingType, cust.meetingTypeArray) Then
            retval.Msg = String.Empty
            retval.Type = ProcessingStatus.SkippedRow
            Return retval 'Meeting type should be skipped
        End If

        InsightProviderID = row(6).ToString.Trim

        Try
            If cust.UseCSV Then
                If Not GetProviderID_FromCSV(IVRProviderID, InsightProviderID, cust) Then
                    retval.Msg = String.Empty
                    retval.Type = ProcessingStatus.SkippedRow
                    Return retval
                End If
            Else
                Dim providerIDs As Array
                Dim reader As New StreamReader("..\ProviderList.txt")
                Dim line As String

                Do Until reader.EndOfStream
                    line = reader.ReadLine
                    providerIDs = Split(line, ",")
                    If providerIDs(1).ToString().Contains(InsightProviderID) Then
                        IVRProviderID = providerIDs(0).ToString()
                        Exit Do
                    End If
                Loop
            End If
        Catch ex As Exception
            If _error = csvProviderProblem Then
                Throw ex
            End If
        End Try

        If IVRProviderID Is Nothing Then
            '  _error = csvProviderProblem
            retval.Msg = String.Empty ' don't print as some doctors don't want calls made
            retval.Type = ProcessingStatus.SkippedRow
            Return retval
        End If

        Try
            If cust.CallLogic = LogicType.AllBut Then
                WriteRecord(row, IVRProviderID, apptDate, cust)
            Else
                WriteRecord(row, IVRProviderID, apptDate, cust)
            End If
        Catch ex As Exception
            Throw ex
            ' Return ProcessingStatus.ErroredRow
        End Try

        retval.Msg = String.Empty
        retval.Type = ProcessingStatus.SuccessfulRow
        Return retval
    End Function

    Public Class OutputResults
        Public Property Msg As String
        Public Property ExceptionCount As Integer
        Public Property CallsCount As Integer
        Public Property FatalError As Boolean
        Public Property CallListPath As String
    End Class
    Public Class Customer
        Public Property ID As String
        Public Property ClinicName As String
        Public Property ErrorPath As String
        Public Property ReportPath As String
        Public Property CSVPath As String
        '  Public Property DataFolderPath As String
        Public Property Engine As String
        Public Property AreaCode As String
        Public Property DaysPrior As Integer
        Public Property ArchivePath As String
        Public Property UseCSV As Boolean
        Public Property RowCount As Integer
        Public Property ProcessedCount As Integer
        Public Property SkippedCount As Integer
        Public Property meetingTypeArray As ArrayList

        Private _CallLogic As String
        Public Property CallLogic() As String
            Get
                Return _CallLogic
            End Get
            Set(ByVal value As String)
                If value.ToUpper = "ALLBUT" Then
                    _CallLogic = LogicType.AllBut
                Else
                    _CallLogic = LogicType.NoneBut
                End If
            End Set
        End Property

        Public Sub New()
            meetingTypeArray = New ArrayList
        End Sub
    End Class
  
    Public Class Phone
        Public Property Type As String
        Public Property Number As String
        Public Property Status As String
        Public Property SMS As Boolean
        Public Sub New()
            Me.Status = "UNKNOWN"
        End Sub
        Public Sub New(ByVal Type As String, ByVal Number As String, ByVal Status As String)
            Me.Type = Type
            Me.Number = Number
            Me.Status = Status
        End Sub
    End Class

    Public Class ProcessedRow
        Public Property Type As ProcessingStatus
        Public Property Msg As String
        Public Sub New()

        End Sub
        Public Sub New(ByVal msg As String, ByVal type As ProcessingStatus)
            Me.Msg = msg
            Me.Type = type
        End Sub
    End Class

    '=================================================================================================
    'Method Name:	Trans.FormatPhones
    'Description:	Takes all spaces out the phone and adds area code if relevant
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Function GetPhone(ByVal name As String, ByRef phones As List(Of Phone), ByVal sms As Boolean) As Phone
        Dim x As Integer
        Dim cur As String
        Dim strNewPhone As String = String.Empty
        Dim retval As New Phone
        'Clean up the data and set the status
        For Each phone As Phone In phones
            With phone
                Do Until x = phone.Number.Length
                    cur = phone.Number.Chars(x)
                    Select Case cur
                        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                            'If this is the first time in the loop, take the zero out
                            If x = 0 Then
                                strNewPhone = ""
                            End If
                            'The character is a number so add it to the new string
                            strNewPhone += cur
                    End Select
                    x += 1
                Loop
                If strNewPhone.Length = 10 Then
                    .Status = "OK"
                    .Number = strNewPhone
                End If
                If strNewPhone.Length = 7 Then
                    .Number = cust.AreaCode & .Number
                    .Status = "OK"
                    .Number = strNewPhone
                End If
                strNewPhone = String.Empty
                x = 0
            End With
        Next

        Dim msg As String = String.Empty
        'To Do -- Loop through numbers and write out exception messages for text if selected , home if text and mobile is missing or invalid, and home if voice 
        'If we get here, there is an issue with the first choice of home or mobile

        For Each phone As Phone In phones
            If sms Then
                If phone.Type = "MOBILE" And phone.Status = "OK" Then
                    Return phone
                ElseIf sms And phone.Type = "HOME" And phone.Status = "OK" Then
                    retval = phone
                ElseIf sms AndAlso retval.Type <> "HOME" And phone.Status = "OK" Then
                    retval = phone
                End If
            Else
                If phone.Type = "HOME" And phone.Status = "OK" Then
                    Return phone
                ElseIf phone.Type = "MOBILE" And phone.Status = "OK" Then
                    retval = phone
                ElseIf retval.Type <> "MOBILE" Then
                    retval = phone
                End If
            End If
        Next

        If retval Is Nothing Then
            exceptionLog.Add("NO CALL will be made to " & name & " as a valid number does not exist.")
        Else
            Dim ok As Boolean = False
            For Each phone As Phone In phones
                With phone
                    If .Status = "OK" Then
                        ok = True
                    End If
                End With
            Next
            If Not ok Then 
                exceptionLog.Add("NO CALL will be made to " & name & " as a valid number does not exist.")
                Return retval
            End If
        End If
        Dim msg2 As String
       
        If sms Then 'Want a text
            msg2 = "Voice call will be made to " & String.Format("{0:(###) ###-####}", Long.Parse(retval.Number)) & " (" & retval.Type & ") instead."
            For Each phone As Phone In phones
                With phone
                    If .Type = "MOBILE" Then
                        msg = "UNABLE to TEXT " & name & " as mobile number is invalid:" & String.Format("{0:(###) ###-####}", Long.Parse(.Number)) & ".  " & msg2
                        Return retval
                    End If
                End With
            Next
            'If we get here, mobile number doesn't exist
            exceptionLog.Add("UNABLE to TEXT " & name & " as mobile number is missing.  " & msg2)
            Return retval
        Else 'They want to call home 
            msg2 = "Will use " & String.Format("{0:(###) ###-####}", Long.Parse(retval.Number)) & " (" & retval.Type & ") instead."
            For Each phone As Phone In phones
                With phone
                    If .Type = "HOME" Then
                        exceptionLog.Add("UNABLE to CALL " & .Type & " of " & name & " as " & .Type & " number is invalid:" & String.Format("{0:(###) ###-####}", Long.Parse(.Number)) & ".  " & msg2)
                        Return retval
                    End If
                End With
            Next
            'If we get here, home number doesn't exist
            exceptionLog.Add("UNABLE to CALL HOME of " & name & ";  " & msg2)
            Return retval
        End If
        'Return retval
    End Function
    Private Sub AddException(ByVal msg As String) 
        exceptionLog.Add(msg)
        exceptionLog.Add("------------------------------------------------------------------------------------------------------------------------------")
    End Sub
    Private Sub WriteRecord(ByVal row As DataRow, ByVal providerID As String, ByVal apptDate As Date, ByVal cust As Customer)
        Dim privateIndicator, okayIndicator, spanishIndicator As Boolean
        Dim sms As Integer = 0
        Dim fullName As String
        Dim phone As Phone
        Dim type As String = "HOME"
        Dim phoneList As New List(Of Phone)

        'Look for 'PR' in the Guar Class1, Guar Class2,	Pat Class1,	Pat Class2, and	Pat Class3 columns
        privateIndicator = (row(16).ToString.Trim.ToUpper = "PR" Or row(17).ToString.Trim.ToUpper = "PR" Or _
            row(18).ToString.Trim.ToUpper = "PR" Or row(19).ToString.Trim.ToUpper = "PR" Or _
            row(20).ToString.Trim.ToUpper = "PR")

        okayIndicator = (row(16).ToString.Trim.ToUpper = "OK" Or row(17).ToString.Trim.ToUpper = "OK" Or _
            row(18).ToString.Trim.ToUpper = "OK" Or row(19).ToString.Trim.ToUpper = "OK" Or _
            row(20).ToString.Trim.ToUpper = "OK")

        If (cust.CallLogic = LogicType.AllBut And Not privateIndicator) Or _
                (cust.CallLogic = LogicType.NoneBut And okayIndicator) Then
            fullName = row(0).ToString.Trim & " " & row(1).ToString.Trim
            Dim home, work, mobile As String
            home = row(3).ToString.Trim()
            work = row(4).ToString.Trim()
            mobile = row(5).ToString.Trim()

            'If TX in any patient fields, SMS = 1 and replace phone with mobile
            If (row(18).ToString.ToUpper.Trim = "TX" Or row(19).ToString.ToUpper.Trim = "TX" Or row(20).ToString.ToUpper.Trim = "TX") Then
                sms = 1
            End If

            If home.Length > 0 Then phoneList.Add(New Phone() With {.Type = "HOME", .Number = home})
            If work.Length > 0 Then phoneList.Add(New Phone() With {.Type = "WORK", .Number = work})
            If mobile.Length > 0 Then phoneList.Add(New Phone() With {.Type = "MOBILE", .Number = mobile})

            phone = GetPhone(fullName, phoneList, (sms = 1))

            If phone.Status = "OK" And Not providerID Is Nothing And providerID.Length > 0 And Left(providerID, 14).ToUpper <> "ENGINEPROVIDER" And apptDate <> Nothing Then

                '**************************************************************************************
                ' English voice sms = 0
                ' English text sms = 1
                ' Spanish voice sms = 2
                ' Spanish text sms = 3
                '**************************************************************************************
                spanishIndicator = row(21).ToString.ToUpper.Trim.StartsWith("SP")
                If sms = 1 Then
                    If phone.Type = "MOBILE" Then
                    Else : sms = 0
                    End If
                End If
                If spanishIndicator And sms = 1 Then sms = 3
                If spanishIndicator And sms = 0 Then sms = 2
                 
                'Phone, LocationID, DocID, ApptDateTime, SpecialMessage, PatientName, SMS, Extra
                Dim srow As String

                srow = phone.Number
                srow += ", " & locID
                srow += ", " & providerID
                srow += ", " & apptDate.ToString("MM/dd/yyyy H:mm")
                srow += ", " & msg
                srow += ", " & row(0).ToString.Trim() 'First Name
                srow += ", " & sms
                srow += ", " & row(2).ToString.Trim() 'Account# / Extra
                srow += ""
                callRows.Add(srow)
            End If
        End If
    End Sub

    Private Function DetermineMeeting(ByVal meetingType As String, ByVal meetingTypeArray As ArrayList) As Boolean
        'Get the meeting type: If the type is listed in the config file as a type to skip, don't print the line
        Dim skipMeeting As Boolean
        For Each mtg As String In meetingTypeArray
            If mtg.Length > 0 Then
                If mtg = meetingType.ToUpper Then
                    skipMeeting = True
                End If
            End If
        Next
        Return skipMeeting
    End Function

    Private Function CheckForNullCharacters(ByVal s As String) As Boolean
        Dim reader As StringReader
        Dim n As Integer
        Try
            reader = New StringReader(s)

            'Loop through each character looking for a null character
            Do
                'Get the character code for the reader
                n = reader.Read

                'If its null, return true, exit do
                If n = 0 Then
                    Return True
                    Exit Do

                    '-1 means the end of the string had been reached
                ElseIf n = -1 Then
                    Return False
                    Exit Do
                End If
            Loop

        Finally
            'Clean up reader
            reader = Nothing
        End Try
    End Function

    'Private Function GetProviderID_FromConfig(ByRef IVRProviderID As String, ByVal InsightProviderID As String) As Boolean
    '    Dim z As Integer = 1
    '    Dim done As Boolean = False
    '    Do Until done = True Or z = CType(ConfigurationManager.AppSettings("EngineProviderTotal"), Integer)
    '        IVRProviderID = "EngineProvider" & z
    '        If ConfigurationManager.AppSettings(IVRProviderID).ToString().ToUpper = InsightProviderID.ToUpper Then
    '            IVRProviderID = z
    '            done = True
    '        Else
    '            z += 1
    '        End If
    '    Loop
    '    If done Then
    '        Return True
    '    Else
    '        _error = DataExport.providerIDNotFoundConfig
    '    End If
    '    Return False
    'End Function

    Private Function GetProviderID_FromCSV(ByRef IVRProviderID As String, ByVal InsightProviderID As String, ByVal cust As Customer) As Boolean
        '*****************************************************************************************************
        '   Find the corresponding IVR provider id based on Insight OPEN provider 
        '   id and meeting type by parsing the xls file
        '   Return True if a match was found; False if no match was found
        '*****************************************************************************************************
        Dim xlsReader As StreamReader
        Dim xlsLine, lookupInsightOpenProviderID, lookupMeetingType As String
        Dim splitOutProviderList As Array
        Try
            Try
                xlsReader = New StreamReader(cust.CSVPath)
                xlsLine = xlsReader.ReadLine
                xlsLine = xlsReader.ReadLine
                splitOutProviderList = Split(xlsLine, ",")
                lookupInsightOpenProviderID = Trim(splitOutProviderList(4)).ToUpper()
                lookupMeetingType = Trim(splitOutProviderList(3)).ToUpper.ToUpper
            Catch ex As Exception
                _error = csvProviderProblem
                Throw ex
            End Try
            Try
                Do While Not xlsLine Is Nothing
                    If xlsLine.Length > 0 Then
                        splitOutProviderList = Split(xlsLine, ",")
                        lookupInsightOpenProviderID = Trim(splitOutProviderList(4)).ToUpper()
                        lookupMeetingType = Trim(splitOutProviderList(3)).ToUpper
                        If InsightProviderID = lookupInsightOpenProviderID And meetingType.ToUpper = lookupMeetingType.ToUpper Then
                            IVRProviderID = Trim(splitOutProviderList(1))
                            msg = Trim(splitOutProviderList(5))
                            locID = Trim(splitOutProviderList(2))
                            Exit Do
                        End If
                    End If
                    xlsLine = xlsReader.ReadLine
                Loop
            Catch e As Exception
                _error = problemReadingProviderXLS
                Throw e
            End Try

            xlsReader = New StreamReader(cust.CSVPath)
            xlsLine = xlsReader.ReadLine
            xlsLine = xlsReader.ReadLine
            If IVRProviderID = Nothing Then
                'loop through the excel spreadsheet again to see if an All-Else exists
                Try
                    Do While Not xlsLine Is Nothing
                        If xlsLine.Length > 0 Then
                            splitOutProviderList = Split(xlsLine, ",")
                            lookupInsightOpenProviderID = Trim(splitOutProviderList(4)).ToUpper()
                            lookupMeetingType = Trim(splitOutProviderList(3)).ToUpper
                            If InsightProviderID = lookupInsightOpenProviderID And lookupMeetingType.ToUpper = "ALL-ELSE" Then
                                IVRProviderID = Trim(splitOutProviderList(1))
                                msg = Trim(splitOutProviderList(5))
                                locID = Trim(splitOutProviderList(2))
                                Exit Do
                            End If
                        End If
                        xlsLine = xlsReader.ReadLine
                    Loop
                Catch e As Exception
                    _error = problemReadingProviderXLS
                    Throw e
                End Try
                If IVRProviderID = Nothing Then
                    _error = providerIDNotFoundCSV
                    Return False
                End If
                If locID = Nothing Then
                    _error = locIDNotFoundCSV
                    Return False
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not xlsReader Is Nothing Then
                xlsReader.Close()
                xlsReader = Nothing
            End If
        End Try
        Return (IVRProviderID <> Nothing)
    End Function
    '=================================================================================================
    'Method Name:	DataExport.CheckForMissingDirectory
    'Description:	Verifies that the directory exists and creates it if necessary
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Sub CheckForMissingDirectory(ByVal filePath As String)

        If filePath IsNot Nothing Then
            Dim directoryEndPosition As Integer
            directoryEndPosition = filePath.LastIndexOf("\")

            'If the necessary directory is not found, it is created
            If Not Directory.Exists(filePath.Substring(0, directoryEndPosition + 1)) Then
                Directory.CreateDirectory(filePath.Substring(0, directoryEndPosition + 1))
            End If
        End If

    End Sub


    Private Function BuildDataTable(ByVal myReader As StreamReader) As DataTable
        Dim myTable As DataTable = New DataTable("MyTable")
        Dim i As Integer
        Dim myRow As DataRow
        Dim fieldValues As String()

        Try

            fieldValues = myReader.ReadLine().Split(vbTab)
            'Create data columns accordingly
            For i = 0 To fieldValues.Length() - 1
                myTable.Columns.Add(New DataColumn("Field" & i))
            Next
            'Adding the first line of data to data table
            myRow = myTable.NewRow
            For i = 0 To fieldValues.Length() - 1
                myRow.Item(i) = fieldValues(i).ToString
            Next
            myTable.Rows.Add(myRow)
            'Now reading the rest of the data to data table
            While myReader.Peek() <> -1
                fieldValues = myReader.ReadLine().Split(vbTab)
                myRow = myTable.NewRow
                For i = 0 To fieldValues.Length() - 1
                    myRow.Item(i) = fieldValues(i).ToString
                Next
                myTable.Rows.Add(myRow)
            End While
        Catch ex As Exception
            MsgBox("Error building datatable: " & ex.Message)
            Return New DataTable("Empty")
        Finally
            myReader.Close()
        End Try

        Return myTable
    End Function

    Public ReadOnly Property ProcessError() As String
        Get
            Return _error
        End Get
    End Property
End Class
