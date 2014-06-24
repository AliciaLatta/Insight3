
Imports System.Configuration
Imports System.IO
Imports Rebex.Net
Imports Microsoft.VisualBasic.FileIO
Imports System.Collections.Generic

'=================================================================================================
'Class Name:	DataExport
'Description:	Holds functions used to parse a report and create a formatted data file.
'Property of Archipelago Systems, LLC.
'=================================================================================================
Public Class DataExport
#Region "Enums"
    Friend Enum ProcessingStatus
        SuccessfulRow
        ErroredRow
        SkippedRow
        WrittenRow
    End Enum
    Private Enum PhoneType
        Mobile
        Home
    End Enum
    Private Enum LogicType
        AllBut
        NoneBut
    End Enum
#End Region
#Region "Variables"
    'Declare variable used as the return string for main
    Dim results As String
    Private Const _className As String = "DataExport"
    Private Const skipReasonMeetingType As String = "Meeting Type"
    Private Const skipReasonPrivPhone As String = "PRIV Keyword"
    Private Const skipReasonNoProvider As String = "Missing Provider ID"
    Private Const skipReasonNoOK As String = "OK Keyword Not Found"
    Private Const invalidCallTime As String = "The call date/time cannot be in the past."
    Private Const skipReasonScheduledDaysPrior As String = "The appointment was created within the timeframe of the 'no reminder necessary' rule."
    Private Const csvProviderProblem As String = "There was a problem with the Provider List xls file.  Be sure it is in the correct directory (as specified in the xlsProviderListFile key of the config file).  Also verify that it is formatted and named properly.  Also, be sure it is not open."
    Private Const xlsReportProblem As String = "There was a problem with the Insight Report xls file.  Be sure it is in the correct directory (as specified in the xlsProviderListFile key of the config file).  Also verify that it is formatted and named properly.  Also, be sure it is not open."
    Private Const ProgramError As String = "There was a problem with the program.  Please try again.  If the problem persists, close then reopen the application."
    Private Const OpenReportFile As String = "The Insight report is open.  Please close it and try again."
    Private Const providerIDNotFoundCSV As String = "A provider id/meeting type combination matching a combination in the Insight OPEN report was not found in the CSV file.  Please review the CSV and Insight report files to be sure all provider ids are accurate."
    Private Const locIDNotFoundCSV As String = "A location id/meeting type combination matching a combination in the Insight OPEN report was not found in the CSV file.  Please review the CSV and Insight report files to be sure all provider ids are accurate."
    Private Const providerIDNotFoundConfig As String = "A provider id/meeting type combination matching a combination in the Insight OPEN report was not found in the config file.  Please review the config file and Insight report files to be sure all provider ids are accurate."
    Private Const problemReadingProviderXLS As String = "There is a problem with the columns in the Provider List CSV file.  Please check that they all exist."
    Private Const problemReadingReportXLS As String = "There is a problem with the columns in the Insight Report tab delimited file.  Please check that they all exist."
    Dim ReportFile, CallListFile As FileInfo
    Dim outputReader As StreamReader
    ' Dim outputWriter As StreamWriter
    Dim exceptions, callrows As ArrayList
    Shared meetingType As String
    Shared msg As String = String.Empty
    Shared locID As Integer = 1
    Private _error As String 'Member level variable used to hold the error message
    Dim gCallDate As String
    Shared callDate As Date
#End Region
    '=================================================================================================
    'Method Name:	DataExport.Main
    'Description:	Entry Point for the executable
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Public Function Main(ByVal custID As String, ByVal daysPrior As Integer, ByVal scheduledDaysPrior As String, _
            ByVal meetingTypeArray() As String, _
            ByVal useCSV As Boolean, ByVal CSVFile As String, ByVal ReportFilePath As String) As String
        Dim exceptionWriter As StreamWriter
        Dim callListFilePath, exceptionFilePath As String 'Strings to store filenames
        Try
            Try
                'Get current settings from the configuration file
                callListFilePath = ConfigurationManager.AppSettings("CallListFile").ToString
                exceptionFilePath = ConfigurationManager.AppSettings("ExceptionFile").ToString
            Catch e As Exception
                'Problem with the config file
                UpdateResults("There was a problem with the configuration file: " & e.ToString)
                Return results
            End Try
            ProcessTransactions(ReportFilePath, callListFilePath, exceptionFilePath, custID, daysPrior, scheduledDaysPrior, meetingTypeArray, useCSV, CSVFile)
            exceptionFilePath = ConfigurationManager.AppSettings("ExceptionFile").ToString()
            exceptionWriter = New StreamWriter(exceptionFilePath, False)
            For Each row As String In exceptions
                exceptionWriter.WriteLine(row)
            Next
            Return results

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
    Public Function ArchiveCallList() As Boolean
        Dim callListFilePath As String
        Dim outputArchivePath As String
        Try
            callListFilePath = ConfigurationManager.AppSettings("CallListFile").ToString
            outputArchivePath = ConfigurationManager.AppSettings("OutputArchive").ToString
        Catch ex As Exception
            Return False
        End Try
        Try
            'Create the directories if they're not there already
            Dim archive As String()
            archive = outputArchivePath.Split(".")
            archive(0) = archive(0) & "_" & Date.Now.Ticks
            archive(0) = archive(0) & ".txt"
            CheckForMissingDirectory(outputArchivePath)
            'Move the file to the archive
            File.Move(callListFilePath, archive(0).ToString())
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
    Private Sub UpdateResults(ByVal status As String)
        results = results & status
        CloseReaders()
    End Sub
    Private Sub UpdateResultsFinal(ByVal status As String)
        results = results & status
    End Sub
    Private Sub CloseReaders()
        Try
            'If Not exceptionWriter Is Nothing Then
            '    exceptionWriter.Close()
            '    exceptionWriter = Nothing
            'End If
            'If Not outputWriter Is Nothing Then
            '    outputWriter.Close()
            '    outputWriter = Nothing
            'End If
            If Not CallListFile Is Nothing Then
                CallListFile = Nothing
            End If
            If Not ReportFile Is Nothing Then
                ReportFile = Nothing
            End If
            If Not outputReader Is Nothing Then
                outputReader.Close()
                outputReader = Nothing
            End If
        Catch e As Exception
            'Problem with the config file
            UpdateResults(ProgramError & ": " & e.Message)
            If Not CallListFile Is Nothing Then
                CallListFile = Nothing
            End If
        End Try
    End Sub
   
    '=================================================================================================
    'Method Name:	DataExport.ProcessTransactions
    'Description:	Processes all transactions in the Input file
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Sub ProcessTransactions(ByVal reportFilePath As String, ByVal callListFilePath As String, _
                                     ByVal exceptionFilePath As String, ByVal custID As String, _
                                     ByVal daysPrior As Integer, ByVal scheduledDaysPrior As String, _
                                     ByVal meetingTypeArray() As String, ByVal useProviderCSV As Boolean, _
                                     ByVal CSVProviderFilePath As String)

        Dim skipCounter As Integer
        Dim processedCounter As Integer
        ' Dim line As String
        Dim splitout As Array
        ' Dim records As ArrayList
        Dim x As Integer
        Dim y As Integer
        Dim exists As Boolean
        Dim xlsReader As StreamReader
        Dim dt As DataTable 
        Try
            'Delete any lines in the output file left from the last run of the program      
            Try
                CallListFile = New FileInfo(callListFilePath)
                If CallListFile.Exists Then
                    CallListFile.Delete()
                End If
            Catch e As Exception
                'Problem with the config file
                _error = "Error deleting Call List from last run.  Make sure it's not open."
                Throw e
            End Try

            'Try
            'Create the directories if they're not there already
            CheckForMissingDirectory(exceptionFilePath)
            'Create an instace of StreamWriter to interact with the log files
            exceptions = New ArrayList()
            'Write out a header in the file to separate runs
            exceptions.Add(New String("*", 100))
            exceptions.Add("* Run Date: " & Date.Now.ToString("f"))
            exceptions.Add(New String("*", 100))
            exceptions.Add("")

            If File.Exists(reportFilePath) Then
                ReportFile = New FileInfo(reportFilePath)
                If ReportFile.LastWriteTime < Date.Now.AddHours(-8) Then
                    UpdateResults("The report was created more than 8 hours ago.  Please recreate it and try again.")
                    Exit Sub
                End If
                'Create the directories if they're not there already
                CheckForMissingDirectory(reportFilePath)
                CheckForMissingDirectory(callListFilePath)

                Try
                    'Get report
                    xlsReader = New StreamReader(reportFilePath)
                    dt = BuildDataTable(xlsReader)
                Catch e As Exception
                    UpdateResults(OpenReportFile & ": " & e.Message)
                    If xlsReader IsNot Nothing Then
                        xlsReader.Close()
                        xlsReader = Nothing
                    End If
                    Exit Sub

                End Try
                Dim rowCounter As Integer

                Try
                    CallListFile = New FileInfo(callListFilePath)
                    '   outputWriter = New StreamWriter(callListFilePath)
                Catch e As Exception
                    UpdateResults("Problem creating the Call List File: " & e.Message)
                    Exit Sub
                End Try

                For Each row As DataRow In dt.Rows
                    rowCounter = (rowCounter + 1)
                    If rowCounter > 1 Then
                        Select Case ProcessRow(row, callListFilePath, _
                           exceptionFilePath, custID, _
                           daysPrior, scheduledDaysPrior, _
                           meetingTypeArray, useProviderCSV, _
                           CSVProviderFilePath, rowCounter)
                            Case ProcessingStatus.SuccessfulRow
                                processedCounter += 1
                            Case ProcessingStatus.SkippedRow
                                skipCounter += 1
                            Case ProcessingStatus.ErroredRow
                                UpdateResults(ProcessError)
                                Exit Sub
                            Case ProcessingStatus.WrittenRow
                        End Select
                    End If
                Next

                ' outputWriter.Close()
                ' If Not outputReader Is Nothing Then outputReader.Dispose()

                'Count the number of lines in the output file
                '   outputReader = New StreamReader(callListFilePath)
                'Add the header lines to the records array
                'records = New ArrayList
                'records.Add(outputReader.ReadLine)
                'records.Add(outputReader.ReadLine)
                'records.Add(outputReader.ReadLine)
                'records.Add(outputReader.ReadLine)
                'line = outputReader.ReadLine
                Dim keepers As New ArrayList
                Dim allrows As New ArrayList
                Dim splitout2 As Array
                Dim dup As Boolean
                Dim time1, time2 As DateTime
                Dim phone1, phone2, name1, name2 As String
                x = 0
                For Each row As String In callrows
                    If x > 0 Then
                        'Count the number records in the call list
                        'Find any duplicate phone number - name combos and only print the first appointment of the day
                        splitout = Split(row, ",")
                        phone1 = splitout(0).ToString.Trim
                        name1 = splitout(5).ToString.Trim
                        time1 = Convert.ToDateTime(splitout(3).ToString.Trim)
                        y = 0
                        For Each row2 As String In callrows
                            exists = False
                            If y > 0 Then 'skip the header
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
                            End If
                            y += 1
                        Next
                        If Not dup Then allrows.Add(row)
                        dup = False
                    Else : x += 1
                    End If
                Next

                callrows = allrows

                Try
                    'x = 0
                    'outputReader.Close()
                    'outputReader = Nothing
                    'outputWriter = New StreamWriter(callListFilePath)
                    'Do Until x = (records.Count)
                    '    outputWriter.WriteLine(records(x))
                    '    x += 1
                    'Loop
                    'This is where we will open and write to the outputWriter
                    Dim callsWriter As New StreamWriter(callListFilePath)
                    callrows.Add("*EOF*")
                    For Each row As String In callrows
                        callsWriter.WriteLine(row)
                    Next
                    callsWriter.Close()
                    'If outputReader IsNot Nothing Then outputReader.Dispose()
                Catch ex As Exception
                    UpdateResults(ProgramError & ": " & ex.Message)
                    Exit Sub
                End Try

                exceptions.Add("")
                exceptions.Add("Rows Written: " & callrows.Count - 1)
                'Need to subtract the last line (*EOF*) from the count
                'Write out the processing results to the screen
                UpdateResultsFinal(New String("-", 100))
                UpdateResultsFinal(vbCrLf & "Run Date: " & Date.Now.ToString("s"))
                UpdateResultsFinal(vbCrLf & "Processing Complete" & vbCrLf)
                'If the row count is less than 0, set it to 0
                UpdateResultsFinal(vbCrLf & "Rows Written: " & callrows.Count - 1 & vbCrLf & vbCrLf)
                UpdateResultsFinal("Call list created in ")
                'Rename file
                Dim newName As String
                newName = "CallList-" & gCallDate & "-" & custID & ".csv"
                If My.Computer.FileSystem.FileExists(ConfigurationManager.AppSettings("DataFolderPath") & newName) Then
                    My.Computer.FileSystem.DeleteFile(ConfigurationManager.AppSettings("DataFolderPath") & newName)
                End If

                My.Computer.FileSystem.RenameFile(ConfigurationManager.AppSettings("CallListFile").ToString(), newName)

                UpdateResultsFinal(ConfigurationManager.AppSettings("CallListFile") & vbCrLf)
                UpdateResultsFinal(New String("-", 100))
            Else
                'Write a line to the log file to indicate that the input file did not exist
                exceptions.Add("Input File: " & reportFilePath & " does not exist.")
                ''Let user know that input file does not exist
                UpdateResults("Input File: " & reportFilePath & " does not exist.")
            End If
        Catch ex As Exception
            UpdateResults(_error & ": " & ex.Message)
        Finally
            CloseReaders()
            ' If Not outputWriter Is Nothing Then outputWriter.Close()
            If Not outputReader Is Nothing Then outputReader.Dispose()
        End Try
    End Sub

    Private Function ProcessRow(ByVal row As DataRow, ByVal callListFilePath As String, _
                                     ByVal exceptionFilePath As String, ByVal custID As String, _
                                     ByVal daysPrior As Integer, ByVal scheduledDaysPrior As String, _
                                     ByVal meetingTypeArray() As String, ByVal useProviderCSV As Boolean, _
                                     ByVal CSVProviderFilePath As String, ByVal rowCounter As Integer) As ProcessingStatus

        Dim InsightProviderID, IVRProviderID As String
        Dim apptDate As Date

        apptDate = Convert.ToDateTime(row(8).ToString.Trim)

        If row(10).ToString.Trim = "12:00n" Then
            apptDate = Convert.ToDateTime(apptDate & " " & "12:00pm")
        Else
            apptDate = Convert.ToDateTime(apptDate & " " & row(10).ToString.Trim)
        End If

        If rowCounter = 2 Then
            callDate = Convert.ToDateTime(apptDate).AddDays(-Convert.ToInt32(daysPrior))
            gCallDate = callDate.ToString("yyyyMMdd")

            If callDate < Today Then
                _error = invalidCallTime
                Return ProcessingStatus.ErroredRow
            End If
            callrows = New ArrayList()
            callrows.Add("Phone,LocationID,DocID,ApptDateTime,SpecialMessage,PatientName,SMS,Extra")
        End If

        meetingType = row(9).ToString.Trim

        If DetermineMeeting(meetingType, meetingTypeArray) Then Return ProcessingStatus.SkippedRow 'Meeting type should be skipped

        InsightProviderID = row(6).ToString.Trim

        Try
            If useProviderCSV = True Then
                If Not GetProviderID_FromCSV(IVRProviderID, InsightProviderID, CSVProviderFilePath) Then
                    Return ProcessingStatus.ErroredRow
                End If
            Else
                If Not GetProviderID_FromConfig(IVRProviderID, InsightProviderID) Then
                    Return ProcessingStatus.ErroredRow
                End If
            End If
        Catch ex As Exception
            If _error = csvProviderProblem Then Throw ex
            Return ProcessingStatus.ErroredRow
        End Try

        If IVRProviderID Is Nothing Then Return ProcessingStatus.ErroredRow

        Try
            If Trim(ConfigurationManager.AppSettings("CallLogic").ToString.ToUpper) = "ALLBUT" Then
                WriteRecord(row, IVRProviderID, daysPrior, apptDate, LogicType.AllBut)
            Else
                WriteRecord(row, IVRProviderID, daysPrior, apptDate, LogicType.NoneBut)
            End If
        Catch ex As Exception
            Throw ex
            Return ProcessingStatus.ErroredRow
        End Try

        Return ProcessingStatus.SuccessfulRow
    End Function

    Public Class CustomRow
        Public Property data As String
        Public Property duplicate As Boolean

        Public Sub New()

        End Sub

        Public Sub New(ByVal data As String, ByVal duplicate As Boolean)
            Me.data = data
            Me.duplicate = duplicate
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
                    .Number = ConfigurationManager.AppSettings("DefaultAreaCode").ToString & .Number
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
            msg = "NO CALL will be made to " & name & " as a valid number does not exist."
            AddException(msg)
            Return retval
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
                msg = "NO CALL will be made to " & name & " as a valid number does not exist."
                AddException(msg)
                Return retval
            End If
        End If
        Dim msg2 As String

        If sms Then 'Want a text
            msg2 = "Voice call will be made to " & String.Format("{0:(###) ###-####}", Long.Parse(retval.Number)) & "(" & retval.Type & ") instead."
            For Each phone As Phone In phones
                With phone
                    If .Type.StartsWith("MOBILE") Then
                        msg = "UNABLE to TEXT " & name & " as mobile number is invalid:" & String.Format("{0:(###) ###-####}", Long.Parse(.Number)) & ".  " & msg2
                        Return retval
                    End If
                End With
            Next
            'If we get here, mobile number doesn't exist
            msg = "UNABLE to TEXT " & name & " as mobile number is missing.  " & msg2
            AddException(msg)
            Return retval
        Else 'They want to call home 
            msg2 = "Will use " & String.Format("{0:(###) ###-####}", Long.Parse(retval.Number)) & "(" & retval.Type & ") instead."
            For Each phone As Phone In phones
                With phone
                    If .Type.StartsWith("HOME") Then
                        msg = "UNABLE to CALL " & .Type & " of " & name & " as " & .Type & " number is invalid:" & String.Format("{0:(###) ###-####}", Long.Parse(.Number)) & ".  " & msg2
                        AddException(msg)
                        Return retval
                    End If
                End With
            Next
            'If we get here, home number doesn't exist
            msg = "UNABLE to CALL HOME of " & name & ";  " & msg2
            AddException(msg)
            Return retval
        End If
        'Return retval
    End Function
    Private Sub AddException(ByVal msg As String)
        exceptions.Add(msg)
        exceptions.Add("------------------------------------------------------------------------------------------------------------------------------")
    End Sub
    Private Sub WriteRecord(ByVal row As DataRow, ByVal providerID As String, ByVal daysPrior As Integer, _
                        ByVal apptDate As Date, ByVal logic As LogicType)
        Dim privateIndicator, okayIndicator, spanishIndicator As Boolean
        Dim sms As Integer = 0
        Dim fullName As String
        Dim phone As Phone
        Dim type As String = PhoneType.Home
        Dim phoneList As New List(Of Phone)

        'Look for 'PR' in the Guar Class1, Guar Class2,	Pat Class1,	Pat Class2, and	Pat Class3 columns
        privateIndicator = (row(16).ToString.Trim.ToUpper = "PR" Or row(17).ToString.Trim.ToUpper = "PR" Or _
            row(18).ToString.Trim.ToUpper = "PR" Or row(19).ToString.Trim.ToUpper = "PR" Or _
            row(20).ToString.Trim.ToUpper = "PR")

        okayIndicator = (row(16).ToString.Trim.ToUpper = "OK" Or row(17).ToString.Trim.ToUpper = "OK" Or _
            row(18).ToString.Trim.ToUpper = "OK" Or row(19).ToString.Trim.ToUpper = "OK" Or _
            row(20).ToString.Trim.ToUpper = "OK")

        If (logic = LogicType.AllBut And Not privateIndicator) Or _
                (logic = LogicType.NoneBut And okayIndicator) Then
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
                spanishIndicator = row(21).ToString.ToUpper.Trim.StartsWith("SP")
                If spanishIndicator Then sms = 2
                If spanishIndicator And phone.Type = "MOBILE" Then sms = 3

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
                callrows.Add(srow)
            End If
        End if
    End Sub
    '=================================================================================================
    'Method Name:	Trans.DetermineMeeting
    'Description:	Returns boolean value of whether the record's meeting type should be skipped
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Function DetermineMeeting(ByVal meetingType As String, ByVal meetingTypeArray() As String) As Boolean
        'Get the meeting type: If the type is listed in the config file as a type to skip, don't print the line
        Dim y = 0
        Dim skipMeeting As Boolean
        Do Until y = (CType(ConfigurationManager.AppSettings("MeetingSkipTypeTotal"), Integer) + 1)
            If Trim(meetingTypeArray(y)) <> "" Then
                If meetingTypeArray(y).ToUpper = meetingType.ToUpper Then
                    skipMeeting = True
                    y = (CType(ConfigurationManager.AppSettings("MeetingSkipTypeTotal"), Integer))
                End If
            End If
            y += 1
        Loop
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
  
    Private Function GetProviderID_FromConfig(ByRef IVRProviderID As String, ByVal InsightProviderID As String) As Boolean
        Dim z As Integer = 1
        Dim done As Boolean = False
        Do Until done = True Or z = CType(ConfigurationManager.AppSettings("EngineProviderTotal"), Integer)
            IVRProviderID = "EngineProvider" & z
            If ConfigurationManager.AppSettings(IVRProviderID).ToString().ToUpper = InsightProviderID.ToUpper Then
                IVRProviderID = z
                done = True
            Else
                z += 1
            End If
        Loop
        If done Then
            Return True
        Else
            _error = DataExport.providerIDNotFoundConfig
        End If
        Return False
    End Function

    Private Function GetProviderID_FromCSV(ByRef IVRProviderID As String, ByVal InsightProviderID As String, ByVal providerCSVFile As String) As Boolean
        '*****************************************************************************************************
        '   Find the corresponding IVR provider id based on Insight OPEN provider 
        '   id and meeting type by parsing the xls file
        '   Return True if a match was found; False if no match was found
        '*****************************************************************************************************
        Dim xlsProviderListFile As String
        Dim xlsReader As StreamReader
        Dim xlsLine As String
        Dim splitOutProviderList As Array
        Dim lookupInsightOpenProviderID As String
        Dim lookupMeetingType As String

        Try
            Try
                xlsProviderListFile = providerCSVFile
                xlsReader = New StreamReader(xlsProviderListFile)
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

            xlsReader = New StreamReader(xlsProviderListFile)
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
        Dim directoryEndPosition As Integer

        directoryEndPosition = filePath.LastIndexOf("\")

        'If the necessary directory is not found, it is created
        If Not Directory.Exists(filePath.Substring(0, directoryEndPosition + 1)) Then
            Directory.CreateDirectory(filePath.Substring(0, directoryEndPosition + 1))
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