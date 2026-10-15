Imports System
Imports System.Globalization
Imports Microsoft.SqlServer.Dts.Runtime


Public Class Generate_BK_ID_Recon

    Public Shared _brs5Generator As BRS5.Generator.GenerateBRS5 = Nothing
    Public Shared _databaseControl As BMSL.DatabaseControl = Nothing
    Public Shared _logger As BMSL.LogControl = Nothing
    Public Shared _Noteslogger As BMSL.LogControl = Nothing

    Public Shared _packageFolder As String = String.Empty


    Public Shared inputArgname As String = ""

    Public Shared inrptstartDate As String ' convert to rptStartDate as DateTime before Generate_BK_ID_Report
    Public Shared inrptendDate As String   ' convert to rptEndDate   as DateTime before GEnerate_BK_ID_Report

    Public Shared rptStartDate As DateTime
    Public Shared rptEndDate As DateTime

    Public Const Status_Complete = 10
    Public Const Status_Error = 99


    Shared Sub Main()
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)
        _Noteslogger = New BMSL.LogControl(LocalConstants.NOTES_PATH)
        _databaseControl = New BMSL.DatabaseControl(_logger, LocalConstants.SQL_CONNECTION_STRING)
        _brs5Generator = New BRS5.Generator.GenerateBRS5(LocalConstants.LOG_PATH, LocalConstants.SQL_CONNECTION_STRING)

        _logger.LogInfo("")

        Dim ArgName As String = ParseCommandLineArgs()

        Select Case Len(ArgName)

            Case 0
                ' no argument supplied

                rptstartDate = "01/01/1900"

                Dim FirstDayCurrentMonth As DateTime = GetFirstDayOfMonth(Today)
                Dim LastDayCurrentMonth As DateTime = GetLastDayOfMonth(Today)

                'rptendDate = Format(Now().AddDays(1), "yyyyMMdd").ToString

                Dim DDToday As Integer = Day(Today)

                rptEndDate = Today.AddDays(1)

                DisplayRunInfo()
                _logger.LogInfo("Report Inclusive Start and End dates for this run are: " & rptstartDate & " and " & rptendDate)
                _Noteslogger.LogInfo("Report Inclusive Start and End dates for this run are: " & rptStartDate & " and " & rptEndDate)

                'continue

            Case 17
                inrptstartDate = Left(ArgName, 8)
                inrptendDate = Right(ArgName, 8)

                If Not (validyyyymmddDate(inrptstartDate) And validyyyymmddDate(inrptendDate)) Then
                    _logger.LogError("Invalid Start and/or End Date Parameter format ")
                    DisplayRunInfo()
                    Dim ERRORCODE As String = 99
                    Dim ErrorDescription As String = "Invalid CommandLine Format.  Input 'Generate_BK_ID_RECON_Rpt.exe' or 'Generate_BK_ID_RECON_Rpt.exe /ReportDates=yyyymmddyyyymmdd'"
                    BK_ID_Recon_Error_Log(ERRORCODE, ErrorDescription)
                    _logger.LogError("Please try again")
                    ' do not continue
                    Exit Sub
                Else
                    ' convert validated string date to datetime 
                    rptStartDate = Date.ParseExact(inrptstartDate, "yyyyMMdd", Nothing)
                    rptEndDate = Date.ParseExact(inrptendDate, "yyyyMMdd", Nothing)
                    ' continue 
                End If

                DisplayRunInfo()
                _logger.LogInfo("Report Inclusive Start and End dates for this run are: " & rptStartDate & " and " & rptEndDate)
                _Noteslogger.LogInfo("Report Inclusive Start and End dates for this run are: " & rptStartDate & " and " & rptEndDate)


            Case Else
                ' wrong argument length
                _logger.LogError("Invalid Start and End Date Parameter length (and/or format)")
                DisplayRunInfo()
                Dim ERRORCODE As String = 99
                Dim ErrorDescription As String = "Invalid CommandLine Format.  Input 'Generate_BK_ID_RECON_Rpt.exe' or 'Generate_BK_ID_RECON_Rpt.exe /ReportDates=yyyymmddyyyymmdd'"
                BK_ID_Recon_Error_Log(ERRORCODE, ErrorDescription)
                _logger.LogError("Please try again")
                ' do not continue
                Exit Sub

        End Select

        Try
            _logger.LogInfo("")
            _logger.LogInfo("BK_ID_Recon_Report Generation starting up")

            Dim Generation_Status As Integer = Status_Error

            Generation_Status = Generate_BK_ID_Report(rptStartDate, rptEndDate)

            _logger.LogInfo("BK_ID_Recon_Report Generation terminating")

        Catch ex As Exception
            _logger.LogError("Exception was unrecoverable: " & ex.Message & " " & ex.StackTrace)
            _Noteslogger.LogError("Exception was unrecoverable: " & ex.Message & " " & ex.StackTrace)
        End Try


    End Sub
    Private Shared Function validyyyymmddDate(ByVal dateString As String) As Boolean
        Dim enUS As New CultureInfo("en-US")
        Dim dateValue As Date

        If Date.TryParseExact(dateString, "yyyyMMdd", enUS, _
                              DateTimeStyles.None, dateValue) Then
            validyyyymmddDate = True
        Else
            validyyyymmddDate = False
        End If

    End Function
    Private Shared Function ParseCommandLineArgs() As String

        Dim inputArgument As String = "/reportdates="

        For Each s As String In My.Application.CommandLineArgs
            If s.ToLower.StartsWith(inputArgument) Then
                inputArgname = s.Remove(0, inputArgument.Length)
            End If
        Next
        Return inputArgname.ToUpper

    End Function

    Private Shared Sub DisplayRunInfo()
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)

        _logger.LogInfo("")
        _logger.LogInfo("Command line format is:")
        _logger.LogInfo("")
        _logger.LogInfo("Option 1: 'Generate_BK_ID_RECON_Rpt.exe /reportdates=yyyymmdd-yyyymmdd'")
        _logger.LogInfo("          Where first yyyymmdd = Start Date")
        _logger.LogInfo("               second yyyymmdd = End Date")
        _logger.LogInfo("")
        _logger.LogInfo("You can also run without a parameter as shown below")
        _logger.LogInfo("")
        _logger.LogInfo("Option2: 'Generate_BK_ID_RECON_Rpt.exe'")
        _logger.LogInfo("         IN this case, Start Date = '01/01/1900'")
        _logger.LogInfo("         and End Date             = tomorrow's date")
        _logger.LogInfo("")

    End Sub


    Public Shared Function Generate_BK_ID_Report(ByVal rptStartDate As String, ByVal rptEndDate As String) As Integer
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)
        _Noteslogger = New BMSL.LogControl(LocalConstants.NOTES_PATH)
        _databaseControl = New BMSL.DatabaseControl(_logger, LocalConstants.SQL_CONNECTION_STRING)

        Try
            ' Run Stored Procedure for Auditing Unique BK_ID Counts

            Dim CustomerCode As String = "PRA"

            _logger.LogInfo("Report Start Date: " & rptStartDate)
            _logger.LogInfo("Report End Date: " & rptEndDate)
            _logger.LogInfo("Customer Code: " & CustomerCode)

            Dim AuditStatus As Integer = Status_Error
            AuditStatus = AuditBK_ID_Counts(rptStartDate, rptEndDate, CustomerCode)

            If AuditStatus = Status_Complete Then
                ' proceed with Report Generation via SSIS
                _logger.LogInfo("Audit of BK_ID Counts Passed.  Will start generating report")

                Dim outputSSIS As New BRS5.Generator.OutputSSISFiles(_databaseControl, _logger, LocalConstants.PACKAGE_FOLDER, LocalConstants.OUTPUT_FOLDER, LocalConstants.DELIVERY_FOLDER, LocalConstants.ON_HOLD_FOLDER)

                Dim Package_Filename As String = LocalConstants.PACKAGE_NAME

                Dim fdate As String = Format(Now(), "yyyyMMdd").ToString
                Dim fExt As String = ".txt"
                Dim fStatic As String = "LCI_TO_PRA_Reconciliation_"
                Dim BK_ID_Recon_Filename As String = LocalConstants.OUTPUT_FOLDER & fStatic & fdate & fExt

                Dim SSIS_Status As Integer = Status_Error

                'SSIS_Status = outputSSIS.GenerateBK_ID_Recon(rptStartDate, rptEndDate, CustomerCode, BK_ID_Recon_Filename, Package_Filename)
                SSIS_Status = lclGenerateBK_ID_Recon(rptStartDate, rptEndDate, CustomerCode, BK_ID_Recon_Filename, Package_Filename)


                If SSIS_Status = Status_Complete Then
                    _logger.LogInfo("Successfully Generated file [" & BK_ID_Recon_Filename & "]")
                    _Noteslogger.LogInfo("Successfully Generated file [" & BK_ID_Recon_Filename & "]")

                    ' on Jul 22, 2013, Tom requested via email to have the output file dropped to the two folder locations below.
                    ' Automation will post the file afterwards.

                    ' Copy to Delivery_Folder
                    Dim Delivery_Filename As String = LocalConstants.DELIVERY_FOLDER & fStatic & fdate & fExt
                    CopyFile(BK_ID_Recon_Filename, Delivery_Filename)
                    _logger.LogInfo("Copied to [" & Delivery_Filename & "]")
                    _Noteslogger.LogInfo("Copied to [" & Delivery_Filename & "]")


                    ' Copy to On_Hold_Folder
                    Dim On_Hold_Filename As String = LocalConstants.ON_HOLD_FOLDER & fStatic & fdate & fExt
                    CopyFile(BK_ID_Recon_Filename, On_Hold_Filename)
                    _logger.LogInfo("Copied to [" & On_Hold_Filename & "]")
                    _Noteslogger.LogInfo("Copied to [" & On_Hold_Filename & "]")

                    Return Status_Complete
                Else
                    ' Failed in Generating a report
                    _logger.LogInfo("SSIS Failed in Generating file [" & BK_ID_Recon_Filename & "]")
                    _logger.LogInfo("Inserting Error Log in BK_ID_Recon_Error_Log Table")

                    _Noteslogger.LogInfo("SSIS Failed in Generating file [" & BK_ID_Recon_Filename & "]")
                    _Noteslogger.LogInfo("Inserting Error Log in BK_ID_Recon_Error_Log Table")

                    Dim ErrorCode As String = Status_Error
                    Dim ErrorDescription As String = "SSIS Failed in Generating file [" & BK_ID_Recon_Filename & "]"
                    BK_ID_Recon_Error_Log(ErrorCode, ErrorDescription)

                    Return Status_Error
                End If

            Else
                ' send email notification
                _logger.LogInfo("Audit of BK_ID Counts Failed.  Cannot generate report")
                _logger.LogInfo("Inserting Error Log in BK_ID_Recon_Error_Log Table")

                _Noteslogger.LogInfo("Audit of BK_ID Counts Failed.  Cannot generate report")
                _Noteslogger.LogInfo("Inserting Error Log in BK_ID_Recon_Error_Log Table")

                Dim ErrorCode As String = Status_Error
                Dim ErrorDescription As String = "Audit of BK_ID Counts Failed.  Cannot generate report"
                BK_ID_Recon_Error_Log(ErrorCode, ErrorDescription)

                Return Status_Error

            End If
        Catch ex As Exception

            _logger.LogInfo("**  EXCEPTION **:  " & ex.Message & ex.StackTrace)
            _logger.LogInfo("Inserting Error Log in BK_ID_Recon_Error_Log Table")

            _Noteslogger.LogInfo("**  EXCEPTION **:  " & ex.Message & ex.StackTrace)
            _Noteslogger.LogInfo("Inserting Error Log in BK_ID_Recon_Error_Log Table")

            Dim ErrorCode As String = Status_Error
            Dim ErrorDescription As String = "Exception Error.  Cannot generate report"
            BK_ID_Recon_Error_Log(ErrorCode, ErrorDescription)

            Return Status_Error
        End Try


    End Function

    Private Shared Function AuditBK_ID_Counts(ByVal rptStartDate As DateTime, ByVal rptEndDate As DateTime, ByVal CustomerCode As String) As Integer
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)

        Dim AuditStatus As Integer = 0
        Try

            _logger.LogInfo("Audit of BK_ID Counts Started")
            Dim sqlString1 As String = "[dbo].[RecordCountAudit_Delivery_Log_Primary_Case_VS_PRIMARY_CASE_History]"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            sqlData1.AddParameter("@rptStartDate", SqlDbType.DateTime, 10, rptStartDate)
            sqlData1.AddParameter("@rptEndDate", SqlDbType.DateTime, 10, rptEndDate)
            sqlData1.AddParameter("@CustomerCode", SqlDbType.VarChar, 10, CustomerCode)
            sqlData1.AddParameter("@AuditStatus", SqlDbType.Int, 2, 0)
            sqlData1.SetOutputParameter("@AuditStatus")

            AuditStatus = sqlData1.ExecuteNonQuery("@AuditStatus")

            Return AuditStatus


        Catch ex As Exception
            _logger.LogInfo("**  EXCEPTION **:  " & ex.Message & ex.StackTrace)
            Throw ex
        End Try



    End Function

    Private Shared Sub BK_ID_Recon_Error_Log(ByVal ErrorCode As String, ByVal ErrorDescription As String)
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)

        Try
            _logger.LogInfo("Writing Error Information into the Error_Log Table")
            Dim sqlstring As String = "BK_ID_Recon_Error_Log__Insert"
            Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlstring)

            sqlData.AddParameter("@ErrorCode", SqlDbType.VarChar, 2, ErrorCode)
            sqlData.AddParameter("@ErrorDescription", SqlDbType.VarChar, 250, ErrorDescription)

            Dim rowsAffected As Integer = sqlData.ExecuteNonQuery
            _logger.LogInfo(rowsAffected & " row(s) affected")

        Catch ex As Exception
            _logger.LogInfo("**  EXCEPTION **:  " & ex.Message & ex.StackTrace)
            Throw ex
        End Try


    End Sub
    Public Shared Function GetFirstDayOfMonth(ByVal dtDate As DateTime) As DateTime

        Dim dtFrom As DateTime = dtDate
        dtFrom = dtFrom.AddDays(-(dtFrom.Day - 1))
        Return dtFrom

    End Function

    Public Shared Function GetLastDayOfMonth(ByVal dtDate As DateTime) As DateTime

        Dim dtTo As New DateTime(dtDate.Year, dtDate.Month, 1)
        dtTo = dtTo.AddMonths(1)
        dtTo = dtTo.AddDays(-(dtTo.Day))
        Return dtTo

    End Function

    Public Shared Function lclGenerateBK_ID_Recon(ByVal rptStartDate As DateTime, ByVal rptEndDate As DateTime, ByVal Customercode As String, ByVal BK_ID_Recon_Filename As String, ByVal Package_Name As String) As Integer

        'check if file already exists
        If System.IO.File.Exists(BK_ID_Recon_Filename) Then
            resolveFileConflict(BK_ID_Recon_Filename)
        End If


        'Check if SSIS package file exists
        If VerifyFileExists(Package_Name) = False Then
            _logger.LogInfo("SSIS Package [" & Package_Name & "] was not found.")
            ' Return STATUS_ERROR
        Else
            '  SSIS package was found
        End If

        'run SSIS package
        _logger.LogInfo("Running SSIS package: " & _packageFolder & Package_Name)

        Return WriteSSISBK_ID_RECON(rptStartDate, rptEndDate, Customercode, BK_ID_Recon_Filename, Package_Name)


    End Function
    Private Shared Function WriteSSISBK_ID_RECON(ByVal rptStartDate As DateTime, ByVal rptEndDate As DateTime, ByVal CustomerCode As String, ByVal BK_ID_Recon_Filename As String, ByVal package_Name As String) As Integer

        Dim eventListener As New SQLServerEventListener()

        'Dim pkgLocation As String
        Dim pkg As New Package
        Dim app As New Application
        Dim pkgResults As DTSExecResult
        'Dim ex As Exception = Nothing

        Try

            'pkgLocation = IS_Package_Location
            'app.PackagePassword = "5p0ng3b0b"
            pkg = app.LoadPackage(package_Name, eventListener)
            pkg.Variables("FileName").Value = BK_ID_Recon_Filename
            pkg.Variables("rptStartDate").Value = rptStartDate
            pkg.Variables("rptEndDate").Value = rptEndDate
            pkg.Variables("CustomerCode").Value = CustomerCode



            pkgResults = pkg.Execute(Nothing, Nothing, eventListener, Nothing, Nothing)

            _logger.LogInfo("SSIS Package Execution Results: " & pkgResults.ToString())
            If (pkgResults = DTSExecResult.Failure) Then
                _logger.LogInfo("SSIS Package Execution Errors: " & pkgResults.ToString())
                Return Status_Error

            Else
                Return Status_Complete
            End If

        Catch ex As Exception
            _logger.LogInfo("WriteSSISBK_ID_Recon generated an exception: " & ex.Message & ex.StackTrace)

            'Throw ex
            Return Status_Error
        End Try

    End Function
    Public Shared Sub resolveFileConflict(ByVal filename As String)

        _logger.LogInfo("Resolving duplicate file conflict: " & filename)

        System.IO.File.Move(filename, filename & ".renamed." & Format(Now(), "yyyyMMddHHmmss"))

    End Sub
    Private Shared Function VerifyFileExists(ByVal filename As String) As Boolean

        If System.IO.File.Exists(filename) Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Sub CopyFile(ByVal sourceFilename As String, ByVal destFilename As String)

        If System.IO.File.Exists(sourceFilename) Then
            _logger.LogInfo("Copying file " & sourceFilename & " to " & destFilename)

            If System.IO.File.Exists(destFilename) Then
                'dest file already exists - rename
                resolveFileConflict(destFilename)
            Else
                'dest file doesn't exist - OK
            End If

            'copy file 
            System.IO.File.Copy(sourceFilename, destFilename)
        Else
            'file doesn't exist
            Throw New Exception("Could not file copy file [" & sourceFilename & "] - File doesn't exist")

        End If

    End Sub


End Class



