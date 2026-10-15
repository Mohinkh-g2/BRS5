Public Class DailyBatch

    Dim _logger As BMSL.LogControl = Nothing
    Dim _databaseControl As BMSL.DatabaseControl = Nothing

    Private Const BATCH_TYPE_DAILY_GENERATION As Integer = 1
    Private Const BATCH_TYPE_AD_HOC As Integer = 2
    Private Const BATCH_TYPE_CLIENT_OUTPUT_NO_BVM As Integer = 3
    Private Const BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM As Integer = 4

    Public Sub New()

        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)
        _databaseControl = New BMSL.DatabaseControl(_logger, LocalConstants.SQL_CONNECTION_STRING)

    End Sub

    'Public Sub Execute_Orig()

    '    _logger.LogInfo("Executing daily batch")

    '    'check if Cervello processing day has changed
    '    If WaitForNewProcessingDay() = False Then
    '        'processing day change timeout
    '        'ERROR
    '        _logger.LogError("Error waiting for new Cervello processing day")

    '    Else
    '        'processing day changed

    '        'hold any currently-ready items
    '        UpdateBatchControl_HoldReady()

    '        Dim parentBatchControlID As Integer = 0

    '        '#1 insert a new batch row for common tasks
    '        parentBatchControlID = InsertBatchControl_CervelloProcessingDateRange(BATCH_TYPE_DAILY_GENERATION)

    '        'start external process
    '        'wait for it to finish
    '        RunBRS5ExecuteDaily()

    '        '#2 insert a new batch row for specific tasks 1
    '        InsertBatchControl(BATCH_TYPE_CLIENT_OUTPUT_NO_BVM, parentBatchControlID)

    '        'start external process
    '        'wait for it to finish
    '        RunBRS5ExecuteDaily()

    '        'wait for accounts to be available
    '        If WaitForBISAccountLoad() = False Then
    '            'BIS account log timeout
    '            'ERROR
    '            _logger.LogError("Error waiting for BIS Accounts to load")

    '        Else
    '            '#3 insert a new batch row for specific tasks 2
    '            InsertBatchControl(BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM, parentBatchControlID)

    '            'start external process
    '            'wait for it to finish
    '            RunBRS5ExecuteDaily()

    '        End If


    '    End If


    'End Sub


    'Deepak Altura-39
    Public Sub Execute()

        _logger.LogInfo("Executing daily batch")
        RunBRS5ExecuteDaily()

        RunBRS5ExecuteDaily()

        'wait for accounts to be available
        If WaitForBISAccountLoad() = False Then
            'BIS account log timeout
            'ERROR
            _logger.LogError("Error waiting for BIS Accounts to load")

        Else
            '#3 insert a new batch row for specific tasks 2
            ' InsertBatchControl(BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM, parentBatchControlID)

            'start external process
            'wait for it to finish
            RunBRS5ExecuteDaily()
        End If
    End Sub
    Private Function WaitForBISAccountLoad() As Boolean

        'Deepak Altura-39
        _logger.LogInfo("Waiting for BIS Accounts to load")

        Dim bisAccountsLoaded As Boolean = False
        Dim startedChecking As Date = Now()

        bisAccountsLoaded = CheckIfBISAccountsLoaded()

        If bisAccountsLoaded = False Then
            Do
                _logger.LogInfo("Waiting for 10 minutes")
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)

                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)

                bisAccountsLoaded = CheckIfBISAccountsLoaded()

                'check if 16 hours have gone by
                If DateDiff(DateInterval.Hour, startedChecking, Now()) > 16 Then
                    _logger.LogInfo("Timed out waiting for BIS Accounts to load")
                    Return False
                End If

            Loop Until bisAccountsLoaded = True
        End If

        Return bisAccountsLoaded

    End Function

    Private Function WaitForNewProcessingDay() As Boolean

        _logger.LogInfo("Waiting for new Cervello processing day")

        Dim timingNewDate As Boolean = False
        Dim startedChecking As Date = Now()

        timingNewDate = CheckIfNewProcessingDay(0, 0, BATCH_TYPE_DAILY_GENERATION)

        If timingNewDate = False Then
            Do
                _logger.LogInfo("Waiting for 10 minutes")
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)

                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)
                Threading.Thread.Sleep(60000)

                timingNewDate = CheckIfNewProcessingDay(0, 0, BATCH_TYPE_DAILY_GENERATION)

                'check if 16 hours have gone by
                If DateDiff(DateInterval.Hour, startedChecking, Now()) > 16 Then
                    _logger.LogInfo("Timed out waiting for new Cervello processing day")
                    Return False
                End If

            Loop Until timingNewDate = True
        End If

        Return timingNewDate

    End Function

    Private Function CheckIfBISAccountsLoaded() As Boolean

        _logger.LogInfo("Checking if BIS Accounts have loaded")

        Dim sqlString2 As String = "app.usp_BIS_Process_Control__Max_Row"
        Dim sqlData2 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString2)
      
        Dim dsDataset2 As DataSet = sqlData2.SelectDataset

        Dim bisProcessingDate As String = String.Empty
        Dim creditVerificationCompletedDate As Date = Nothing
        
        If Not dsDataset2.Tables(0).Rows(0).IsNull("BIS_Processing_Date") Then bisProcessingDate = dsDataset2.Tables(0).Rows(0).Item("BIS_Processing_Date")
        If Not dsDataset2.Tables(0).Rows(0).IsNull("CreditVerification_Completed") Then creditVerificationCompletedDate = dsDataset2.Tables(0).Rows(0).Item("CreditVerification_Completed")

        _logger.LogInfo("BIS Processing Date is [" & bisProcessingDate & "]")

        If creditVerificationCompletedDate = Nothing Then
            'credit verification date is null
            _logger.LogInfo("Credit Verification last run day is null")
            _logger.LogInfo("BIS Accounts have not been loaded yet.")

            Return False
        Else
            'credit verification date is not null 
            'check the year/month/day 
            Dim today_Year As Integer = DatePart(DateInterval.Year, Now())
            Dim today_Month As Integer = DatePart(DateInterval.Month, Now())
            Dim today_DayOfMonth As Integer = DatePart(DateInterval.Day, Now())

            Dim creditVerificationCompleted_Year As Integer = DatePart(DateInterval.Year, creditVerificationCompletedDate)
            Dim creditVerificationCompleted_Month As Integer = DatePart(DateInterval.Month, creditVerificationCompletedDate)
            Dim creditVerificationCompleted_DayOfMonth As Integer = DatePart(DateInterval.Day, creditVerificationCompletedDate)

            If creditVerificationCompleted_Year = today_Year And creditVerificationCompleted_Month = today_Month And creditVerificationCompleted_DayOfMonth = today_DayOfMonth Then
                'credit verification date is today.
                'this was completed today
                _logger.LogInfo("Credit Verification last run day [" & creditVerificationCompletedDate & "] is today's date")
                _logger.LogInfo("BIS Accounts have been loaded.")
                Return True

            Else
                'credit vertification date is not today
                'this was completed on a different day.
                _logger.LogInfo("Credit Verification last run day [" & creditVerificationCompletedDate & "] is not today's date")
                _logger.LogInfo("BIS Accounts have not been loaded yet.")

                Return False
            End If
            
        End If

    End Function

    Private Function CheckIfNewProcessingDay(ByVal clientID As Integer, ByVal fileID As Integer, ByVal batchType As Integer) As Boolean
        _logger.LogInfo("Checking Cervello processing day timing")

        'Cervello data
        Dim sqlString1 As String = "app.usp_Cervello_Processing_Days__Select_Max"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        Dim dsDataset1 As DataSet = sqlData1.SelectDataset

        'BRS5 data
        Dim sqlString2 As String = "app.usp_Batch_Control__Select_Max_Completed"
        Dim sqlData2 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString2)
        sqlData2.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
        sqlData2.AddParameter("@File_ID", SqlDbType.SmallInt, 2, fileID)
        sqlData2.AddParameter("@Batch_Type", SqlDbType.SmallInt, 2, batchType)

        Dim dsDataset2 As DataSet = sqlData2.SelectDataset

        Dim cervelloProcessingStartDateTime As Date = Nothing
        Dim cervelloProcessingEndDateTime As Date = Nothing

        Dim batchRangeStartDT As Date = Nothing
        Dim batchRangeEndDT As Date = Nothing

        If Not dsDataset1.Tables(0).Rows(0).IsNull("Cervello_Processing_Start_Date_Time") Then cervelloProcessingStartDateTime = dsDataset1.Tables(0).Rows(0).Item("Cervello_Processing_Start_Date_Time")
        If Not dsDataset1.Tables(0).Rows(0).IsNull("Cervello_Processing_End_Date_Time") Then cervelloProcessingEndDateTime = dsDataset1.Tables(0).Rows(0).Item("Cervello_Processing_End_Date_Time")

        If Not dsDataset2.Tables(0).Rows(0).IsNull("Batch_Range_Start_DT") Then batchRangeStartDT = dsDataset2.Tables(0).Rows(0).Item("Batch_Range_Start_DT")
        If Not dsDataset2.Tables(0).Rows(0).IsNull("Batch_Range_End_DT") Then batchRangeEndDT = dsDataset2.Tables(0).Rows(0).Item("Batch_Range_End_DT")

        If cervelloProcessingStartDateTime > batchRangeStartDT Then
            'cervello has advanced
            _logger.LogInfo("Cervello starting date [" & cervelloProcessingStartDateTime & "] has changed from BRS5 [" & batchRangeStartDT & "]")

            If cervelloProcessingStartDateTime = batchRangeEndDT Then
                'time periods are contiguous
                Return True
            Else
                'time periods are not contiguous
                Throw New Exception("Cervello Processing Start Date Time [" & cervelloProcessingStartDateTime & "] is greater than the Batch Range End DT [" & batchRangeEndDT & "].  A time period may have been missed.")

            End If
        Else
            'cervello has not advanced since the last check
            _logger.LogInfo("Cervello and BRS5 starting dates are the same [" & cervelloProcessingStartDateTime & "]")
            Return False
        End If

    End Function

    Public Function InsertBatchControl_CervelloProcessingDateRange(ByVal batchType As Integer) As Integer

        Dim batchControlID As Integer = 0

        Try

            'batch control ID output

            _logger.LogInfo("Inserting batch control row - Daily Cervello Processing Date Range")
            _logger.LogInfo("Batch Type [" & batchType & "]")

            'Insert processing date and get new Batch_Control_ID
            Dim sqlString1 As String = "brs5.usp_Batch_Control__Insert_Cervello_Processing_Date_Range"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, 0)
            sqlData1.AddParameter("@File_ID", SqlDbType.Int, 4, 0)
            sqlData1.AddParameter("@Batch_Type", SqlDbType.Int, 4, batchType)
            sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            sqlData1.SetOutputParameter("@Batch_Control_ID")

            batchControlID = sqlData1.ExecuteNonQuery("@Batch_Control_ID")

            _logger.LogInfo("New Batch Control ID created: [" & batchControlID & "]")

            Return batchControlID

        Catch ex As Exception
            _logger.LogInfo("** EXCEPTION **: " & ex.Message & ex.StackTrace)

            Throw ex

        End Try

    End Function

    Public Function InsertBatchControl(ByVal batchType As Integer, parentBatchControlID As Integer) As Integer

        Dim batchControlID As Integer = 0

        Try

            'batch control ID output

            _logger.LogInfo("Inserting batch control row")
            _logger.LogInfo("Batch Type [" & batchType & "]")

            'Insert processing date and get new Batch_Control_ID
            Dim sqlString1 As String = "brs5.usp_Batch_Control__Insert"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, 0)
            sqlData1.AddParameter("@File_ID", SqlDbType.Int, 4, 0)
            sqlData1.AddParameter("@Batch_Type", SqlDbType.Int, 4, batchType)
            sqlData1.AddParameter("@Parent_Batch_Control_ID", SqlDbType.Int, 4, parentBatchControlID)
            sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            sqlData1.SetOutputParameter("@Batch_Control_ID")

            batchControlID = sqlData1.ExecuteNonQuery("@Batch_Control_ID")

            _logger.LogInfo("New Batch Control ID created: [" & batchControlID & "]")

            Return batchControlID

        Catch ex As Exception
            _logger.LogInfo("** EXCEPTION **: " & ex.Message & ex.StackTrace)

            Throw ex

        End Try

    End Function

    Public Sub UpdateBatchControl_HoldReady()

        Try

            _logger.LogInfo("Updating batch control row - hold current ready items")

            Dim sqlString1 As String = "brs5.usp_Batch_Control__Hold_Ready"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            Dim rowsAffected As Integer = 0
            rowsAffected = sqlData1.ExecuteNonQuery()

        Catch ex As Exception
            _logger.LogInfo("** EXCEPTION **: " & ex.Message & ex.StackTrace)

            Throw ex

        End Try

    End Sub

    Private Sub RunBRS5ExecuteDaily()
        'run the scheduled event
        _logger.LogInfo("Starting BRS5 Execute Daily process [" & LocalConstants.EXECUTE_FILENAME & "]")

        Dim p As Process = System.Diagnostics.Process.Start(LocalConstants.EXECUTE_FILENAME)

        'wait for it to finish a max 20 hours
        Dim result As Boolean = False
        result = p.WaitForExit(72000000)

        If result = False Then
            'finished/timed out
            _logger.LogInfo("Timed out waiting for BRS5 Execute Daily process to exit")
        Else
            'finished
            _logger.LogInfo("BRS5 Execute Daily process complete")

        End If
    End Sub

End Class
