Public Class Execute

    Dim _databaseControl As BMSL.DatabaseControl = Nothing
    Dim _logger As BMSL.LogControl = Nothing

    Public Sub New()
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)
        _databaseControl = New BMSL.DatabaseControl(_logger, LocalConstants.SQL_CONNECTION_STRING)

        'ExecuteDaily("BAL", 1, 1, 1)
        _logger.LogInfo("Program starting up")

    End Sub

    Public Sub Execute_TimingCheck_CreateBatch_RunDaily(ByVal clientCode As String, ByVal clientID As Integer, ByVal fileID As Integer, ByVal batchType As Integer)

        Dim timingNewDate As Boolean = False
        Dim startedChecking As Date = Now()

        timingNewDate = TimingCheckForNewDate(clientID, fileID, batchType)
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

                timingNewDate = TimingCheckForNewDate(clientID, fileID, batchType)

                'check if 16 hours have gone by
                If DateDiff(DateInterval.Hour, startedChecking, Now()) > 16 Then
                    _logger.LogInfo("Timed out waiting for Cervello Daily Processing date to change")
                    End
                End If

            Loop Until timingNewDate = True
        End If

        If timingNewDate Then
            Dim batchControlID As Integer = CreateBatch_CervelloProcessingDateRange(clientCode, clientID, fileID, batchType)

            Execute_Daily(clientCode, clientID, fileID, batchType, batchControlID)

        Else
            Throw New System.Exception("Execute Daily with Cervello Check - Check timing error")

        End If

    End Sub

    Public Sub Execute_ExistingBatch_RunDaily(ByVal clientCode As String, ByVal clientID As Integer, ByVal fileID As Integer, ByVal batchType As Integer)

        _logger.LogInfo("Checking for existing ready batch")

        'BRS5 data
        Dim sqlString2 As String = "app.usp_Batch_Control__Select_Min_Ready"
        Dim sqlData2 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString2)
        sqlData2.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
        sqlData2.AddParameter("@File_ID", SqlDbType.SmallInt, 2, fileID)
        sqlData2.AddParameter("@Batch_Type", SqlDbType.SmallInt, 2, batchType)

        Dim dsDataset2 As DataSet = sqlData2.SelectDataset

        If dsDataset2.Tables(0).Rows.Count = 0 Then
            _logger.LogInfo("There are no existing ready batches for Client ID [" & clientID & "] File ID [" & fileID & "] Batch Type [" & batchType & "]")

        Else
            Dim dsRow As DataRow = dsDataset2.Tables(0).Rows(0)

            Dim batchControlID As Integer = 0
            If Not dsRow.IsNull("Batch_Control_ID") Then batchControlID = dsRow.Item("Batch_Control_ID")

            If batchControlID = 0 Then
                Throw New Exception("Batch Control ID 0 is not valid.")
            Else
                Execute_Daily(clientCode, clientID, fileID, batchType, batchControlID)

            End If
            
        End If
        
        
    End Sub

    Public Function TimingCheckForNewDate(ByVal clientID As Integer, ByVal fileID As Integer, ByVal batchType As Integer) As Boolean
        _logger.LogInfo("Checking Cervello Daily Processing timing")

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

    Public Sub Execute_Migration(ByVal clientCode As String, ByVal clientID As Integer, ByVal fileID As Integer, ByVal batchType As Integer)

        Dim batchControlID As Integer = 0

        Try

            'batch control ID output

            _logger.LogInfo("Executing migration output for Client Code [" & clientCode & "] Client ID [" & clientID & "] File ID [" & fileID & "] Batch Type [" & batchType & "]")

            ''----------------------
            ''execute part 1
            ''----------------------
            ''Insert processing date and get new Batch_Control_ID
            'Dim sqlString1 As String = "brs5.usp_Batch_Control__Insert_Cervello_Processing_Date_Range"
            'Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            'sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
            'sqlData1.AddParameter("@File_ID", SqlDbType.Int, 4, fileID)
            'sqlData1.AddParameter("@Batch_Type", SqlDbType.Int, 4, batchType)
            'sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            'sqlData1.SetOutputParameter("@Batch_Control_ID")

            'batchControlID = sqlData1.ExecuteNonQuery("@Batch_Control_ID")

            '_logger.LogInfo("New Batch Control ID created: [" & batchControlID & "]")

            'get next batch ID
            _logger.LogInfo("Selecting next job from Batch Control")

            Dim sqlString1 As String = "brs5.usp_Batch_Control__Select_Next_Ready"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
            sqlData1.AddParameter("@File_ID", SqlDbType.Int, 4, fileID)
            sqlData1.AddParameter("@Batch_Type", SqlDbType.Int, 4, batchType)

            Dim dsDataset As DataSet = sqlData1.SelectDataset
            If dsDataset.Tables(0).Rows.Count = 0 Then
                'no data found
                _logger.LogInfo("No ready job found")
                Threading.Thread.Sleep(5000)
                Exit Sub
            End If
            'get data
            If Not dsDataset.Tables(0).Rows(0).IsNull("Batch_Control_ID") Then
                batchControlID = dsDataset.Tables(0).Rows(0).Item("Batch_Control_ID")

            End If

            'check validity of data
            If batchControlID = 0 Then
                _logger.LogInfo("Batch Control ID found [0] is invalid")
                Threading.Thread.Sleep(5000)
                Exit Sub
            Else
                _logger.LogInfo("Running batch [" & batchControlID & "]")

            End If

            'update batch to in progress
            UpdateBatchStatus(batchControlID, 2)

            ''Update job control and move to history
            '_logger.LogInfo("Updating job control/moving job control to history")
            'InsertCurrentStatus(batchControlID, "D1 Update Job Control/Move to History")

            'Dim sqlString2 As String = "client.usp_Job_Control__Move_Prior_Complete_To_History"
            'Dim sqlData2 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString2)
            'sqlData2.AddParameter("@UserCode", SqlDbType.VarChar, 50, clientCode)
            'Dim rowsAffected2 As Integer = sqlData2.ExecuteNonQuery()

            '_logger.LogInfo(rowsAffected2 & " row(s) affected")

            'Dim sqlString3 As String = "client.usp_Job_Control__Update_With_BK_Production_Status"
            'Dim sqlData3 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString3)
            'sqlData3.AddParameter("@UserCode", SqlDbType.VarChar, 50, clientCode)
            'Dim rowsAffected3 As Integer = sqlData3.ExecuteNonQuery()

            '_logger.LogInfo(rowsAffected3 & " row(s) affected")

            'Move latest BRS5 to history
            _logger.LogInfo("Moving latest BRS5 to history")
            InsertCurrentStatus(batchControlID, "D1 Move Latest BRS5 to History")

            Dim sqlString4 As String = "run.usp__Move_Latest_BRS5_to_History"
            Dim sqlData4 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString4)
            Dim rowsAffected4 As Integer = sqlData4.ExecuteNonQuery()

            _logger.LogInfo(rowsAffected4 & " row(s) affected")

            'Run production daily Cervello part 1
            _logger.LogInfo("Running Cervello Migration Part1")
            InsertCurrentStatus(batchControlID, "M1 Run All Cervello Custom Set Part1")

            Dim sqlString5 As String = "run.usp__Run_All__Cervello__Custom_Set_Part1"
            Dim sqlData5 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString5)
            sqlData5.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            Dim rowsAffected5 As Integer = sqlData5.ExecuteNonQuery()

            _logger.LogInfo(rowsAffected5 & " row(s) affected")

            '----------------------
            'format addressess
            '----------------------
            _logger.LogInfo("Formatting Addresses")

            Dim formatter As New FormatBRS5Address
            formatter.Execute()

            '----------------------
            'execute part 2
            '----------------------
            _logger.LogInfo("Running Cervello Migration Part2")
            InsertCurrentStatus(batchControlID, "D2 Run All Cervello Custom Set Part2")

            Dim sqlString6 As String = "run.usp__Run_All__Cervello__Custom_Set_Part2"
            Dim sqlData6 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString6)
            sqlData6.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            sqlData6.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
            Dim rowsAffected6 As Integer = sqlData6.ExecuteNonQuery()

            _logger.LogInfo(rowsAffected6 & " row(s) affected")

            '----------------------
            'generate output
            '----------------------

            'run SSIS packages
            'rename files
            'move files
            _logger.LogInfo("Generating SSIS output")

            Dim ssisOutput As New OutputSSISFiles(_databaseControl, _logger)
            ssisOutput.Execute(clientCode)

            'update batch to in progress
            UpdateBatchStatus(batchControlID, 10)

            _logger.LogInfo("Execute daily output complete")

        Catch ex As Exception
            _logger.LogInfo("** EXCEPTION **: " & ex.Message & ex.StackTrace)

            'update batch to in progress
            UpdateBatchStatus(batchControlID, 99)

            Throw ex

        End Try

    End Sub

    Public Function CreateBatch_CervelloProcessingDateRange(ByVal clientCode As String, ByVal clientID As Integer, ByVal fileID As Integer, ByVal batchType As Integer) As Integer
        Dim batchControlID As Integer = 0

        Try

            'batch control ID output

            _logger.LogInfo("Creating batch (Cervello Processing Date Range) for Client Code [" & clientCode & "] Client ID [" & clientID & "] File ID [" & fileID & "] Batch Type [" & batchType & "]")

            '----------------------
            'execute part 1
            '----------------------
            'Insert processing date and get new Batch_Control_ID
            Dim sqlString1 As String = "brs5.usp_Batch_Control__Insert_Cervello_Processing_Date_Range"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
            sqlData1.AddParameter("@File_ID", SqlDbType.Int, 4, fileID)
            sqlData1.AddParameter("@Batch_Type", SqlDbType.Int, 4, batchType)
            sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            sqlData1.SetOutputParameter("@Batch_Control_ID")

            batchControlID = sqlData1.ExecuteNonQuery("@Batch_Control_ID")

            _logger.LogInfo("New Batch Control ID created: [" & batchControlID & "]")

            Return batchControlID

        Catch ex As Exception
            _logger.LogInfo("** EXCEPTION **: " & ex.Message & ex.StackTrace)

            'update batch to in progress
            UpdateBatchStatus(batchControlID, 99)

            Throw ex

        End Try

    End Function

    Public Sub Execute_Daily(ByVal clientCode As String, ByVal clientID As Integer, ByVal fileID As Integer, ByVal batchType As Integer, ByVal batchControlID As Integer)

        ' Dim batchControlID As Integer = 0

        Try

            ''batch control ID output

            '_logger.LogInfo("Executing daily output for Client Code [" & clientCode & "] Client ID [" & clientID & "] File ID [" & fileID & "] Batch Type [" & batchType & "]")

            ''----------------------
            ''execute part 1
            ''----------------------
            ''Insert processing date and get new Batch_Control_ID
            'Dim sqlString1 As String = "brs5.usp_Batch_Control__Insert_Cervello_Processing_Date_Range"
            'Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            'sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
            'sqlData1.AddParameter("@File_ID", SqlDbType.Int, 4, fileID)
            'sqlData1.AddParameter("@Batch_Type", SqlDbType.Int, 4, batchType)
            'sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            'sqlData1.SetOutputParameter("@Batch_Control_ID")

            'batchControlID = sqlData1.ExecuteNonQuery("@Batch_Control_ID")

            '_logger.LogInfo("New Batch Control ID created: [" & batchControlID & "]")

            'update batch to in progress
            UpdateBatchStatus(batchControlID, 2)

            'Update job control and move to history
            _logger.LogInfo("Updating job control/moving job control to history")
            InsertCurrentStatus(batchControlID, "D1 Update Job Control/Move to History")

            Dim sqlString2 As String = "client.usp_Job_Control__Move_Prior_Complete_To_History"
            Dim sqlData2 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString2)
            sqlData2.AddParameter("@UserCode", SqlDbType.VarChar, 50, clientCode)
            Dim rowsAffected2 As Integer = sqlData2.ExecuteNonQuery()

            _logger.LogInfo(rowsAffected2 & " row(s) affected")

            Dim sqlString3 As String = "client.usp_Job_Control__Update_With_BK_Production_Status"
            Dim sqlData3 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString3)
            sqlData3.AddParameter("@UserCode", SqlDbType.VarChar, 50, clientCode)
            Dim rowsAffected3 As Integer = sqlData3.ExecuteNonQuery()

            _logger.LogInfo(rowsAffected3 & " row(s) affected")

            'Move latest BRS5 to history
            _logger.LogInfo("Moving latest BRS5 to history")
            InsertCurrentStatus(batchControlID, "D1 Move Latest BRS5 to History")

            Dim sqlString4 As String = "run.usp__Move_Latest_BRS5_to_History"
            Dim sqlData4 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString4)
            Dim rowsAffected4 As Integer = sqlData4.ExecuteNonQuery()

            _logger.LogInfo(rowsAffected4 & " row(s) affected")

            'Run production daily Cervello part 1
            _logger.LogInfo("Running Cervello Daily Production Part1")
            InsertCurrentStatus(batchControlID, "D1 Run All Cervello Daily Production Part1")

            Dim sqlString5 As String = "run.usp__Run_All__Cervello__Daily_Production_Part1"
            Dim sqlData5 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString5)
            sqlData5.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            Dim rowsAffected5 As Integer = sqlData5.ExecuteNonQuery()

            _logger.LogInfo(rowsAffected5 & " row(s) affected")

            '----------------------
            'format addressess
            '----------------------
            _logger.LogInfo("Formatting Addresses")

            Dim formatter As New FormatBRS5Address
            formatter.Execute()

            '----------------------
            'execute part 2
            '----------------------
            _logger.LogInfo("Running Cervello Daily Production Part2")
            InsertCurrentStatus(batchControlID, "D2 Run All Cervello Daily Production Part2")

            Dim sqlString6 As String = "run.usp__Run_All__Cervello__Daily_Production_Part2"
            Dim sqlData6 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString6)
            sqlData6.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            sqlData6.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
            Dim rowsAffected6 As Integer = sqlData6.ExecuteNonQuery()

            _logger.LogInfo(rowsAffected6 & " row(s) affected")

            '----------------------
            'generate output
            '----------------------

            'run SSIS packages
            'rename files
            'move files
            _logger.LogInfo("Generating SSIS output")

            Dim ssisOutput As New OutputSSISFiles(_databaseControl, _logger)
            ssisOutput.Execute(clientCode)

            'update batch to in progress
            UpdateBatchStatus(batchControlID, 10)

            _logger.LogInfo("Execute daily output complete")

        Catch ex As Exception
            _logger.LogInfo("** EXCEPTION **: " & ex.Message & ex.StackTrace)

            'update batch to in progress
            UpdateBatchStatus(batchControlID, 99)

            Throw ex

        End Try

    End Sub

    Public Sub Execute_OutputOnly(ByVal clientCode As String)
        Try
            _logger.LogInfo("Generating SSIS output")

            Dim ssisOutput As New OutputSSISFiles(_databaseControl, _logger)
            ssisOutput.Execute(clientCode)

        Catch ex As Exception
            _logger.LogInfo("** EXCEPTION **: " & ex.Message & ex.StackTrace)
            Throw ex


        End Try
    End Sub

    Public Sub InsertCurrentStatus(ByVal batchControlID As Integer, ByVal currentStatus As String)

        Dim sqlString1 As String = "run.usp_Current_Status__Insert"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        sqlData1.AddParameter("@Current_Status", SqlDbType.VarChar, 250, currentStatus)

        Dim rowsAffected1 As Integer = 0
        rowsAffected1 = sqlData1.ExecuteNonQuery()

    End Sub

    Public Sub UpdateBatchStatus(ByVal batchControlID As Integer, ByVal status As Integer)

        'usp_Batch_Control__Update_Status
        Dim sqlString1 As String = "brs5.usp_Batch_Control__Update_Status"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        sqlData1.AddParameter("@Batch_Status", SqlDbType.SmallInt, 2, status)

        Dim rowsAffected1 As Integer = 0
        rowsAffected1 = sqlData1.ExecuteNonQuery()

    End Sub
End Class
