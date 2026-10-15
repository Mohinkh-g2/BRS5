Imports System.Threading
Imports BRS5.Controller

Public Class ExecuteDaily

    Public Const BATCH_TYPE_DAILY_GENERATION As Integer = 1
    Public Const BATCH_TYPE_CLIENT_OUTPUT_NO_BVM As Integer = 3
    Public Const BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM As Integer = 4
    Private Shared _databaseControl As BMSL.DatabaseControl = Nothing
    Private Shared _logger As BMSL.LogControl = Nothing
    Private Shared _brs5Generator As BRS5.Generator.GenerateBRS5 = Nothing
    Private Shared Function AlreadyRunning() As Boolean
        ' Get our process name.
        Dim my_proc As Process = Process.GetCurrentProcess
        Dim my_name As String = my_proc.ProcessName
        ' Get information about processes with this name.
        Dim procs() As Process =
            Process.GetProcessesByName(my_name)

        ' If there is only one, it's us.
        If procs.Length = 1 Then Return False

        ' If there is more than one process,
        ' see if one has a StartTime before ours.
        Dim i As Integer
        For i = 0 To procs.Length - 1
            If procs(i).StartTime < my_proc.StartTime Then _
                Return True
        Next i

        ' If we get here, we were first.
        Return False
    End Function

    Shared Sub Main()

        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)

        Do While (AlreadyRunning())
            _logger.LogInfo("Another Instance is still running , Sleeping for some time to finish existing process")
            Thread.Sleep(60000)
        Loop

        Try
            _databaseControl = New BMSL.DatabaseControl(_logger, LocalConstants.SQL_CONNECTION_STRING)
            _brs5Generator = New BRS5.Generator.GenerateBRS5(LocalConstants.LOG_PATH, LocalConstants.SQL_CONNECTION_STRING)
            _logger.LogInfo("Program starting up")

            Dim ClientToProcess As String = System.Environment.GetCommandLineArgs(1)
            Dim dsBatch As DataSet

            dsBatch = _brs5Generator.GetBatchToProcess(ClientToProcess)

            If dsBatch.Tables.Count = 0 Then
                _logger.LogInfo("No batch found! (dsBatch dataset do not have any table)")
                ' _logger.LogInfo("Generate email")
            ElseIf dsBatch.Tables(0).Rows.Count < 1 Then
                _logger.LogInfo("No batch found!")
            ElseIf dsBatch.Tables(0).Rows.Count > 1 Then
                _logger.LogInfo("Multiple batches found when only one batch is expected. Batch Count: " + dsBatch.Tables(0).Rows.Count.ToString())
            Else
                Dim BatchControlID As Integer = dsBatch.Tables(0).Rows(0).Item("Batch_Control_ID")
                _logger.LogInfo("Batch found: BatchControlID " + BatchControlID.ToString())

                Try
                    _logger.LogInfo("Starting Executing ClientID: " + ClientToProcess + "| BatchControlID: " + BatchControlID.ToString())

                    Execute_Daily(BatchControlID, ClientToProcess)

                Catch ex As Exception

                    '_brs5Generator.UpdateBatchStatus(BatchControlID, 99) 'failed / error - Other than that is already handled.

                    _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)

                End Try

            End If

            _logger.LogInfo("Program shutting down")

        Catch ex As Exception
            _logger.LogError("Exception was unrecoverable: " & ex.Message & " " & ex.StackTrace)
        End Try
    End Sub
    'Deepak Altura-39
    Private Shared Sub Execute_Daily(ByVal BatchControlID, ByVal ClientToProcess)

        _logger.LogInfo("Updating the BatchControlID:" + BatchControlID.ToString() + " Status : 2")

        _brs5Generator.UpdateBatchStatus(BatchControlID, 2)


        Dim dsBatchExecutionSteps As DataSet
        dsBatchExecutionSteps = _brs5Generator.GetBatchExecutionSteps(BatchControlID)

        'Check if Client Is subscribed to BIS , BIS is loaded or not

        For Each execution_step As DataRow In dsBatchExecutionSteps.Tables(0).Rows
            If execution_step.Item("Description").contains("BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM") Then
                If Not CheckIfBISAccountsLoaded() Then
                    _logger.LogInfo("BIS Data is not loaded yet , exiting the Application")
                    _brs5Generator.UpdateBatchStatus(BatchControlID, 9)
                    _brs5Generator.UpdateNextBatchDate(BatchControlID)
                    Return
                End If
            End If
        Next

        For Each execution_step As DataRow In dsBatchExecutionSteps.Tables(0).Rows
            Select Case execution_step.Item("Description")

                Case "BATCH_TYPE_DAILY_GENERATION"

                    _logger.LogInfo("Starting: BATCH_TYPE_DAILY_GENERATION")
                    Try

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 2)
                        CommonTasksandOutput(BatchControlID, ClientToProcess)
                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 10)

                    Catch ex As Exception

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 99)
                        _brs5Generator.UpdateBatchStatus(BatchControlID, 99)
                        ' _brs5Generator.GenerateEmail(BatchControlID)

                        _brs5Generator.InsertCurrentStatus(BatchControlID, "Common Tasks and Output Generation Error")
                        _brs5Generator.UpdateNextBatchDate(BatchControlID)
                        _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)
                        Throw ex

                    End Try

                    _logger.LogInfo("Done: BATCH_TYPE_DAILY_GENERATION")

                Case "TRANSFER_DATA_FROM_ALTURA_TO_ALTURA_LITE"

                    _logger.LogInfo("Starting Step: TRANSFER_DATA_FROM_ALTURA_TO_ALTURA_LITE")

                    Try
                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 2)
                        _brs5Generator.TransferDataToAlturaLite(BatchControlID)
                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 10)
                        _logger.LogInfo("Data Transferred From Altura to Altura Lite")

                    Catch ex As Exception

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 99)
                        _brs5Generator.UpdateBatchStatus(BatchControlID, 99)
                        _brs5Generator.InsertCurrentStatus(BatchControlID, "Error in transferring data to Altura_Lite")
                        _brs5Generator.UpdateNextBatchDate(BatchControlID)
                        _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)
                        Throw ex
                    End Try

                    _logger.LogInfo("Done Step: TRANSFER_DATA_FROM_ALTURA_TO_ALTURA_LITE")

                Case "BATCH_TYPE_CLIENT_OUTPUT_NO_BVM"

                    _logger.LogInfo("Starting: BATCH_TYPE_CLIENT_OUTPUT_NO_BVM")

                    Try
                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 2)
                        ProcessBatchType3(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")))
                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 10)

                    Catch ex As Exception

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 99)
                        _brs5Generator.UpdateBatchStatus(BatchControlID, 99)
                        _brs5Generator.InsertCurrentStatus(BatchControlID, "Error in BATCH_TYPE_CLIENT_OUTPUT_NO_BVM")
                        _brs5Generator.UpdateNextBatchDate(BatchControlID)
                        _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)
                        Throw ex
                    End Try

                    _logger.LogInfo("Done: BATCH_TYPE_CLIENT_OUTPUT_NO_BVM")

                Case "BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM"

                    _logger.LogInfo("Starting: BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM")

                    Try

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 2)
                        ProcessBatchType4(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")))

                        _logger.LogInfo("Updating the BatchControlID: " + BatchControlID.ToString() + "| Status : 10")
                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 10)

                    Catch ex As Exception

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 99)
                        _brs5Generator.UpdateBatchStatus(BatchControlID, 99)
                        _brs5Generator.InsertCurrentStatus(BatchControlID, "Error in BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM")
                        _brs5Generator.UpdateNextBatchDate(BatchControlID)
                        _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)
                        Throw ex
                    End Try

                    _logger.LogInfo("Done: BATCH_TYPE_CLIENT_OUTPUT_WITH_BVM")

                Case "BRS5_Batch_Audit"

                    _logger.LogInfo("Starting: BRS5_Batch_Audit")

                    Try

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 2)
                        _brs5Generator.BRS5AuditProcess(BatchControlID)
                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 10)

                    Catch ex As Exception

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 99)
                        _brs5Generator.UpdateBatchStatus(BatchControlID, 98)
                        _logger.LogInfo("Updating the BatchControlID (Error in Audit): " + BatchControlID.ToString() + "| Status : 98")
                        _brs5Generator.InsertCurrentStatus(BatchControlID, "Error in BRS5_Batch_Audit")

                        _logger.LogError("Warning:  " & ex.Message & " " & ex.StackTrace)
                        Throw ex

                    End Try

                    _logger.LogInfo("Done: BRS5_Batch_Audit")

                Case "EXECUTE_TU_BRS5_SSIS_PACKAGE"

                    _logger.LogInfo("Starting: EXECUTE_TU_BRS5_SSIS_PACKAGE")

                    Try

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 2)
                        Dim outputBRS5SSIS As New BRS5.Generator.OutputSSISFiles(_databaseControl, _logger, LocalConstants.PACKAGE_FOLDER, LocalConstants.OUTPUT_FOLDER, LocalConstants.DELIVERY_FOLDER, LocalConstants.ON_HOLD_FOLDER)
                        outputBRS5SSIS.OutputFile_TU_BRS5("TU_BRS5.dtsx")
                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 10)
                    Catch ex As Exception

                        _brs5Generator.UpdateBatchExecutionStep(BatchControlID, Convert.ToInt16(execution_step.Item("Settings_Id")), 99)
                        _brs5Generator.UpdateBatchStatus(BatchControlID, 97)
                        _logger.LogInfo("Updating the BatchControlID (Error in TU SSIS Execution): " + BatchControlID.ToString() + "| Status : 97")
                        _brs5Generator.InsertCurrentStatus(BatchControlID, "Error in EXECUTE_TU_BRS5_SSIS_PACKAGE")

                        _logger.LogError("Warning:  " & ex.Message & " " & ex.StackTrace)
                        Throw ex

                    End Try

                    _logger.LogInfo("Done: EXECUTE_TU_BRS5_SSIS_PACKAGE")

                Case Else
                    Throw New Exception("Invalid Batch Type: " + 1)
            End Select
        Next

        _brs5Generator.UpdateBatchStatus(BatchControlID, 10)

    End Sub
    Private Shared Sub Execute_WaitForNewProcessingDay()

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
                    _logger.LogInfo("Timed out waiting for Cervello Daily Processing date to change")
                    End
                End If

            Loop Until timingNewDate = True
        End If

        If timingNewDate Then
            'create a new batch control row 
            Dim batchControlID As Integer = _brs5Generator.InsertBatchControl_CervelloProcessingDateRange(BATCH_TYPE_DAILY_GENERATION)

            Try
                'start batch - update status
                _brs5Generator.UpdateBatchStatus(batchControlID, 2)
                _brs5Generator.InsertCurrentStatus(batchControlID, "Daily Generation Starting")

                '-----------------------
                'DAILY PROCESS START
                '-----------------------
                'Execute_Daily(clientCode, clientID, fileID, batchType, batchControlID)

                'Update client-specific tables 
                _brs5Generator.UpdateClientSpecificTables(batchControlID, "BAL")

                'Move previous BRS5 data to history, and clear staging tables 
                _brs5Generator.MovePreviousBRS5ToHistory_ClearPreviousStagingTables(batchControlID)

                'Insert cervello data for new batch control row 
                _brs5Generator.InsertBatchCases_CervelloProcessingDateRange(batchControlID)

                'Insert override cases
                _brs5Generator.UpdateBatchCases_Override(batchControlID)

                'Insert data from Cervello into BRS5 staging
                _brs5Generator.InsertCervelloIntoBRS5Staging(batchControlID)

                'Insert data from BRS5 staging into BRS5
                _brs5Generator.InsertBRS5StagingIntoBRS5(batchControlID, 1)

                'Output SSIS
                Dim outputSSIS As New BRS5.Generator.OutputSSISFiles(_databaseControl, _logger, LocalConstants.PACKAGE_FOLDER, LocalConstants.OUTPUT_FOLDER, LocalConstants.DELIVERY_FOLDER, LocalConstants.ON_HOLD_FOLDER)
                'OLD GENERATE SSIS BEGIN''''''''''''''''
                'outputSSIS.Execute("BAL")
                'OLD GENERATE SSIS END ''''''''''''''''

                'NEW GENERATE SSIS BEGIN''''''''''''''''

                'generate SSIS for all
                outputSSIS.GenerateAllSSIS_exceptManifest()

                ' ******** Start
                'get SQL dataset with client codes info
                Dim sqlString1 As String = "client.usp_Delivery_Options__Select_Summary"
                Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

                Dim dsDataset1 As DataSet = sqlData1.SelectDataset

                If dsDataset1.Tables(0).Rows.Count = 0 Then
                    _logger.LogInfo("There is no existing record to deliver")

                Else
                    Dim clientCode As String = ""
                    Dim clientID As Integer = 0
                    For Each dsRow1 As DataRow In dsDataset1.Tables(0).Rows
                        If Not dsRow1.IsNull("UserCode") Then clientCode = dsRow1.Item("UserCode")

                        'populate BRS5 table with manifest for this client
                        outputSSIS.insertManifesttoBRS5(batchControlID, clientID, clientCode)

                        'generate SOURCE manifest for this client
                        outputSSIS.GenerateSSIS("Manifest.txt", "Manifest.dtsx")


                        'copy SOURCE ssis files to DEST output
                        Dim blnSavetoOnHOLD As Boolean = False
                        outputSSIS.CopyClientSSIS(clientCode, blnSavetoOnHOLD)

                        'delete SOURCE manifest
                        outputSSIS.DeleteSSIS("Manifest.txt")

                    Next dsRow1    ' next clientCode
                End If



                'delete all SOURCE SSIS output
                outputSSIS.DeleteAllSSIS()

                'NEW GENERATE SSIS END''''''''''''''''

                '-----------------------
                'DAILY PROCESS END
                '-----------------------

                'finish batch - update status 
                _brs5Generator.InsertCurrentStatus(batchControlID, "Daily Generation Complete")
                _brs5Generator.UpdateBatchStatus(batchControlID, 10)

            Catch ex As Exception
                'error batch - update status 
                _brs5Generator.UpdateBatchStatus(batchControlID, 99)
                _brs5Generator.InsertCurrentStatus(batchControlID, "Daily Generation Error")

                'Altura-51
                'Update the next client batch to current batch start date for reprocessing date range.
                _brs5Generator.UpdateNextBatchDate(batchControlID)

                Throw ex

            End Try

        Else
            Throw New System.Exception("Execute Daily with Cervello Check - Check timing error")

        End If

    End Sub


    Public Shared Sub CommonTasksandOutput(ByVal batchControlID As Integer, ByVal ClientToProcess As String)

        Dim outputSSIS As New BRS5.Generator.OutputSSISFiles(_databaseControl, _logger, LocalConstants.PACKAGE_FOLDER, LocalConstants.OUTPUT_FOLDER, LocalConstants.DELIVERY_FOLDER, LocalConstants.ON_HOLD_FOLDER)

        _logger.LogInfo("Deleting all SSIS output files")
        outputSSIS.DeleteAllSSIS()

        _brs5Generator.InsertCurrentStatus(batchControlID, "BATCH_TYPE_DAILY_GENERATION")

        If ClientToProcess.ToUpper = "BAL" Then
            _brs5Generator.UpdateClientSpecificTables(batchControlID, "BAL")
        End If

        _logger.LogInfo("Move Previous run to History")
        _brs5Generator.MovePreviousBRS5ToHistory_ClearPreviousStagingTables(batchControlID)     ' excluding BVM

        _logger.LogInfo("Inserting Batch Cases")
        _brs5Generator.InsertBatchCases_CervelloProcessingDateRange(batchControlID)

        If ClientToProcess.Trim.ToLower() = "pra" Then
            _logger.LogInfo(ClientToProcess + ": Determine BVM Eligible Cases")
            _brs5Generator.BvmDetermineEligibleCases(batchControlID)                            ' All BA cases are eligible
            _logger.LogInfo(ClientToProcess + "Insert Cases to Override")
            _brs5Generator.BvmInsertCasesToOverride(batchControlID)                             'Override BVM eligible cases that do not exist in batch case
        End If

        _logger.LogInfo("Add overridden cases to Batch Cases")
        _brs5Generator.UpdateBatchCases_Override(batchControlID)

        _logger.LogInfo("Fetch data from Cervello into Staging tables")
        _brs5Generator.InsertCervelloIntoBRS5Staging(batchControlID)                            ' excluding BVM

        _logger.LogInfo("Inserting Staging data to BRS5 tables")
        _brs5Generator.InsertBRS5StagingIntoBRS5(batchControlID, 1)                             ' excluding BVM

        If ClientToProcess.Trim.ToLower() = "pra" Or ClientToProcess.Trim.ToLower() = "bal" Then

            _logger.LogInfo(ClientToProcess + ": Generating output files")
            outputSSIS.GenerateAllSSIS_exceptManifest()                                         ' excluding BVM
        End If


    End Sub

    Private Shared Sub ProcessBatchType3(ByVal batchControlID As Integer, ByVal batchTypeID As Integer)

        Dim outputSSIS As New BRS5.Generator.OutputSSISFiles(_databaseControl, _logger, LocalConstants.PACKAGE_FOLDER, LocalConstants.OUTPUT_FOLDER, LocalConstants.DELIVERY_FOLDER, LocalConstants.ON_HOLD_FOLDER)

        Dim sqlString1 As String = "client.usp_Delivery_Options__Select_Summary"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        'sqlData1.AddParameter("@BatchType", SqlDbType.SmallInt, 2, batchTypeID)
        sqlData1.AddParameter("@BatchControlID", SqlDbType.SmallInt, 2, batchControlID)


        Dim dsDataset1 As DataSet = sqlData1.SelectDataset

        If dsDataset1.Tables(0).Rows.Count = 0 Then
            _logger.LogInfo("There is no existing record to deliver")

        Else
            Dim clientCode As String = ""
            Dim clientID As Integer = 0
            For Each dsRow1 As DataRow In dsDataset1.Tables(0).Rows
                If Not dsRow1.IsNull("UserCode") Then clientCode = dsRow1.Item("UserCode")

                outputSSIS.insertManifesttoBRS5(batchControlID, clientID, clientCode)

                outputSSIS.GenerateSSIS("Manifest.txt", "Manifest.dtsx")

                Dim blnSavetoOnHOLD As Boolean = False

                outputSSIS.CopyClientSSIS(clientCode, blnSavetoOnHOLD)

                outputSSIS.DeleteSSIS("Manifest.txt")

            Next dsRow1    ' next clientCode
        End If

    End Sub

    Private Shared Sub ProcessBatchType4(ByVal batchControlID As Integer, ByVal batchType As Integer)

        ' Procedure to Load into BVM Subsystem the case list 
        ' The procedure below temporarily passes a static Debtor Case Indicator of 1
        ' add a feature that will determine Debtor Case Indicator and depending upon its value, create either one or two records.
        '_logger.LogInfo("Insert cases to add in BVM Case list")
        '_brs5Generator.InsertCasesToAddInBvmCaseList(batchControlID)'deprecated; eongsioco

        '_logger.LogInfo("Insert case list to BVM subsystem")
        '_brs5Generator.InsertBVMSubsystemCaseListfromBRS5(batchControlID) 'deprecated; eongsioco

        '-- STEP 2: Queue case list for processing
        'EXEC @Return = input.usp_QueueBatch @BatchQueueID OUTPUT
        'PRINT @Return
        'Dim ParentBatchControlID As Int32
        'ParentBatchControlID = _brs5Generator.GetParentBatchControlID(batchControlID)

        _brs5Generator.BvmQueueEligibleCases(batchControlID)

        Dim BatchQueueID As Integer = 0
        BatchQueueID = _brs5Generator.BvmQueueCaseListForProcessing()
        If BatchQueueID = -1 Then
            _logger.LogInfo("Could not find any Bvm eligible cases ")
            Return
        End If

        _brs5Generator.BvmValidateCaseBusinessRules(BatchQueueID)

        '-- STEP 3: BVM calculation
        'DECLARE @BatchQueueStatus AS int
        'EXEC bvm.usp_PullBisAccountsAndCalculateBvms @BatchQueueID, @BatchQueueStatus OUTPUT

        Dim BatchQueueStatus As Integer = 0
        BatchQueueStatus = _brs5Generator.BvmCalculate(BatchQueueID)

        'throw exception if Calculate BVM returned 99
        If BatchQueueStatus = 99 Then Throw New Exception("BVM BatchQueueStatus returned 99 (unhandled exception)")


        _logger.LogInfo("Move Previous BVM to History")
        _brs5Generator.BvmMoveToHistory(batchControlID)

        _logger.LogInfo("Insert BVMs from BVM subsystem to staging")
        ' Populate BVM Staging Table and Generate BVM SSIS
        _brs5Generator.BvmInsertStagingFromBVMSubsystem(batchControlID, BatchQueueID)


        ' Copy Client Specific

        'get SQL dataset with client codes info
        'Dim Usercode As String = "PRA"
        Dim sqlString1 As String = "client.usp_Delivery_Options__Select_Summary"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        'sqlData1.AddParameter("@BatchType", SqlDbType.SmallInt, 2, batchType)
        sqlData1.AddParameter("@BatchControlID", SqlDbType.SmallInt, 2, batchControlID)

        Dim dsDataset1 As DataSet = sqlData1.SelectDataset

        If dsDataset1.Tables(0).Rows.Count = 0 Then
            _logger.LogInfo("There is no existing record to deliver")

        Else
            Dim clientCode As String = ""
            Dim clientID As Integer = 0
            For Each dsRow1 As DataRow In dsDataset1.Tables(0).Rows
                If Not dsRow1.IsNull("UserCode") Then clientCode = dsRow1.Item("UserCode")

                _logger.LogInfo("Insert BVMs from BVM staging")
                ' Populate BVM table with filtered Data from BVM_Staging  ' add filtering 
                _brs5Generator.BvmInsertFromStaging(batchControlID, clientCode)

                _logger.LogInfo("Run SSIS to generate output")
                Dim outputSSIS As New BRS5.Generator.OutputSSISFiles(_databaseControl, _logger, LocalConstants.PACKAGE_FOLDER, LocalConstants.OUTPUT_FOLDER, LocalConstants.DELIVERY_FOLDER, LocalConstants.ON_HOLD_FOLDER)
                outputSSIS.GenerateSSIS("BVM.txt", "BVM.dtsx")


                'populate BRS5 table with manifest for this client
                outputSSIS.insertManifesttoBRS5(batchControlID, clientID, clientCode)

                'generate SOURCE manifest for this client
                outputSSIS.GenerateSSIS("Manifest.txt", "Manifest.dtsx")


                'copy SOURCE ssis files to DEST output
                ' added below option to save the output text files in an alternate folder
                Dim blnSavetoOnHOLD As Boolean = False

                If BatchQueueStatus = 11 Then
                    blnSavetoOnHOLD = True
                End If
                outputSSIS.CopyClientSSIS(clientCode, blnSavetoOnHOLD)

                'delete SOURCE manifest
                outputSSIS.DeleteSSIS("Manifest.txt")

            Next dsRow1    ' next clientCode

        End If

        _brs5Generator.InsertCurrentStatus(batchControlID, "Completed Batch Type 4")

    End Sub

    Private Shared Function CheckIfNewProcessingDay(ByVal clientID As Integer, ByVal fileID As Integer, ByVal batchType As Integer) As Boolean
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
    Private Shared Function CheckIfBISAccountsLoaded() As Boolean
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

End Class
