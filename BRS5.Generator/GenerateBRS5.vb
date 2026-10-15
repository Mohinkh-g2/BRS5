Imports System.Net
Imports System.Net.Mail
Imports System.Configuration
Imports System.Threading

Public Class GenerateBRS5

    Dim _databaseControl As BMSL.DatabaseControl = Nothing
    Dim _logger As BMSL.LogControl = Nothing
    Dim ClientToProcess As String = System.Environment.GetCommandLineArgs(0)

    Public Sub New(ByVal logPath As String, ByVal sqlConnectionString As String)

        _logger = New BMSL.LogControl(logPath)
        _databaseControl = New BMSL.DatabaseControl(_logger, sqlConnectionString)

        Dim args As String() = System.Environment.GetCommandLineArgs()

        If args.Length > 1 Then
            ClientToProcess = args(1)
            Console.WriteLine("ClientToProcess:" & args(1))

        End If

    End Sub

    Public Sub InsertCervelloIntoBRS5Staging(ByVal batchControlID As Integer)

        _logger.LogInfo("Batch [" & batchControlID & "]" & "Inserting Cervello data into BRS5 Staging")

        Dim sqlString1 As String = "brs5_staging.usp__Run_All"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub

    Public Sub InsertBRS5StagingIntoBRS5(ByVal batchControlID As Integer, ByVal clientID As Integer)

        _logger.LogInfo("Batch [" & batchControlID & "]" & "Inserting BRS5 Staging data into BRS5")

        Dim sqlString1 As String = "brs5.usp__Run_All"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub

    Public Sub MovePreviousBRS5ToHistory_ClearPreviousStagingTables(ByVal batchControlID As Integer)

        'Move latest BRS5 to history
        _logger.LogInfo("Moving previous BRS5 to history")
        _logger.LogInfo("Clearing previous BRS5 from staging tables")

        InsertCurrentStatus(batchControlID, "D1 Move Latest BRS5 to History")

        Dim sqlString4 As String = "run.usp__Move_Latest_BRS5_to_History"
        Dim sqlData4 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString4)
        Dim rowsAffected4 As Integer = sqlData4.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected4 & " row(s) affected")

    End Sub


    Public Sub UpdateClientSpecificTables(ByVal batchControlID As Integer, ByVal clientCode As String)
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


    End Sub

    Public Sub InsertBatchCases_Migration(ByVal batchControlID As Integer)

        _logger.LogInfo("Batch [" & batchControlID & "]" & "Inserting batch cases - migration")

        Dim sqlString1 As String = "brs5.usp_Batch_Cases__Insert_Migration_Cases"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub

    Public Sub UpdateBatchCases_Override(ByVal batchControlID As Integer)

        _logger.LogInfo("Batch [" & batchControlID & "]" & "Updating batch cases - inserting override cases")

        Dim sqlString1 As String = "brs5.usp_Batch_Cases__Update_w_Override"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub
    Public Function InsertBatchCases_CervelloProcessingDateRange(ByVal batchControlID As Integer) As Integer

        _logger.LogInfo("Batch [" & batchControlID & "]" & "Inserting batch cases - Daily Cervello Processing Date Range")

        Dim sqlString1 As String = "brs5.usp_Batch_Cases__Insert_Cervello_Processing_Date_Range"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()
        _logger.LogInfo(rowsAffected & " row(s) affected")
        Return rowsAffected
    End Function

    Public Function InsertBatchControl_CervelloProcessingDateRange(ByVal batchType As Integer) As Integer

        Dim batchControlID As Integer = 0

        Try

            'batch control ID output

            _logger.LogInfo("Inserting batch control row - Daily Cervello Processing Date Range")

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

    Public Sub InsertCurrentStatus(ByVal batchControlID As Integer, ByVal currentStatus As String)
        Try


            Dim sqlString1 As String = "run.usp_Current_Status__Insert"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            sqlData1.AddParameter("@Current_Status", SqlDbType.VarChar, 250, currentStatus)

            Dim rowsAffected1 As Integer = 0
            rowsAffected1 = sqlData1.ExecuteNonQuery()
        Catch ex As Exception
            GenerateEmail(batchControlID, ClientToProcess)
            _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)
            Throw ex
        End Try


    End Sub

    Public Sub UpdateBatchStatus(ByVal batchControlID As Integer, ByVal status As Integer)
        Try

            Dim isEmailEnabled As Boolean = Convert.ToBoolean(ConfigurationManager.AppSettings("IsEmailEnabled"))

            Dim sqlString1 As String = "brs5.usp_Batch_Control__Update_Status"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            sqlData1.AddParameter("@Batch_Status", SqlDbType.SmallInt, 2, status)
            If status = 2 And isEmailEnabled = True Then
                GenerateEmail(batchControlID, ClientToProcess)
            End If

            Dim rowsAffected1 As Integer = 0
            rowsAffected1 = sqlData1.ExecuteNonQuery()



        Catch ex As Exception

            GenerateEmail(batchControlID, ClientToProcess)
            _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)
            Throw ex

        End Try


    End Sub

    Public Sub UpdateNextBatchDate(ByVal batchControlID As Integer)
        Try
            'usp_Batch_Control__Update_Status
            Dim sqlString1 As String = "brs5.usp_Batch_Control__Update_Date_NextRun"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

            sqlData1.AddParameter("@batchControlID", SqlDbType.Int, 4, batchControlID)
            sqlData1.AddParameter("@execResult", SqlDbType.VarChar, 100, "")
            sqlData1.SetOutputParameter("@execResult")

            Dim strMessage As String = String.Empty
            strMessage = sqlData1.ExecuteNonQuery("@execResult")

            _logger.LogInfo(strMessage)
        Catch ex As Exception
            GenerateEmail(batchControlID, ClientToProcess)
            _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)
            Throw ex
        End Try


    End Sub

    Public Function GetParentBatchControlID(ByVal BatchControlID As Int32) As Int32

        _logger.LogInfo("get ParentBatchControlID of batch [" & BatchControlID & "]")

        Dim sqlString1 As String = "app.usp_Get_Parent_Batch_Control_ID"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@BatchControlID", SqlDbType.Int, 4, BatchControlID)
        sqlData1.AddParameter("@ParentBatchControlID", SqlDbType.Int, 4, 0)
        sqlData1.SetOutputParameter("@ParentBatchControlID")
        'sqlData1.SetOutputParameter("@Return")
        Dim parentBatchControlID As Integer = 0
        parentBatchControlID = Integer.Parse(sqlData1.ExecuteNonQuery("@ParentBatchControlID"))

        If parentBatchControlID = 0 Then
            Throw New ApplicationException("***BRS5.ExecuteDaily: No valid parent batch control ID returned from app.usp_Get_Parent_Batch_Control_ID stored procedure.")
        End If
        _logger.LogInfo("ParentBatchControlID is [" & parentBatchControlID.ToString() & "]")

        Return parentBatchControlID
    End Function

#Region "BVM Functions and Subroutines"

    Public Sub BvmInsertCaseListFromBRS5(ByVal batchControlID As Integer)

        _logger.LogInfo("Batch [" & batchControlID & "]" & "Inserting Case List into BVM Subsystem from BRS5")

        Dim sqlString1 As String = "[brs5].[usp_BVM_Batch_Cases__Insert]"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub

    Public Function GetBatchToProcess(ByVal ClientCode As String) As DataSet

        _logger.LogInfo("Getting batch to process.")
        Dim sqlString1 As String = "brs5.usp_Get_Batch_To_Process"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@ClientCode", SqlDbType.VarChar, 50, ClientCode)
        Dim dsBatch As New DataSet
        dsBatch = sqlData1.SelectDataset()
        Return dsBatch
    End Function


    Public Sub BvmDetermineEligibleCases(ByVal batchControlID As Int32)

        _logger.LogInfo("Get Bvm Eligible Cases")

        Dim sqlString1 As String = "Altura_BVM.Brs5Control.usp_GetEligibleCase"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@BatchControlID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub


    Public Sub BvmInsertCasesToOverride(ByVal batchControlID As Int32)

        _logger.LogInfo("Insert BVM cases to override")

        Dim sqlString1 As String = "Altura_BVM.Brs5Control.usp_Brs5OverrideBvmEligible"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@BatchControlID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub


    Public Sub BvmQueueEligibleCases(ByVal batchControlID As Int32)

        _logger.LogInfo("Queue eligible cases")

        Dim sqlString1 As String = "Altura_BVM.Brs5Control.usp_QueueEligibleCases"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@BatchControlID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()
        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub
    Public Function BvmQueueCaseListForProcessing() As Int32
        _logger.LogInfo("Inserting Case List into BVM Subsystem from BRS5")
        Dim sqlString1 As String = "Altura_BVM.input.usp_QueueBatch"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@pBatchQueueID", SqlDbType.Int, 4, 0)
        sqlData1.SetOutputParameter("@pBatchQueueID")
        Dim Output As String = sqlData1.ExecuteNonQuery("@pBatchQueueID")
        Dim BatchQueueID As Integer = Integer.Parse(Output)
        'If BatchQueueID = -1 Then
        '    Throw New ApplicationException("***BRS5.ExecuteDaily: Error occured when Altura_BVM.input.usp_QueueBatch stored procedure was executed. This could be no cases were entered in input.CaseList table.")
        'End If
        _logger.LogInfo("Queue case batch. BatchQueueID=[" & BatchQueueID & "]")
        Return BatchQueueID
    End Function

    Public Sub BvmValidateCaseBusinessRules(ByVal BatchQueueID As Int32)

        _logger.LogInfo("Business rules validation of batchQueueID = [" & BatchQueueID & "] ")

        Dim sqlString1 As String = "Altura_BVM.Brs5Control.usp_ValidateCaseBusinessRules"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@BatchQueueID", SqlDbType.Int, 4, BatchQueueID)
        sqlData1.ExecuteNonQuery()

    End Sub

    Public Function BvmCalculate(ByVal BatchQueueID As Int32) As Int32

        _logger.LogInfo("Calculating BVM for BatchQueueID = [" & BatchQueueID & "] ")

        Dim sqlString1 As String = "Altura_BVM.bvm.usp_PullBisAccountsAndCalculateBvms"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@BatchQueueID", SqlDbType.Int, 4, BatchQueueID)
        sqlData1.AddParameter("@BatchQueueStatus", SqlDbType.Int, 4, BatchQueueID)
        sqlData1.SetOutputParameter("@BatchQueueStatus")
        Dim BatchQueueStatus As Int32 = 0
        BatchQueueStatus = sqlData1.ExecuteNonQuery("@BatchQueueStatus")

        _logger.LogInfo("BVM Calculation done. BatchQueueStatus = [" & BatchQueueStatus & "]")

        Return BatchQueueStatus
    End Function

    Public Sub BvmMoveToHistory(ByVal batchControlID As Integer)
        ' 20120920 remove reference to batchcontrolID in thelog text below
        '_logger.LogInfo("Batch [" & batchControlID & "]" & "Moving BVM to History")
        _logger.LogInfo("Moving BVM to History")

        Dim sqlString1 As String = "brs5_history.usp_BVM__Move"    ' edit SP to handle or remove BatchControlID
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        ' 20102920 commented out (1) line below (for @BatchControlID)
        'sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)

        'sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub

    Public Sub BvmInsertStagingFromBVMSubsystem(ByVal batchControlID As Integer, ByVal batchQueueID As Integer)

        _logger.LogInfo("Batch [" & batchControlID & "]" & "Inserting to BVM Staging from BVM Subsystem")

        Dim sqlString1 As String = "[brs5_staging].[usp_Staging_BVM__Insert]"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        sqlData1.AddParameter("@BatchQueueID", SqlDbType.Int, 4, batchQueueID)
        'sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub
    Public Sub BvmInsertFromStaging(ByVal batchControlID As Integer, ByVal userCode As String)

        _logger.LogInfo("Batch [" & batchControlID & "]" & "Inserting to BVM from BVM Staging")

        Dim sqlString1 As String = "brs5.usp_BVM__Insert"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        sqlData1.AddParameter("@UserCode", SqlDbType.VarChar, 10, userCode)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub

    'Public Sub InsertCasesToAddInBvmCaseList(ByVal batchControlID As Integer) 'eongsioco; 02/12/2013; deprecated
    '    _logger.LogInfo("executing brs5_staging.usp_Insert_Cases_To_Add_In_Bvm_Case_List")

    '    Dim sqlString1 As String = "brs5_staging.usp_Insert_Cases_To_Add_In_Bvm_Case_List"
    '    Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
    '    sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
    '    Dim rowsAffected As Integer = 0
    '    rowsAffected = sqlData1.ExecuteNonQuery()

    '    _logger.LogInfo(rowsAffected & " row(s) affected")
    'End Sub

    'Public Sub InsertCasesToOverrideForBvmCalculations(ByVal batchControlID As Int32) 'eongsioco; 02/12/2013; deprecated
    '    _logger.LogInfo("Inserting Cases to update BVM data")

    '    Dim sqlString1 As String = "brs5_staging.usp_Insert_Cases_To_Override_For_Bvm_Calculation"
    '    Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
    '    sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
    '    Dim rowsAffected As Integer = 0
    '    rowsAffected = sqlData1.ExecuteNonQuery()

    '    _logger.LogInfo(rowsAffected & " row(s) affected")
    'End Sub


    'Deepak Altura-39
    Public Function GetBatchExecutionSteps(ByVal batchControlID As Int32) As DataSet

        _logger.LogInfo("Getting batch's execution steps to process.")

        Dim sqlString1 As String = "brs5.usp_Get_Execution_Steps_Of_Batch"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        Dim dsBatch As New DataSet
        dsBatch = sqlData1.SelectDataset()
        Return dsBatch


    End Function
    'Deepak Altura-39
    Public Sub UpdateBatchExecutionStep(ByVal batchControlID As Int32, ByVal settingsID As Int16, ByVal batchStatus As Int16)

        Try
            _logger.LogInfo("Updating  batch's execution steps to process.")
            Dim sqlString1 As String = "brs5.usp_Update_Execution_Steps_Of_Batch"
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
            sqlData1.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
            sqlData1.AddParameter("@Setting_ID", SqlDbType.Int, 4, settingsID)
            sqlData1.AddParameter("@Status", SqlDbType.Int, 4, batchStatus)
            Dim rowsAffected As Integer = 0
            rowsAffected = sqlData1.ExecuteNonQuery()
            _logger.LogInfo(rowsAffected & " row(s) affected")

        Catch ex As Exception
            GenerateEmail(batchControlID, ClientToProcess)
            _logger.LogError("Exception:  " & ex.Message & " " & ex.StackTrace)
            Throw ex

        End Try

    End Sub
    Public Sub TransferDataToAlturaLite(ByVal batchControlID As Integer)

        Dim sqlString1 As String = "brs5.usp_Insert_Data_From_BRS5_To_Altura_Lite"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@Batch_Control_Id", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()

    End Sub
    Public Function BVM_Eligible_Case_List_Check() As Int32
        _logger.LogInfo("Checking Case List for cases to process")
        Dim sqlString1 As String = "Altura_BVM.input.usp_BVM_Eligible_Case_List_Check"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@CasesExists", SqlDbType.Int, 4, 0)
        sqlData1.SetOutputParameter("@CasesExists")
        Dim Output As String = sqlData1.ExecuteNonQuery("@CasesExists")
        Dim CaseList As Integer = Integer.Parse(Output)
        Return CaseList
    End Function
    Public Sub BRS5AuditProcess(ByVal batchControlID As Int32)

        Dim sqlString1 As String = "Altura.dbo.Daily_BRS5_Audit"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)
        sqlData1.AddParameter("@BatchControlID", SqlDbType.Int, 4, batchControlID)
        Dim rowsAffected As Integer = 0
        rowsAffected = sqlData1.ExecuteNonQuery()
        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub

    'G2BR-34
    Public Sub GenerateEmail(ByVal batchControlID As Integer, ByVal ClientToProcess As String)
        Try

            _logger.LogInfo("Sending Email")
            Dim smtp As New SmtpClient("email2.lci.com")
            Dim smtpPort As Integer = Convert.ToInt32(ConfigurationManager.AppSettings("SMTPPort"))
            Dim mailFromAddress As String = Convert.ToString(ConfigurationManager.AppSettings("MailFromAddress"))
            Dim mailToAddress As String = Convert.ToString(ConfigurationManager.AppSettings("MailToAddress"))
            Dim title As String = Convert.ToString(ConfigurationManager.AppSettings("Subject"))
            smtp.Port = smtpPort
            'smtp.EnableSsl = True
            'smtp.

            'smtp.Credentials = New NetworkCredential("santhini.paramasivam@g2risksolutions.com", "Psant123!@#")
            'smtp.Timeout = 60000 ' Set timeout to 60 seconds


            Dim mail As New MailMessage()
            mail.From = New MailAddress(mailFromAddress)
            mail.To.Add(mailToAddress)
            mail.Subject = title & " " & batchControlID
            mail.Body = "The Batch id which is stuck on status 2 is: " & batchControlID & vbCrLf & vbCrLf & "The ClientCode is: " & ClientToProcess & vbCrLf & vbCrLf & "Please execute the below stored procedure to clean up the stuck batch: " & vbCrLf & vbCrLf & "EXEC USP_ReInsert_Failed_Batch_To_Batch_Control_Tbl " & batchControlID & vbCrLf & vbCrLf & "Note: Rerun the application after clean up."

            ' Send the email
            smtp.Send(mail)
            _logger.LogInfo("Email sent successfully!")


        Catch ex As Exception
            _logger.LogError("Error: " & ex.Message & " " & ex.StackTrace)


        End Try


    End Sub
#End Region
End Class
