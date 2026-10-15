Public Class ExecuteAdHoc

    Public Const BATCH_TYPE_DAILY_GENERATION As Integer = 1
    Public Const BATCH_TYPE_MIGRATION As Integer = 2

    Private Shared _databaseControl As BMSL.DatabaseControl = Nothing
    Private Shared _logger As BMSL.LogControl = Nothing

    Private Shared _brs5Generator As BRS5.Generator.GenerateBRS5 = Nothing

    Shared Sub Main()
        Try
            _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)
            _databaseControl = New BMSL.DatabaseControl(_logger, LocalConstants.SQL_CONNECTION_STRING)

            _brs5Generator = New BRS5.Generator.GenerateBRS5(LocalConstants.LOG_PATH, LocalConstants.SQL_CONNECTION_STRING)

            _logger.LogInfo("Program starting up")

            _logger.LogInfo("Checking for migration ready jobs")
            Execute_AdHoc(BATCH_TYPE_MIGRATION)
            _logger.LogInfo("Checking for daily ready jobs")
            Execute_AdHoc(BATCH_TYPE_DAILY_GENERATION)
        Catch ex As Exception
            _logger.LogError("Exception was unrecoverable: " & ex.Message & " " & ex.StackTrace)
        End Try

    End Sub

    Public Shared Sub Execute_AdHoc(ByVal batchType As Integer)

        Dim batchControlID As Integer = 0

        'get next batch ID
        _logger.LogInfo("Selecting next job from Batch Control")

        Dim sqlString1 As String = "brs5.usp_Batch_Control__Select_Next_Ready"
        Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString1)

        sqlData1.AddParameter("@Client_ID", SqlDbType.Int, 4, 0)
        sqlData1.AddParameter("@File_ID", SqlDbType.Int, 4, 0)
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

        If LocalConstants.EXECUTE_GENERATION = "1" Then

            Try
                'start batch - update status
                _brs5Generator.UpdateBatchStatus(batchControlID, 2)
                _brs5Generator.InsertCurrentStatus(batchControlID, "Ad-Hoc Generation Starting")

                '-----------------------
                'GENERATION PROCESS START
                '-----------------------

                'Move previous BRS5 data to history, and clear staging tables 
                _brs5Generator.MovePreviousBRS5ToHistory_ClearPreviousStagingTables(batchControlID)

                'Insert migration cases
                _brs5Generator.InsertBatchCases_Migration(batchControlID)

                'Insert data from Cervello into BRS5 staging
                _brs5Generator.InsertCervelloIntoBRS5Staging(batchControlID)

                'Insert data from BRS5 staging into BRS5
                _brs5Generator.InsertBRS5StagingIntoBRS5(batchControlID, 1)

                '-----------------------
                'GENERATION PROCESS END
                '-----------------------

                _logger.LogInfo("Completed batch [" & batchControlID & "]")

                'finish batch - update status 
                _brs5Generator.InsertCurrentStatus(batchControlID, "Ad-Hoc Generation Complete")
                _brs5Generator.UpdateBatchStatus(batchControlID, 10)

            Catch ex As Exception
                'error batch - update status 
                _brs5Generator.UpdateBatchStatus(batchControlID, 99)
                _brs5Generator.InsertCurrentStatus(batchControlID, "Ad-Hoc Generation Error")

                'Altura-51
                'Update the next client batch to current batch start date for reprocessing date range.
                _brs5Generator.UpdateNextBatchDate(batchControlID)
                Throw ex

            End Try

        End If

        If LocalConstants.EXECUTE_OUTPUT = "1" Then

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
            Dim sqlString2 As String = "client.usp_Delivery_Options__Select_Summary"
            Dim sqlData2 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString2)

            Dim dsDataset1 As DataSet = sqlData2.SelectDataset

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

                    ' below is for option to copy the output txt files in an alternate folder
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

        End If

    End Sub
End Class
