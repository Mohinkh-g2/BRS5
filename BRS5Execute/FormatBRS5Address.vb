Public Class FormatBRS5Address

    Dim _databaseControl As BMSL.DatabaseControl = Nothing
    Dim _logger As BMSL.LogControl = Nothing

    Dim _addrFormatter As LCINameAddrStd.AddrFormatter = Nothing

    Public Sub New()
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)
        _databaseControl = New BMSL.DatabaseControl(_logger, LocalConstants.SQL_CONNECTION_STRING)

        _addrFormatter = New LCINameAddrStd.AddrFormatter()

    End Sub

    Public Sub Execute()
        Dim logSummary As Boolean = False

        _logger.LogInfo("Process started")

        Dim dsDataset As DataSet = Nothing

        'DO LOOP get data (UNTIL ZERO ROWS RETURNED)
        Do
            _logger.LogInfo("Selecting dataset")
            dsDataset = SelectAddressBlock()

            _logger.LogInfo("Selected dataset: " & dsDataset.Tables(0).Rows.Count & " row(s) returned")

            If dsDataset.Tables(0).Rows.Count > 0 Then

                If dsDataset.Tables(0).Rows.Count > 1000 Then logSummary = True
                If logSummary Then _logger.LogInfo("Logging every 100th success entry")

                'FOR EACH LOOP format and update (UNTIL ALL PROCESSED)
                For Each dsRow As DataRow In dsDataset.Tables(0).Rows


                    Dim addressID As Integer = 0
                    Dim addressLine1 As String = ""
                    Dim addressLine2 As String = ""
                    Dim postalFormatCode As String = ""
                    Dim consumerAddressFormatCode As String = ""

                    If Not dsRow.IsNull("Staging_Debtor_Address_ID") Then addressID = dsRow.Item("Staging_Debtor_Address_ID")


                    'format
                    Dim dsRowResult As DataRow = FormatAddress(dsRow)

                    If logSummary Then
                        If addressID Mod 100 = 0 Then
                            _logger.LogInfo("Formatting addr: " & addressID)
                        Else
                            Console.WriteLine("Formatting addr: " & addressID)
                        End If
                    Else
                        _logger.LogInfo("Formatting addr: " & addressID)
                    End If

                    If Not dsRowResult.IsNull("Consumer_Address_Line_1") Then addressLine1 = dsRowResult.Item("Consumer_Address_Line_1")
                    If Not dsRowResult.IsNull("Consumer_Address_Line_2") Then addressLine2 = dsRowResult.Item("Consumer_Address_Line_2")
                    If Not dsRowResult.IsNull("Postal_Format_Code") Then postalFormatCode = dsRowResult.Item("Postal_Format_Code")
                    If Not dsRowResult.IsNull("Consumer_Address_Format_Code") Then consumerAddressFormatCode = dsRowResult.Item("Consumer_Address_Format_Code")

                    'update data
                    UpdateAddress(addressID, addressLine1, addressLine2, postalFormatCode & ":" & consumerAddressFormatCode, logSummary)

                Next dsRow
            End If

        Loop Until dsDataset.Tables(0).Rows.Count = 0

        _logger.LogInfo("Process complete")


    End Sub

    Public Function SelectAddressBlock() As DataSet

        Dim sqlString As String = "app.usp_Format_BRS5_Address__Select"
        Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)

        Dim dsDataset As DataSet = sqlData.SelectDataset

        Return dsDataset

    End Function

    Public Function FormatAddress(ByVal dsRow As DataRow) As DataRow
        Dim range As String = ""
        Dim preDirection As String = ""
        Dim streetName As String = ""
        Dim suffix As String = ""
        Dim postDirection As String = ""
        Dim suiteName As String = ""
        Dim suiteNumber As String = ""
        Dim boxName As String = ""
        Dim boxNumber As String = ""
        Dim city As String = ""
        Dim state As String = ""

        If Not dsRow.IsNull("Address_Range") Then range = dsRow.Item("Address_Range")
        If Not dsRow.IsNull("Address_Pre_Direction") Then preDirection = dsRow.Item("Address_Pre_Direction")
        If Not dsRow.IsNull("Address_Street_Name") Then streetName = dsRow.Item("Address_Street_Name")
        If Not dsRow.IsNull("Address_Suffix") Then suffix = dsRow.Item("Address_Suffix")
        If Not dsRow.IsNull("Address_Post_Direction") Then postDirection = dsRow.Item("Address_Post_Direction")
        If Not dsRow.IsNull("Address_Suite_Name") Then suiteName = dsRow.Item("Address_Suite_Name")
        If Not dsRow.IsNull("Address_Suite_Number") Then suiteNumber = dsRow.Item("Address_Suite_Number")
        If Not dsRow.IsNull("Address_Mailbox_Name") Then boxName = dsRow.Item("Address_Mailbox_Name")
        If Not dsRow.IsNull("Address_Mailbox_Number") Then boxNumber = dsRow.Item("Address_Mailbox_Number")
        If Not dsRow.IsNull("City") Then city = dsRow.Item("City")
        If Not dsRow.IsNull("State") Then state = dsRow.Item("State")

        Dim returnDS As DataSet = _addrFormatter.Format_AddrStd(range _
            , preDirection _
            , streetName _
            , suffix _
            , postDirection _
            , suiteName _
            , boxName _
            , boxNumber _
            , suiteNumber _
            , city _
            , state)

        Return returnDS.Tables(0).Rows(0)

    End Function

    Public Sub UpdateAddress(ByVal addressID As Integer, ByVal addressLine1 As String, ByVal addressLine2 As String, ByVal formatResult As String, ByVal logSummary As Boolean)
        Try
            Dim sqlString As String = "app.usp_Format_BRS5_Address__Update"
            Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)
            sqlData.AddParameter("@Address_ID", SqlDbType.Int, 4, addressID)
            sqlData.AddParameter("@Address_Line_1", SqlDbType.VarChar, 100, Left(addressLine1, 100))
            sqlData.AddParameter("@Address_Line_2", SqlDbType.VarChar, 100, Left(addressLine2, 100))
            sqlData.AddParameter("@Transacted_Result", SqlDbType.VarChar, 10, Left(formatResult, 10))

            Dim rowsAffected As Integer = sqlData.ExecuteNonQuery
            If logSummary Then
                If addressID Mod 100 = 0 Then
                    _logger.LogInfo("Updating " & addressID & ": " & rowsAffected & " row(s) affected")
                Else
                    Console.WriteLine("Updating " & addressID & ": " & rowsAffected & " row(s) affected")
                End If
            Else
                _logger.LogInfo("Updating " & addressID & ": " & rowsAffected & " row(s) affected")
            End If

        Catch ex As Exception
            _logger.LogError("Updating " & addressID & " FAILED: " & ex.Message)
            Throw ex
        End Try

    End Sub

End Class
