Public Class BMS_SQLDataset
    Implements IDisposable

    Private _sql As String = ""
    Private _connection As Data.SqlClient.SqlConnection = Nothing
    Private _command As Data.SqlClient.SqlCommand = Nothing
    Private _dataAdapter As Data.SqlClient.SqlDataAdapter = Nothing
    Private _dataReader As Data.SqlClient.SqlDataReader = Nothing

    Private _cachedDataset As DataSet = Nothing

    ''' <summary>
    ''' Instantiates an LCI_SQLNDataset.
    ''' </summary>
    ''' <param name="connectionString">Connection string to connect to SQL database.</param>
    ''' <param name="sqlString">SQL statement to execute on SQL server.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub New(ByVal connectionString As String, ByVal sqlString As String) ', Optional ByVal commandTimeout As Integer = 600)
        _sql = sqlString

        Try
            _connection = New Data.SqlClient.SqlConnection(connectionString)
            _command = New Data.SqlClient.SqlCommand(_sql, _connection)
            If Not _sql.ToUpper.StartsWith("SELECT") _
            And Not _sql.ToUpper.StartsWith("INSERT") _
            And Not _sql.ToUpper.StartsWith("UPDATE") _
            And Not _sql.ToUpper.StartsWith("DELETE") Then
                _command.CommandType = CommandType.StoredProcedure
            End If
            _command.CommandTimeout = 12000 'commandTimeout

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)
        End Try

    End Sub

    ''' <summary>
    ''' Instantiates an LCI_SQLNDataset.
    ''' </summary>
    ''' <param name="connectionString">Connection string to connect to SQL database.</param>
    ''' <param name="sqlString">SQL statement to execute on SQL server.</param>
    ''' <param name="commandTimeout">Command timeout when executing on SQL server.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub New(ByVal connectionString As String, ByVal sqlString As String, ByVal commandTimeout As Integer)
        _sql = sqlString

        Try
            _connection = New Data.SqlClient.SqlConnection(connectionString)
            _command = New Data.SqlClient.SqlCommand(_sql, _connection)
            If Not _sql.ToUpper.StartsWith("SELECT ") _
            And Not _sql.ToUpper.StartsWith("INSERT ") _
            And Not _sql.ToUpper.StartsWith("UPDATE ") _
            And Not _sql.ToUpper.StartsWith("DELETE ") Then
                _command.CommandType = CommandType.StoredProcedure
            End If
            _command.CommandTimeout = commandTimeout

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)
        End Try

    End Sub

    ''' <summary>
    ''' Adds a parameter to SQL command, prior to execution.
    ''' </summary>
    ''' <param name="paramName">Parameter name to add.</param>
    ''' <param name="paramType">Parameter type to add. (SQL data type)</param>
    ''' <param name="paramSize">Parameter size to add.</param>
    ''' <param name="paramValue">Parameter value to add.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Overloads Sub AddParameter(ByVal paramName As String, ByVal paramType As SqlDbType, ByVal paramSize As Integer, ByVal paramValue As Object)
        Try
            _command.Parameters.Add(paramName, paramType, paramSize).Value = paramValue

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)
        End Try

    End Sub

    ''' <summary>
    ''' Sets a parameter type to output parameter, prior to execution.
    ''' </summary>
    ''' <param name="paramName">Parameter name to change to output type.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub SetOutputParameter(ByVal paramName As String)
        Try
            _command.Parameters(paramName).Direction = ParameterDirection.Output
        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)
        End Try

    End Sub

    ''' <summary>
    ''' Executes a SQL non-query on the SQL server.
    ''' </summary>
    ''' <returns>Number of rows affected on the SQL server by the non-query.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function ExecuteNonQuery() As Integer

        Dim rowsAffected As Integer = 0

        Try
            _dataAdapter = New Data.SqlClient.SqlDataAdapter
            _dataAdapter.SelectCommand = _command

            _connection.Open()
            rowsAffected = _command.ExecuteNonQuery
            _connection.Close()

            _command.Dispose()
            _command = Nothing
            _dataAdapter.Dispose()
            _dataAdapter = Nothing

            Return rowsAffected

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        Finally
            _connection.Close()

        End Try

    End Function

    ''' <summary>
    ''' Executes a SQL non-query on the SQL server and returns a parameter.
    ''' </summary>
    ''' <param name="paramName">SQL parameter that is being returned.</param>
    ''' <returns>Value of return parameter.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function ExecuteNonQueryGetParameter(ByVal paramName As String) As String

        Dim rowsAffected As Integer = 0
        Dim parameterValue As String = ""

        Try
            _dataAdapter = New Data.SqlClient.SqlDataAdapter
            _dataAdapter.SelectCommand = _command

            _connection.Open()
            rowsAffected = _command.ExecuteNonQuery
            parameterValue = _command.Parameters(paramName).Value

            _connection.Close()

            _command.Dispose()
            _command = Nothing
            _dataAdapter.Dispose()
            _dataAdapter = Nothing

            Return parameterValue

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        Finally
            _connection.Close()

        End Try

    End Function


    ''' <summary>
    ''' Selects a SQL dataset from the SQL server.
    ''' </summary>
    ''' <returns>SQL Dataset.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function SelectDataset() As DataSet

        Dim dsDataset As New DataSet

        Try
            _dataAdapter = New Data.SqlClient.SqlDataAdapter
            _dataAdapter.SelectCommand = _command

            _connection.Open()
            _dataAdapter.Fill(dsDataset)
            _connection.Close()

            _command.Dispose()
            _command = Nothing
            _dataAdapter.Dispose()
            _dataAdapter = Nothing

            Return dsDataset

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        Finally
            _connection.Close()

        End Try

    End Function

    ''' <summary>
    ''' Selects a SQL dataset from the SQL server.  Catches SQL exceptions and
    ''' retries SQL query, up to a specified number of times.
    ''' </summary>
    ''' <param name="logger">Logger passed in by caller.</param>
    ''' <param name="catchExceptionMax">Number of times to retry SQL query before throwing exception.</param>
    ''' <returns>SQL Dataset.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function SelectDatasetCatchExceptions(ByVal logger As BMS_Log, ByVal catchExceptionMax As Integer) As DataSet

        Dim dsDataset As New DataSet
        catchExceptionMax = catchExceptionMax + 1

        Try
            _dataAdapter = New Data.SqlClient.SqlDataAdapter
            _dataAdapter.SelectCommand = _command

            Dim exitLoop As Integer = 0
            Do Until exitLoop >= catchExceptionMax
                Try
                    _connection.Open()
                    _dataAdapter.Fill(dsDataset)

                    exitLoop = catchExceptionMax
                Catch ex As Exception
                    logger.Log("Caught exception while selecting dataset: " & ex.Message & ex.StackTrace)
                    exitLoop = exitLoop + 1
                    If exitLoop < catchExceptionMax Then
                        logger.Log("Retrying query")
                    Else
                        logger.Log("Query retry max reached")
                        Throw ex
                    End If

                Finally
                    _connection.Close()
                End Try
            Loop

            _command.Dispose()
            _command = Nothing
            _dataAdapter.Dispose()
            _dataAdapter = Nothing

            Return dsDataset

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        Finally
            _connection.Close()

        End Try

    End Function

    'Public editable As Boolean = False
    'Public Function selectDatasetEditable() As DataSet

    '    Dim dsDataset As New DataSet

    '    Try
    '        _dataAdapter = New SqlDataAdapter
    '        _dataAdapter.SelectCommand = _command
    '        _dataAdapter.UpdateCommand = _command

    '        _connection.Open()
    '        _dataAdapter.Fill(dsDataset)
    '        _connection.Close()

    '        editable = True
    '        '_command.Dispose()
    '        '_command = Nothing
    '        '_dataAdapter.Dispose()
    '        '_dataAdapter = Nothing

    '        Return dsDataset

    '    Catch ex As Exception
    '        Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

    '    Finally
    '        _connection.Close()

    '    End Try

    'End Function

    'Public Sub updateDataset(ByVal dsTable As DataTable)
    '    If editable = False Then
    '        Throw New System.Exception("** DLL EXCEPTION ** Dataset is not editable")
    '    Else
    '        _connection.Open()
    '        _dataAdapter.Update(dsTable)
    '        _connection.Close()

    '        editable = False

    '        _command.Dispose()
    '        _command = Nothing
    '        _dataAdapter.Dispose()
    '        _dataAdapter = Nothing

    '    End If
    'End Sub

    ''' <summary>
    ''' Selects a SQL dataset from the SQL server.
    ''' </summary>
    ''' <param name="tableName">Table name to return.</param>
    ''' <param name="maxRows">Maximum number of rows to return.</param>
    ''' <returns>SQL Dataset.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function SelectDataset(ByVal tableName As String, ByVal maxRows As Integer) As DataSet

        Dim dsDataset As New DataSet

        Try
            _dataAdapter = New Data.SqlClient.SqlDataAdapter
            _dataAdapter.SelectCommand = _command

            _connection.Open()
            _dataAdapter.Fill(dsDataset, 0, maxRows, tableName)
            _connection.Close()

            _command.Dispose()
            _command = Nothing
            _dataAdapter.Dispose()
            _dataAdapter = Nothing

            Return dsDataset

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        Finally
            _connection.Close()

        End Try

    End Function

    ''' <summary>
    ''' Disposes LCI_SQLNDataset by disposing data reader, adapter, command, and connection objects.
    ''' </summary>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub Dispose() Implements System.IDisposable.Dispose
        Try
            If Not _dataReader Is Nothing Then
                _dataReader.Close()
                _dataReader = Nothing
            End If
            If Not _dataAdapter Is Nothing Then
                _dataAdapter.Dispose()
                _dataAdapter = Nothing
            End If
            If Not _command Is Nothing Then
                _command.Dispose()
                _command = Nothing
            End If
            If Not _connection Is Nothing Then
                _connection.Close()
                _connection.Dispose()
                _connection = Nothing
            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    ''' <summary>
    ''' Executes a SQL non-query on the SQL server and returns a list of paramters.
    ''' </summary>
    ''' <returns>Value of return parameter.</returns>
    ''' <remarks>Author: Edward Ongsioco</remarks>
    Public Function ExecuteNonQueryGetParameter() As List(Of KeyValuePair(Of String, Object))

        Dim rowsAffected As Integer = 0
        Dim parameterValue As String = ""
        Dim outputParameters As New List(Of KeyValuePair(Of String, Object))
        Try
            _dataAdapter = New Data.SqlClient.SqlDataAdapter
            _dataAdapter.SelectCommand = _command

            _connection.Open()
            rowsAffected = _command.ExecuteNonQuery

            For Each param As SqlClient.SqlParameter In _command.Parameters
                If param.Direction <> ParameterDirection.Input Then

                    Dim parameterKvp As New KeyValuePair(Of String, Object)(param.ParameterName, param.Value)
                    outputParameters.Add(parameterKvp)
                End If
            Next

            _connection.Close()

            _command.Dispose()
            _command = Nothing
            _dataAdapter.Dispose()
            _dataAdapter = Nothing

            Return outputParameters

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        Finally
            _connection.Close()

        End Try

    End Function
End Class
