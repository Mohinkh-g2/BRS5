Public Class DatabaseSession

    Dim _connectionString As String = ""
    Dim _logger As LogControl = Nothing

    Dim sqlData As BMS_SQLDataset = Nothing

    Public Sub New(ByVal mLogControl As LogControl, ByVal connectionString As String, ByVal sqlString As String)
        _connectionString = connectionString
        _logger = mLogControl

        sqlData = New BMS_SQLDataset(connectionString, sqlString)

    End Sub

    Public Sub New(ByVal mLogger As LogControl, ByVal connectionString As String, ByVal sqlString As String, ByVal timeout As Integer)
        Try
            _connectionString = connectionString
            _logger = mLogger

            sqlData = New BMS_SQLDataset(connectionString, sqlString, timeout)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Function AddParameter(ByVal name As String, ByVal type As SqlDbType, ByVal size As Integer, ByVal value As Object) As Boolean
        Try
            sqlData.AddParameter(name, type, size, value)
            Return True

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function SetOutputParameter(ByVal name As String) As Boolean

        Try
            sqlData.SetOutputParameter(name)
            Return True

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function ExecuteNonQuery() As Integer

        Try
            Return sqlData.ExecuteNonQuery()

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function ExecuteNonQuery(ByVal paramName As String) As String

        Try
            Return sqlData.ExecuteNonQueryGetParameter(paramName)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function SelectDataset() As DataSet

        'Try
        '    Return sqlData.SelectDataset

        'Catch ex As Exception
        '    _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
        '    Throw ex

        'End Try

        Return SelectDataset(1)

    End Function

    Public Function SelectDataset(ByVal retryCount As Integer) As DataSet

        Try
            Return sqlData.SelectDatasetCatchExceptions(_logger.pBMS_Log, retryCount)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function
End Class
