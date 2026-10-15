Public Class DatabaseControl

    Dim _logger As LogControl = Nothing
    Dim _connectionString As String = ""
    Public Sub New(ByVal mLogger As LogControl, ByVal connectionString As String)
        Try
            _logger = mLogger
            _connectionString = connectionString

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Function CreateNewSession(ByVal sqlString As String) As DatabaseSession

        Try
            Return New DatabaseSession(_logger, _connectionString, sqlString)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function CreateNewSession(ByVal sqlString As String, ByVal timeout As Integer) As DatabaseSession

        Try
            Return New DatabaseSession(_logger, _connectionString, sqlString, timeout)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

End Class
