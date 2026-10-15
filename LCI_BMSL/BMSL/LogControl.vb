Public Class LogControl

    Dim _bms_Log As BMS_Log = Nothing

    Public ReadOnly Property pBMS_Log() As BMS_Log
        Get
            Return _bms_Log
        End Get
    End Property

    Public Sub New(ByVal filepath As String)

        Try
            _bms_Log = New BMS_Log(filepath)

        Catch ex As Exception
            LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)

        End Try

    End Sub

    Public Sub LogProgress(ByVal indicator As String)
        System.Console.Write(indicator)

    End Sub

    Public Sub LogDetail(ByVal message As String)
        Try
            _bms_Log.Log(message)

        Catch ex As Exception
            LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)

        End Try

    End Sub

    Public Sub LogInfo(ByVal message As String)

        Try
            _bms_Log.Log(message)

        Catch ex As Exception
            LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)

        End Try

    End Sub

    Public Sub LogWarning(ByVal message As String)
        Try
            _bms_Log.Log("WARNING: " & message)

        Catch ex As Exception
            LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)

        End Try

    End Sub

    Public Sub LogError(ByVal message As String)
        Try
            _bms_Log.Log("ERROR: " & message)

        Catch ex As Exception
            LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)

        End Try

    End Sub

    Public Sub LogMethodException(ByVal methodName As String, ByVal exception As String, ByVal stackTrace As String)
        'System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString()
        Try
            LogError("** EXCEPTION [" & methodName & "]: " & exception) '& " [" & stackTrace.Trim & "]")

        Catch ex As Exception
            LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Shared Sub LogStaticMethodException(ByVal methodName As String, ByVal exception As String, ByVal stackTrace As String)

        Console.WriteLine("** " & Format(Now, "H:mm:ss.ff") & "|" & "** EXCEPTION [" & methodName & "]: " & exception & " [" & stackTrace.Trim & "]")

    End Sub
End Class
