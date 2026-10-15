Public Class LocalConstants
    Public Const MY_MODULE = "BRS5Service"
    Public Const INI_FILENAME = "c:\brs5\bin\settings.ini"

    Public Shared ReadOnly Property LOG_PATH()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "LOG_PATH", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find LOG_PATH in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property SERVICE_EXECUTE_FILENAME()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "SERVICE_EXECUTE_FILENAME", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find SERVICE_EXECUTE_FILENAME in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property SCHEDULED_START_TIME()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "SCHEDULED_START_TIME", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find SCHEDULED_START_TIME in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

End Class
