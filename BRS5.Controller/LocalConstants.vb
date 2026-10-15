Public Class LocalConstants
    Public Const MY_MODULE = "BRS5Controller"
    Public Const INI_FILENAME = "c:\brs5\bin\settings.ini"

    Public Shared ReadOnly Property SQL_CONNECTION_STRING()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "SQL_CONNECTION_STRING", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find SQL_CONNECTION_STRING in " & INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property LOG_PATH()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "LOG_PATH", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find LOG_PATH in " & INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property EXECUTE_FILENAME()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "EXECUTE_FILENAME", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find EXECUTE_FILENAME in " & INI_FILENAME & ", this is a required value."
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
                Const errorMessage = "Cannot find SCHEDULED_START_TIME in " & INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property RUN_ON_STARTUP()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "RUN_ON_STARTUP", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Return "0"
                'Const errorMessage = "Cannot find RUN_ON_STARTUP in " & INI_FILENAME & ", this is a required value."
                'BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                'Threading.Thread.Sleep(5000)
                'Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property
End Class
