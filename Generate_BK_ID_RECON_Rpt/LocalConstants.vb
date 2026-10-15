Public Class LocalConstants
    Public Const MY_MODULE = "BK_ID_Recon"
    Public Const INI_FILENAME = "c:\brs5\bin\settings.ini"

    Public Shared ReadOnly Property SQL_CONNECTION_STRING()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "SQL_CONNECTION_STRING", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find SQL_CONNECTION_STRING in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
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
                Const errorMessage = "Cannot find LOG_PATH in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property
    Public Shared ReadOnly Property NOTES_PATH()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "NOTES_PATH", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find NOTES_PATH in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property
    Public Shared ReadOnly Property PACKAGE_FOLDER()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "PACKAGE_FOLDER", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find PACKAGE_FOLDER in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property OUTPUT_FOLDER()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "OUTPUT_FOLDER", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find OUTPUT_FOLDER in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property DELIVERY_FOLDER()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "DELIVERY_FOLDER", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find DELIVERY_FOLDER in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property
    Public Shared ReadOnly Property ON_HOLD_FOLDER()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "ON_HOLD_FOLDER", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find ON_HOLD_FOLDER in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property
    Public Shared ReadOnly Property PACKAGE_NAME()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "PACKAGE_NAME", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find PACKAGE_NAME in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

End Class
