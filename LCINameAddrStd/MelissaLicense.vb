Public Class MelissaLicense
    'Altura-12: Mellisa License string will be extracted from the settings.ini
    Public Const MY_MODULE = "LCINameAddrStd"
    Public Const INI_FILENAME = "c:\brs5\bin\settings.ini"

    Public Shared ReadOnly Property MELISSA_LICENSE_STRING()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "MELISSA_LICENSE_STRING", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find MELISSA_LICENSE_STRING in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property MELISSA_NAMEEXE_FILEPATH()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "MELISSA_NAMEEXE_FILEPATH", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find MELISSA_NAMEEXE_FILEPATH in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

    Public Shared ReadOnly Property MELISSA_ADDRESSEXE_FILEPATH()
        Get
            Dim iniValue As String
            Dim ini As New BMSL.INIFile(INI_FILENAME)

            iniValue = ini.GetString(MY_MODULE, "MELISSA_ADDRESSEXE_FILEPATH", "")
            If Len(Trim(iniValue)) > 0 Then
                Return iniValue
            Else
                Const errorMessage = "Cannot find MELISSA_ADDRESSEXE_FILEPATH in " & BMSL.Constants.INI_FILENAME & ", this is a required value."
                BMSL.LogControl.LogStaticMethodException(MY_MODULE, errorMessage, "")
                Threading.Thread.Sleep(5000)
                Throw New System.Exception(errorMessage)
            End If
        End Get
    End Property

End Class
