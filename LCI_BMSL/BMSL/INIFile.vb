Public Class INIFile

 
    ' API functions
    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal lpDefault As String, _
      ByVal lpReturnedString As System.Text.StringBuilder, _
      ByVal nSize As Integer, ByVal lpFileName As String) _
      As Integer
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal lpString As String, _
      ByVal lpFileName As String) As Integer
    Private Declare Ansi Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal nDefault As Integer, _
      ByVal lpFileName As String) As Integer
    Private Declare Ansi Function FlushPrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer, _
      ByVal lpKeyName As Integer, ByVal lpString As Integer, _
      ByVal lpFileName As String) As Integer

    Dim _filename As String

    ' Constructor, accepting a filename
    ''' <summary>
    ''' Instantiates an INIFile.
    ''' </summary>
    ''' <param name="Filename">Filename of INI file to read/write from.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub New(ByVal filename As String)

        _filename = filename

    End Sub

    ' Read-only filename property
    ''' <summary>
    ''' Returns value Filename.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Filename</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public ReadOnly Property pFilename() As String
        Get
            Return _filename
        End Get
    End Property

    ''' <summary>
    ''' Gets a string value from an INI file that corresponds to the section and key provided.
    ''' </summary>
    ''' <param name="Section">Section of INI file to read from.</param>
    ''' <param name="Key">Key of INI value.</param>
    ''' <param name="Default">Default value to return if requested value not found.</param>
    ''' <returns>String value from INI file.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function GetString(ByVal section As String, _
      ByVal key As String, ByVal [default] As String) As String
        Try
            ' Returns a string from your INI file
            Dim characterCount As Integer
            Dim strResult As New System.Text.StringBuilder(256)

            characterCount = GetPrivateProfileString(section, key, _
               [default], strResult, strResult.Capacity, _filename)

            If characterCount > 0 Then
                Return Left(strResult.ToString, characterCount)
            Else
                Return ""
            End If

        Catch ex As Exception
            LogControl.LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    ''' <summary>
    ''' Gets an integer value from an INI file that corresponds to the section and key provided.
    ''' </summary>
    ''' <param name="Section">Section of INI file to read from.</param>
    ''' <param name="Key">Key of INI value.</param>
    ''' <param name="Default">Default value to return if requested value not found.</param>
    ''' <returns>Integer value from INI file.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function GetInteger(ByVal section As String, _
      ByVal key As String, ByVal [default] As Integer) As Integer

        Try
            ' Returns an integer from your INI file
            Return GetPrivateProfileInt(section, key, _
                   [default], _filename)

        Catch ex As Exception
            LogControl.LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    ''' <summary>
    ''' Gets a boolean value from an INI file that corresponds to the section and key provided.
    ''' </summary>
    ''' <param name="Section">Section of INI file to read from.</param>
    ''' <param name="Key">Key of INI value.</param>
    ''' <param name="Default">Default value to return if requested value not found.</param>
    ''' <returns></returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function GetBoolean(ByVal section As String, _
      ByVal key As String, ByVal [default] As Boolean) As Boolean

        Try
            ' Returns a boolean from your INI file
            Return (GetPrivateProfileInt(section, key, _
               CInt([default]), _filename) = 1)

        Catch ex As Exception
            LogControl.LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    ''' <summary>
    ''' Writes a string value to an INI file that corresponds to the section and key provided.
    ''' </summary>
    ''' <param name="Section">Section of INI file to write to.</param>
    ''' <param name="Key">Key of INI value.</param>
    ''' <param name="Value">Value to write to INI file.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub WriteString(ByVal section As String, _
      ByVal key As String, ByVal value As String)

        Try
            ' Writes a string to your INI file
            WritePrivateProfileString(section, key, value, _filename)
            Flush()

        Catch ex As Exception
            LogControl.LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    ''' <summary>
    ''' Writes an INI value to an INI file that corresponds to the section and key provided.
    ''' </summary>
    ''' <param name="Section">Section of INI file to write to.</param>
    ''' <param name="Key">Key of INI value.</param>
    ''' <param name="Value">Value to write to INI file.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub WriteInteger(ByVal section As String, _
      ByVal key As String, ByVal value As Integer)

        Try
            ' Writes an integer to your INI file
            WriteString(section, key, CStr(value))
            Flush()

        Catch ex As Exception
            LogControl.LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    ''' <summary>
    ''' Writes a boolean value to an INI file that corresponds to the section and key provided.
    ''' </summary>
    ''' <param name="Section">Section of INI file to write to.</param>
    ''' <param name="Key">Key of INI value.</param>
    ''' <param name="Value">Value to write to INI file.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub WriteBoolean(ByVal section As String, _
      ByVal key As String, ByVal value As Boolean)

        Try
            ' Writes a boolean to your INI file
            WriteString(section, key, CStr(CInt(value)))
            Flush()

        Catch ex As Exception
            LogControl.LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    ''' <summary>
    ''' Flushes file cache.
    ''' </summary>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Sub Flush()

        Try
            ' Stores all the cached changes to your INI file
            FlushPrivateProfileString(0, 0, 0, _filename)

        Catch ex As Exception
            LogControl.LogStaticMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

End Class
