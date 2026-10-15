Public Class BMS_Log

    Private _filestub As String = ""
    Private _eventListener As Boolean = False

    ''' <summary>
    ''' Instantiates an LCI_Log.
    ''' </summary>
    ''' <param name="filestub">Folder and partial filename to write log file to.</param>
    ''' <param name="eventListener">True if should raise log events; otherwise, false.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub New(ByVal filestub As String, Optional ByVal eventListener As Boolean = False)
        _filestub = filestub
        _eventListener = eventListener

    End Sub

    ''' <summary>
    ''' Log a text message.
    ''' </summary>
    ''' <param name="message">Message to log.</param>
    ''' <param name="priority">Priority of log message.</param>
    ''' <remarks>Default priority 1.  Priority 1 message are written to console and file.  Priority 0 messages are written to file only.
    ''' <para>Author: Nick Washburn</para></remarks>
    Public Sub Log(ByVal message As String, Optional ByVal priority As Integer = 1)

        Dim textToWrite As String = ""

        textToWrite = Format(Now, "H:mm:ss.ff") & "|" & message

        If priority >= 1 Then WriteToConsole(textToWrite)
        If priority >= 0 Then WriteToTextFile(_filestub, textToWrite)

    End Sub

    'Write to Console event
    ''' <summary>
    ''' Log event raised when message logged.  
    ''' </summary>
    ''' <param name="logEntry">Message to raise in event.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Event evtLogOutput(ByVal logEntry As String)

    ''' <summary>
    ''' Writes a log message to the console.
    ''' </summary>
    ''' <param name="textToWrite">Text to log.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub WriteToConsole(ByVal textToWrite As String)
        System.Console.WriteLine()
        System.Console.Write(textToWrite)

        If _eventListener Then RaiseEvent evtLogOutput(textToWrite)
    End Sub

    ''' <summary>
    ''' Writes a log message to a text file.
    ''' </summary>
    ''' <param name="fileStub">Folder and partial filename to write log message to.</param>
    ''' <param name="textToWrite">Text to log.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub WriteToTextFile(ByVal fileStub As String, ByVal textToWrite As String)
        Try
            Dim objStreamWriter As IO.StreamWriter = Nothing
            Dim openedStream As Boolean = False
            Dim openAttempts As Integer = 0

            Dim filename As String = fileStub & Format(Now(), "yyyyMMdd") & ".txt"

            Do
                Try
                    If IO.File.Exists(filename) Then
                        'try opening file
                    End If
                    'Open new streamwriter
                    objStreamWriter = New IO.StreamWriter(filename, True)
                    openedStream = True
                Catch ex1 As UnauthorizedAccessException
                    openedStream = False
                    WriteToConsole("Access denied opening log file (" & ex1.Message & " [" & ex1.StackTrace & "])")
                    openAttempts = openAttempts + 5
                Catch ex As IO.IOException
                    openedStream = False
                    WriteToConsole("Retry opening log file")
                    Threading.Thread.Sleep(100)
                End Try
                openAttempts = openAttempts + 1
                If openAttempts > 5 Then Exit Do
            Loop Until openedStream = True

            If objStreamWriter Is Nothing Then
                WriteToConsole("WARNING: failed writing to log file (" & filename & ")")
            Else
                objStreamWriter.WriteLine(textToWrite)
                objStreamWriter.Close()
            End If

        Catch ex As Exception
            Throw New Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

End Class
