Public Class ControllerMain

    Private Shared _logger As BMSL.LogControl = Nothing

    Shared Sub Main()

        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)

        _logger.LogInfo("Program starting up")

        ' See if another instance is already running.
        If AlreadyRunning() Then
            _logger.LogWarning("There is already another instance of this program running.")
            _logger.LogWarning("This program will shut down in 10 seconds.")
            Threading.Thread.Sleep(10000)
        Else

            'create and start timer loop
            Dim timer1 As New TimerLoop
            timer1.StartTimerLoop()

        End If

        _logger.LogInfo("Program shutting down")

    End Sub

    ' Return True if another instance
    ' of this program is already running.
    Private Shared Function AlreadyRunning() As Boolean
        ' Get our process name.
        Dim my_proc As Process = Process.GetCurrentProcess
        Dim my_name As String = my_proc.ProcessName

        ' Get information about processes with this name.
        Dim procs() As Process = _
            Process.GetProcessesByName(my_name)

        ' If there is only one, it's us.
        If procs.Length = 1 Then Return False

        ' If there is more than one process,
        ' see if one has a StartTime before ours.
        Dim i As Integer
        For i = 0 To procs.Length - 1
            If procs(i).StartTime < my_proc.StartTime Then _
                Return True
        Next i

        ' If we get here, we were first.
        Return False
    End Function


End Class
