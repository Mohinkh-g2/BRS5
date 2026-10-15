Public Class ServiceLoop

    Dim _logger As BMSL.LogControl = Nothing
    Dim _scheduledStartTime As String = ""

    Dim _nextStartTime As Date = Nothing
    Dim _serviceRunning As Boolean = False

    Public Sub New()
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)
        _logger.LogInfo("Program starting up")

        _scheduledStartTime = LocalConstants.SCHEDULED_START_TIME

        _logger.LogInfo("Scheduled start time: " & _scheduledStartTime)

        StartService()

    End Sub

    Public Sub StartService()
        _logger.LogInfo("Starting service")

        Dim buildNextStartTime As Date = Nothing
        Dim scheduledStartTimeHours As Integer = 0
        Dim scheduledStartTimeMinutes As Integer = 0

        scheduledStartTimeHours = Left(_scheduledStartTime, 2)
        scheduledStartTimeMinutes = Right(_scheduledStartTime, 2)

        buildNextStartTime = New Date(Now.Year, Now.Month, Now.Day, scheduledStartTimeHours, scheduledStartTimeMinutes, 1)

        If buildNextStartTime > Now() Then
            'in future today
            _nextStartTime = buildNextStartTime
        Else
            'in the past today; schedule for tomorrow
            _nextStartTime = DateAdd(DateInterval.Day, 1, buildNextStartTime)

        End If

        _logger.LogInfo("Next scheduled start time: " & _nextStartTime.ToString)


        _serviceRunning = True
        RunService()

    End Sub

    Public Sub StopService()
        _logger.LogInfo("Stopping service")
        _serviceRunning = False

    End Sub

    Private Sub RunService()

        While _serviceRunning

            If Now() > _nextStartTime Then
                'start time passed
                _logger.LogInfo("Scheduled start time passed")

                'reschedule next run
                _nextStartTime = DateAdd(DateInterval.Day, 1, _nextStartTime)
                _logger.LogInfo("Next scheduled start time: " & _nextStartTime.ToString)

                'run the scheduled event
                _logger.LogInfo("Starting external process")
                Dim p As Process = System.Diagnostics.Process.Start(LocalConstants.SERVICE_EXECUTE_FILENAME)

                'wait for it to finish a max 20 hours
                Dim result As Boolean = False
                result = p.WaitForExit(72000000)

                If result = False Then
                    'finished/timed out
                    _logger.LogInfo("Timed out waiting for external process to exit")
                Else
                    'finished
                    _logger.LogInfo("External process complete")

                End If
            Else
                'start time not passed
                _logger.LogInfo("Sleeping")

            End If

            'sleep
            Threading.Thread.Sleep(45000)
        End While

        'service stopped
        _logger.LogInfo("Service stopped")


    End Sub

End Class
