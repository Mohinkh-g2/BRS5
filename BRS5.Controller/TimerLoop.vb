Public Class TimerLoop

    Dim _logger As BMSL.LogControl = Nothing
    Dim _scheduledStartTime As String = ""

    Dim _nextStartTime As Date = Nothing
    Dim _timerLoopRunning As Boolean = False

    Public Sub New()
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)

        _scheduledStartTime = LocalConstants.SCHEDULED_START_TIME

        _logger.LogInfo("Scheduled start time: " & _scheduledStartTime)

    End Sub

    Dim runOnStartup As Boolean = False

    Public Sub StartTimerLoop()
        _logger.LogInfo("Starting timer loop")

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

        If LocalConstants.RUN_ON_STARTUP = "1" Then
            _logger.LogInfo("Run on startup command received")
            runOnStartup = True
        End If

        _timerLoopRunning = True
        RunTimerLoop()

    End Sub

    Public Sub StopTimerLoop()
        _logger.LogInfo("Stopping timer loop")
        _timerLoopRunning = False

    End Sub

    Private Sub RunTimerLoop()
        _logger.LogInfo("Running timer loop")

        While _timerLoopRunning
            Try
                If Now() > _nextStartTime Then
                    'start time passed
                    _logger.LogInfo("Scheduled start time passed")

                    'reschedule next run
                    _nextStartTime = DateAdd(DateInterval.Day, 1, _nextStartTime)
                    _logger.LogInfo("Next scheduled start time: " & _nextStartTime.ToString)

                    'execute BRS5 batch run
                    Dim batch1 As New DailyBatch
                    batch1.Execute()

                Else

                    If runOnStartup Then
                        _logger.LogInfo("Running on startup")
                        'execute BRS5 batch run
                        Dim batch1 As New DailyBatch
                        batch1.Execute()
                    End If

                    'start time not passed
                    _logger.LogInfo("Sleeping")

                End If


            Catch ex As Exception
                _logger.LogError("Caught exception while running timer loop: " & ex.Message & ex.StackTrace)
            Finally
                'sleep
                Threading.Thread.Sleep(45000)

                'reset flag
                runOnStartup = False

            End Try

        End While

        'timer loop stopped
        _logger.LogInfo("Timer loop stopped")


    End Sub

End Class
