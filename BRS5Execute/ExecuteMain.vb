Public Class ExecuteMain

    Shared Sub Main()
        Try
            Dim exec As New Execute()

            ' exec.ExecuteDaily("BAL", 1, 1, 1)

            If LocalConstants.OUTPUT_ONLY = "1" Then
                exec.Execute_OutputOnly("BAL")

            Else
                If LocalConstants.EXISTING_BATCH = "1" Then
                    exec.Execute_ExistingBatch_RunDaily("BAL", 1, 1, 1)

                ElseIf LocalConstants.EXISTING_BATCH = "0" And LocalConstants.RUN_DAILY = "1" Then
                    exec.Execute_TimingCheck_CreateBatch_RunDaily("BAL", 1, 1, 1)

                ElseIf LocalConstants.EXISTING_BATCH = "0" And LocalConstants.RUN_DAILY = "0" And LocalConstants.RUN_MIGRATION = "1" Then

                    'exec.Execute_Migration("Q3G", 2, 1, 1)
                    Throw New Exception("Run migration not implemented")

                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message + ex.StackTrace)
            Threading.Thread.Sleep(10000)

        End Try

    End Sub

End Class
