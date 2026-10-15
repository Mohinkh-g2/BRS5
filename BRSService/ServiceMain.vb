Public Class ServiceMain

    Shared Sub Main()
        Try
            Dim sa As New ServiceLoop
            sa.StartService()
        Catch ex As Exception
            Console.WriteLine(ex.Message + ex.StackTrace)
            Threading.Thread.Sleep(10000)

        End Try

    End Sub


End Class
