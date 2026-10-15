
Imports Microsoft.SqlServer.Dts.Runtime
Public Class SQLServerEventListener

    Inherits DefaultEvents
    'NLog
    Dim _logger As BMSL.LogControl = Nothing

    Public Overrides Function OnError(ByVal source As Microsoft.SqlServer.Dts.Runtime.DtsObject, _
      ByVal errorCode As Integer, ByVal subComponent As String, ByVal description As String, _
      ByVal helpFile As String, ByVal helpContext As Integer, _
      ByVal idofInterfaceWithError As String) As Boolean
        _logger = New BMSL.LogControl(LocalConstants.LOG_PATH)


        ' Add application–specific diagnostics here.
        _logger.loginfo("SSIS Error - SQLServerEventListener")
        Return False

    End Function
End Class
