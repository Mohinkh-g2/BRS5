Public Class WebSession

    Private _ieWindowStack As Stack = Nothing
    Private _ieDocumentStack As Stack = Nothing

    Private _ieWindow As WebWindow = Nothing
    Private _ieDocument As WebDocument = Nothing

    Private _popupAcquired As Boolean = False

    Private _logger As LogControl = Nothing

    Public ReadOnly Property pIEWindow() As WebWindow
        Get
            Return _ieWindow
        End Get
    End Property

    Public ReadOnly Property pIEDocument() As WebDocument
        Get
            Return _ieDocument
        End Get
    End Property

    Public ReadOnly Property pPopupAcquired() As Boolean
        Get
            Return _popupAcquired
        End Get
    End Property

    Public Sub New(ByVal mLogger As LogControl)
        Try
            _ieWindowStack = New Stack
            _ieDocumentStack = New Stack

            _logger = mLogger

            _ieWindow = New WebWindow(_logger)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Function AcquirePopup(ByVal urlParts As String()) As Boolean

        If _popupAcquired = True Then Throw New HTMLInteractionException("Popup window is already acquired.")

        Dim mWebWindow As New WebWindow(_logger)

        If mWebWindow.AttachToPopupBrowser(urlParts, _ieWindow.pLocationURL) = True Then
            'Push current variables to stack
            PushIEDocument(_ieDocument)
            PushIEWindow(_ieWindow)

            'Reassign IE
            _ieWindow = mWebWindow

            'Get new document
            _ieDocument = _ieWindow.GetIEDocument()

            _popupAcquired = True

            Return True
        End If

        Return False


    End Function

    Public Sub DiscardPopup()

        If _popupAcquired = False Then Throw New HTMLInteractionException("Popup window is not acquired.")

        If _popupAcquired = True Then
            _popupAcquired = False

            _ieWindow.CloseBrowser()
            _ieWindow = Nothing
            _ieDocument = Nothing

            _ieWindow = PopIEWindow()
            _ieDocument = PopIEDocument()
        End If


    End Sub

    Public Sub Navigate(ByVal url As String)
        Try
            _ieWindow.NavigateBrowser(url)

            Const RETRY_MAX As Integer = 5
            Dim success As Integer = 0
            _logger.LogInfo("[WebSession] Navigating [" & url & "]")
            While success < RETRY_MAX
                Threading.Thread.Sleep(1500)
                success = success + 1
                Try
                    _ieDocument = _ieWindow.GetIEDocument()
                    success = RETRY_MAX + 1
                Catch ex As Exception
                    'try again
                    _logger.LogInfo("[WebSession] Waiting for IE document")
                End Try

            End While

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Sub Quit()
        Try
            _ieWindow.CloseBrowser()
            _ieWindow = Nothing
            _ieDocument = Nothing

            If _popupAcquired = True Then
                _popupAcquired = False

                _ieWindow = PopIEWindow()
                _ieDocument = PopIEDocument()

                _ieWindow.CloseBrowser()
                _ieWindow = Nothing
                _ieDocument = Nothing

            End If

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Sub PushIEWindow(ByVal window As WebWindow)
        Try
            _ieWindowStack.Push(window)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Function PopIEWindow() As WebWindow
        Try
            Return _ieWindowStack.Pop()

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Sub PushIEDocument(ByVal document As WebDocument)
        Try
            _ieDocumentStack.Push(document)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Function PopIEDocument() As WebDocument
        Try
            Return _ieDocumentStack.Pop

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

End Class
