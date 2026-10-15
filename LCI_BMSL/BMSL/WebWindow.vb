Public Class WebWindow

    Private objMyLock As New Object
    Private objMyLock2 As New Object

    Private _logger As LogControl = Nothing
    Private _ie As SHDocVw.InternetExplorer = Nothing
    Private _ieHandlePtr As IntPtr

    Private evtBrowserReady As New Threading.AutoResetEvent(False)
    'Private _browserReadyCheckpoints As Collection
    'Private _unsolicitedBrowerReadyCheckpoints As Collection
    Private _readyDocumentCount As Integer = 0

    'Fill filebox
    Private Declare Auto Function PostMessage Lib "user32" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As UInt32, ByVal lParam As UInt32) As Boolean
    Private Declare Auto Function FindWindow Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Auto Function FindWindowEx Lib "user32" (ByVal hWnd1 As IntPtr, ByVal hWnd2 As IntPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As IntPtr

    Private Const WM_CHAR = &H102
    Const WM_GETTEXT As Long = &HD
    Const WM_GETTEXTLENGTH As Long = &HE
    Const GW_ENABLEDPOPUP As Long = 6
    Const GW_HWNDFIRST As Long = 0
    Const BM_CLICK As Long = &HF5&
    Const GW_CHILD As Long = 5
    Const GW_HWNDNEXT As Long = 2
    Const WM_KEYDOWN As Integer = &H100

    'Retrieve the popup window associated with the form, as well as to find the child windows of the popup...
    Private Declare Auto Function GetWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal uCmd As UInt32) As IntPtr
    'Sendmessage overload that is used to send messages to the button on the dialog window...
    Private Declare Auto Function SendMessage Lib "user32.dll" Alias "SendMessage" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByRef lParam As IntPtr) As IntPtr
    'Sendmessage overloads used to retrieve the window text...
    Private Declare Auto Function SendMessageA Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByRef lParam As IntPtr) As IntPtr
    <Runtime.InteropServices.DllImport("User32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto, Entrypoint:="SendMessage")> Public Shared Function SendMessageString(ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wparam As Integer, ByVal lparam As System.Text.StringBuilder) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=Runtime.InteropServices.CharSet.Auto)> Private Shared Function FindWindowByClass(ByVal lpClassName As String, ByVal zero As IntPtr) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=Runtime.InteropServices.CharSet.Auto)> Private Shared Function FindWindowByCaption(ByVal zero As IntPtr, ByVal lpWindowName As String) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetActiveWindow", SetLastError:=True, CharSet:=Runtime.InteropServices.CharSet.Auto)> Private Shared Function SetActiveWindow(ByVal hWnd As IntPtr) As IntPtr
    End Function

    'Public ReadOnly Property pIE() As SHDocVw.InternetExplorer
    '    Get
    '        Return _ie
    '    End Get
    'End Property

    Public Sub New(ByVal logger As LogControl)
        Try
            _logger = logger
            '_browserReadyCheckpoints = New Collection
            '_unsolicitedBrowerReadyCheckpoints = New Collection

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Function HWND_GetChildWindowHandles(ByVal parentWindowHandle As IntPtr) As ArrayList
        Try
            Dim ptrChild As IntPtr
            Dim clsRet As New ArrayList

            'get first child handle...
            ptrChild = HWND_GetChildWindowHandle(parentWindowHandle)

            Do Until ptrChild.Equals(IntPtr.Zero)
                'add to collection of handles...
                clsRet.Add(ptrChild)
                'get next child...
                ptrChild = HWND_GetNextWindowHandle(ptrChild)
            Loop

            'return...
            Return clsRet
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function HWND_GetChildWindowHandle(ByVal parentWindowHandle As IntPtr) As IntPtr
        Try
            Return GetWindow(parentWindowHandle, GW_CHILD)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function HWND_GetNextWindowHandle(ByVal currentWindowhandle As IntPtr) As IntPtr
        Try
            Return GetWindow(currentWindowhandle, GW_HWNDNEXT)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function HWND_GetWindowText(ByVal windowHandle As IntPtr) As String
        Try
            Dim ptrRet As IntPtr
            Dim ptrLength As IntPtr

            'get length for buffer...
            ptrLength = SendMessageA(windowHandle, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero)

            'create buffer for return value...
            Dim sb As New System.Text.StringBuilder(ptrLength.ToInt32 + 1)

            'get window text...
            ptrRet = SendMessageString(windowHandle, WM_GETTEXT, ptrLength.ToInt32 + 1, sb)

            'get return value...
            Return sb.ToString
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Sub HWND_ClickButton(ByVal windowHandle As IntPtr)
        Try
            SendMessage(windowHandle, BM_CLICK, 0, IntPtr.Zero)
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Sub

    Public Function HWND_FindWindowByCaption(ByVal windowText As String) As IntPtr
        Try
            Dim zeroPtr As IntPtr = 0
            Dim windowHWND As IntPtr = FindWindowByCaption(zeroPtr, windowText)

            Return windowHWND

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function HWND_DismissIEDialog() As Boolean
        Try
            Dim ptrDialogWindow As IntPtr = GetWindow(_ieHandlePtr, GW_ENABLEDPOPUP)

            'If the popup window is one displayed by the browser, then send the close message to the window...
            If HWND_GetWindowText(ptrDialogWindow) = "Microsoft Internet Explorer" Then
                Dim clsChildHandles As ArrayList = HWND_GetChildWindowHandles(ptrDialogWindow)

                Dim buildMessage As String = ""
                'look through all of the child handles of the window for an "OK" button (this method 
                'can also be used to gather more specific information about the dialog itself, such as 
                'the message being displayed)...
                For Each ptrHandle As IntPtr In clsChildHandles
                    buildMessage = "[" & HWND_GetWindowText(ptrHandle) & "] " & buildMessage
                Next

                _logger.LogInfo("[WebWindow] IE dialog message encountered: " & buildMessage)

                For Each ptrHandle As IntPtr In clsChildHandles
                    'if the OK button is found, click it...
                    If HWND_GetWindowText(ptrHandle) = "OK" Then
                        _logger.LogInfo("[WebWindow] Clicking [OK] on dialog message")
                        HWND_ClickButton(ptrHandle)
                        Return True

                    ElseIf HWND_GetWindowText(ptrHandle) = "&OK" Then
                        _logger.LogInfo("[WebWindow] Clicking [OK] on dialog message")
                        HWND_ClickButton(ptrHandle)
                        Return True

                    End If
                Next
            ElseIf HWND_GetWindowText(ptrDialogWindow) = "Security Alert" Then
                Dim clsChildHandles As ArrayList = HWND_GetChildWindowHandles(ptrDialogWindow)

                Dim buildMessage As String = ""
                'look through all of the child handles of the window for an "OK" button (this method 
                'can also be used to gather more specific information about the dialog itself, such as 
                'the message being displayed)...
                For Each ptrHandle As IntPtr In clsChildHandles
                    buildMessage = "[" & HWND_GetWindowText(ptrHandle) & "] " & buildMessage
                Next

                _logger.LogInfo("[WebWindow] IE dialog message encountered: " & buildMessage)

                For Each ptrHandle As IntPtr In clsChildHandles
                    'if the OK button is found, click it...
                    If HWND_GetWindowText(ptrHandle) = "Yes" Then
                        _logger.LogInfo("[WebWindow] Clicking [Yes] on dialog message")
                        HWND_ClickButton(ptrHandle)
                        Return True

                    ElseIf HWND_GetWindowText(ptrHandle) = "&Yes" Then
                        _logger.LogInfo("[WebWindow] Clicking [Yes] on dialog message")
                        HWND_ClickButton(ptrHandle)
                        Return True

                    End If
                Next

            ElseIf HWND_GetWindowText(ptrDialogWindow) = "Security Information" Then
                Dim clsChildHandles As ArrayList = HWND_GetChildWindowHandles(ptrDialogWindow)

                Dim buildMessage As String = ""
                'look through all of the child handles of the window for an "OK" button (this method 
                'can also be used to gather more specific information about the dialog itself, such as 
                'the message being displayed)...
                For Each ptrHandle As IntPtr In clsChildHandles
                    buildMessage = "[" & HWND_GetWindowText(ptrHandle) & "] " & buildMessage
                Next

                _logger.LogInfo("[WebWindow] IE dialog message encountered: " & buildMessage)

                For Each ptrHandle As IntPtr In clsChildHandles
                    'if the OK button is found, click it...
                    If HWND_GetWindowText(ptrHandle) = "Yes" Then
                        _logger.LogInfo("[WebWindow] Clicking [Yes] on dialog message")
                        HWND_ClickButton(ptrHandle)
                        Return True

                    ElseIf HWND_GetWindowText(ptrHandle) = "&Yes" Then
                        _logger.LogInfo("[WebWindow] Clicking [Yes] on dialog message")
                        HWND_ClickButton(ptrHandle)
                        Return True

                    End If
                Next

            End If

            Return False
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    'USE THIS to emulate ClickSaveDialogChild
    Public Function HWND_GetPopupWindowHandle(ByVal parentWindowHandle As IntPtr) As IntPtr
        Try
            Return GetWindow(parentWindowHandle, GW_ENABLEDPOPUP)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function HWND_ClickButtonWithText(ByVal buttonText As String, ByVal parentWindowHandle As IntPtr) As Boolean
        Try
            Dim clsChildHandles As ArrayList = HWND_GetChildWindowHandles(parentWindowHandle)

            Dim buildMessage As String = ""
            'look through all of the child handles of the window for an "OK" button (this method 
            'can also be used to gather more specific information about the dialog itself, such as 
            'the message being displayed)...
            For Each ptrHandle As IntPtr In clsChildHandles
                buildMessage = "[" & HWND_GetWindowText(ptrHandle) & "] " & buildMessage
            Next

            _logger.LogInfo("[WebWindow] IE dialog message encountered: " & buildMessage)

            For Each ptrHandle As IntPtr In clsChildHandles
                'if the OK button is found, click it...
                If HWND_GetWindowText(ptrHandle) = buttonText Then
                    _logger.LogInfo("[WebWindow] Clicking [" & buttonText & "] on dialog message")
                    HWND_ClickButton(ptrHandle)
                    Return True
                ElseIf HWND_GetWindowText(ptrHandle) = "&" & buttonText Then
                    _logger.LogInfo("[WebWindow] Clicking [" & buttonText & "] on dialog message")
                    HWND_ClickButton(ptrHandle)
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function HWND_ClickButtonWithText_RetrieveFilename(ByVal buttonText As String, ByVal parentWindowHandle As IntPtr) As String
        Try
            Dim clsChildHandles As ArrayList = HWND_GetChildWindowHandles(parentWindowHandle)

            Dim buildMessage As String = ""

            Dim filename As String = ""
            Dim getNextMessage As Boolean = False

            'look through all of the child handles of the window for an "OK" button (this method 
            'can also be used to gather more specific information about the dialog itself, such as 
            'the message being displayed)...
            For Each ptrHandle As IntPtr In clsChildHandles
                buildMessage = buildMessage & "[" & HWND_GetWindowText(ptrHandle) & "] "

                If getNextMessage = True Then
                    filename = HWND_GetWindowText(ptrHandle)
                    _logger.LogInfo("[WebWindow] Retrieved filename: " & filename)
                    getNextMessage = False
                End If
                If HWND_GetWindowText(ptrHandle) = "Name:" Then getNextMessage = True
            Next

            _logger.LogInfo("[WebWindow] IE dialog message encountered: " & buildMessage)

            For Each ptrHandle As IntPtr In clsChildHandles
                'if the OK button is found, click it...
                If HWND_GetWindowText(ptrHandle) = buttonText Then
                    _logger.LogInfo("[WebWindow] Clicking [" & buttonText & "] on dialog message")
                    HWND_ClickButton(ptrHandle)
                    Return filename
                ElseIf HWND_GetWindowText(ptrHandle) = "&" & buttonText Then
                    _logger.LogInfo("[WebWindow] Clicking [" & buttonText & "] on dialog message")
                    HWND_ClickButton(ptrHandle)
                    Return filename
                End If
            Next

            Return filename
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function HWND_TypeText(ByVal text As String) As Boolean

        Dim shellDocObjectView As IntPtr
        Dim internetExplorerServer As IntPtr

        shellDocObjectView = FindWindowEx(_ieHandlePtr, Nothing, "shell docobject view", vbNullString)
        internetExplorerServer = FindWindowEx(shellDocObjectView, Nothing, "internet explorer_server", vbNullString)

        Dim messageResult As Boolean = Nothing

        If Len(text) > 0 Then
            Dim i As Integer = 1
            While i <= Len(text)
                Dim c As String = ""
                c = Mid(text, i, 1)
                messageResult = PostMessage(internetExplorerServer, WM_CHAR, Asc(c), 0)
                i = i + 1
            End While
        End If

        Return True
    End Function

    Public Function HWND_KeyEnter()
        Try
            Dim shellDocObjectView As IntPtr
            Dim internetExplorerServer As IntPtr

            shellDocObjectView = FindWindowEx(_ieHandlePtr, Nothing, "shell docobject view", vbNullString)
            internetExplorerServer = FindWindowEx(shellDocObjectView, Nothing, "internet explorer_server", vbNullString)
            PostMessage(internetExplorerServer, WM_KEYDOWN, 13, 0)

            Return True
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function HWND_KeyDown()
        Try
            Dim shellDocObjectView As IntPtr
            Dim internetExplorerServer As IntPtr

            shellDocObjectView = FindWindowEx(_ieHandlePtr, Nothing, "shell docobject view", vbNullString)
            internetExplorerServer = FindWindowEx(shellDocObjectView, Nothing, "internet explorer_server", vbNullString)
            PostMessage(internetExplorerServer, WM_KEYDOWN, 40, 0)

            Return True
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Private Sub InitializeBrowser()
        Try
            _logger.LogInfo("[WebWindow] Initializing browser")

            If Not _ie Is Nothing Then Throw New HTMLInteractionException("IE object is already initialized")

            _ie = New SHDocVw.InternetExplorer
            '_ie.Silent = True
            _ie.Visible = True
            _ieHandlePtr = _ie.HWND
            AddHandler _ie.DocumentComplete, AddressOf evtHandler_DocumentComplete

            Threading.Thread.Sleep(1500)
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Sub

    Public Function AttachToExistingBrowser(ByVal windowHandle As Integer) As Boolean
        Try
            Dim allBrowsers As New SHDocVw.ShellWindows

            _logger.LogInfo("[WebWindow] Attaching to existing browser [" & windowHandle & "]")

            If Not _ie Is Nothing Then Throw New HTMLInteractionException("IE object is already initialized")

            For i As Integer = 0 To (allBrowsers.Count)
                Dim e As SHDocVw.InternetExplorer

                e = allBrowsers.Item(i)
                If Not e Is Nothing Then
                    _logger.LogInfo("[WebWindow] Analyzing browser [" & e.HWND & "]")

                    If e.HWND = windowHandle Then 'p.MainWindowHandle.ToInt32 Then
                        _logger.LogInfo("[WebWindow] Attaching to browser [" & e.HWND & "]")

                        _ie = e
                        '_ie.Silent = True
                        _ieHandlePtr = _ie.HWND 'p.MainWindowHandle
                        AddHandler _ie.DocumentComplete, AddressOf evtHandler_DocumentComplete
                        Return True
                        Exit For
                    End If
                End If

            Next i

            _logger.LogInfo("[WebWindow] Could not find browser with handle [" & windowHandle & "]")
            Return False

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public ReadOnly Property pLocationURL() As String
        Get
            Return _ie.LocationURL
        End Get
    End Property

    Public Function AttachToPopupBrowser(ByVal urlParts As String(), ByVal existingLocationURL As String) As Boolean

        Try
            _logger.LogInfo("[WebWindow] Attaching to popup browser")

            Dim thisURL As String = existingLocationURL '_ie.LocationURL
            Dim browserCount As Integer = 0

            Dim allBrowsers As New SHDocVw.ShellWindows
            browserCount = allBrowsers.Count

            If browserCount = 0 Then
                Throw New BMSL.HTMLInteractionException("No browser windows found")
            ElseIf browserCount = 1 Then
                Throw New BMSL.HTMLInteractionException("No popup windows found")
            Else

                For i As Integer = 0 To browserCount - 1

                    Dim e As SHDocVw.InternetExplorer 'WebBrowser 
                    e = allBrowsers.Item(i)
                    If Not e Is Nothing Then
                        _logger.LogInfo("[WebWindow] Analyzing browser [" & e.HWND & "]")

                        'Check to make sure it's not the current URL
                        If e.LocationURL.Trim.ToLower = thisURL.Trim.ToLower Then
                            'Disregard
                        Else
                            'Check locationURL for parts of base URL
                            Dim totalComponents As Integer = urlParts.Length
                            Dim foundComponents As Integer = 0

                            For j As Integer = 0 To totalComponents - 1
                                If e.LocationURL.Trim.ToLower.Contains(urlParts(j).Trim.ToLower) Then
                                    foundComponents += 1

                                End If

                            Next j

                            If foundComponents = totalComponents Then
                                _logger.LogInfo("[WebWindow] Attaching to browser [" & e.HWND & "]")

                                'Found all needed components
                                'This is the right window
                                _ie = e
                                '_ie.Silent = True
                                _ieHandlePtr = _ie.HWND 'p.MainWindowHandle
                                AddHandler _ie.DocumentComplete, AddressOf evtHandler_DocumentComplete
                                Return True
                                Exit For

                            End If
                        End If

                    End If
                Next i
            End If

            _logger.LogInfo("[WebWindow] Could not find popup browser")
            Return False

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Sub NavigateBrowser(ByVal navigateURL As String)
        Try
            _logger.LogInfo("[WebWindow] Navigating browser")

            If _ie Is Nothing Then
                InitializeBrowser()
            End If

            _ie.Navigate(navigateURL, Nothing, Nothing, Nothing, Nothing)
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Sub

    Public Sub NavigateBrowserBack()
        Try
            _logger.LogInfo("[WebWindow] Navigating browser back")

            If _ie Is Nothing Then
                Throw New System.Exception("Browser is not initialized.")
            End If

            _ie.GoBack()
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Sub

    Public Sub CloseBrowser()
        Try
            _logger.LogInfo("[WebWindow] Closing browser")

            _ie.Quit()

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Sub

    Public Function GetIEDocument() As WebDocument
        Try
            Return New WebDocument(_ie.Document, _logger)
        Catch ex As Exception
            '_logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    'Event handler
    Public ReadOnly Property pReadyDocuments() As Integer
        Get
            Return _readyDocumentCount
        End Get
    End Property

    Public Sub ClearReadyDocuments()
        _readyDocumentCount = 0
    End Sub

    Private Sub evtHandler_DocumentComplete(ByVal pDisp As Object, ByRef URL As Object)
        SyncLock objMyLock2
            Try
                _logger.LogInfo("[IE] Document complete")
                Threading.Interlocked.Increment(_readyDocumentCount)
                Me.evtBrowserReady.Set()

                'Indicate that a document complete event occurred
                'UpdateNextBrowserReadyCheckpoint()

            Catch ex As Exception
                _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
                Throw ex

            End Try
        End SyncLock

    End Sub

    Public Function BrowserReadyWait(Optional ByVal timeout As Integer = 60000) As Boolean
        SyncLock objMyLock
            Try
                Threading.Thread.Sleep(100)
                _logger.LogInfo("[WebWindow] Waiting for browser ready state")
                Dim success As Boolean = Me.evtBrowserReady.WaitOne(timeout, False)
                If success = False Then _logger.LogInfo("[IE] The wait request timed out")

                Return success
            Catch ex As Exception
                _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
                Throw ex

            End Try
        End SyncLock
    End Function

    'Public Sub AddBrowserReadyCheckpoint(ByVal id As String)
    '    SyncLock objMyLock
    '        Try
    '            If _browserReadyCheckpoints.Contains(id) Then Throw New HTMLInteractionException("BrowserReady [" & id & "] already exists")

    '            _logger.LogInfo("[WebWindow] BrowserReady [" & id & "] queued")

    '            'Add zero - not checked off yet
    '            _browserReadyCheckpoints.Add(New WebCheckpoint(id), id)

    '        Catch ex As Exception
    '            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '            Throw ex

    '        End Try
    '    End SyncLock
    'End Sub


    'Public Sub UpdateNextBrowserReadyCheckpoint() 'Optional ByVal tryAgain As Boolean = True)
    '    SyncLock objMyLock

    '    Try
    '            Dim found As Boolean = False

    '            'Update next checkpoint that is not checked off yet
    '            For Each checkpoint As WebCheckpoint In _browserReadyCheckpoints
    '                If checkpoint.Checked = False Then
    '                    checkpoint.Checked = True
    '                    found = True

    '                    _logger.LogInfo("[WebWindow] BrowserReady [" & checkpoint.ID & "] cleared")
    '                    Threading.Thread.Sleep(500)

    '                    Exit For
    '                End If
    '            Next checkpoint

    '            If found = False Then
    '                _logger.LogWarning("WARNING: BrowserReady does not contain an available checkpoint.")
    '                Me._unsolicitedBrowerReadyCheckpoints.Add(New WebCheckpoint("Error" & Format(Now(), "yyyyMMddhhmmssffff")))

    '                Threading.Thread.Sleep(2000)
    '                Throw New HTMLInteractionException("BrowserReady does not contain an available checkpoint.")
    '            End If

    '        Catch ex As Exception
    '            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '            Throw ex

    '        End Try
    '    End SyncLock

    'End Sub

    'Public Function GetBrowserReadyCheckpoint(ByVal id As String) As WebCheckpoint
    '    SyncLock objMyLock
    '        Try
    '            If Not _browserReadyCheckpoints.Contains(id) Then Throw New HTMLInteractionException("BrowserReady does not contain this checkpoint.  [" & id & "]")

    '            Return _browserReadyCheckpoints.Item(id)
    '        Catch ex As Exception
    '            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '            Throw ex

    '        End Try
    '    End SyncLock
    'End Function

    'Public Sub RemoveBrowserReadyCheckpoint(ByVal id As String)
    '    SyncLock objMyLock
    '        Try
    '            If Not _browserReadyCheckpoints.Contains(id) Then Throw New HTMLInteractionException("BrowserReady does not contain this checkpoint.  [" & id & "]")

    '            If _browserReadyCheckpoints.Contains(id) Then _browserReadyCheckpoints.Remove(id)
    '        Catch ex As Exception
    '            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '            Throw ex

    '        End Try
    '    End SyncLock
    'End Sub

    Public Function BrowserBusy() As Boolean
        Return _ie.Busy
    End Function

End Class
