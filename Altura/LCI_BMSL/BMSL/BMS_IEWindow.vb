Public Class hideBMS_IEWindow

    Protected Const NO_FRAME = ""

    Private _logger As BMS_Log = Nothing
    Private _ie As SHDocVw.InternetExplorer = Nothing
    Private _ieHandlePtr As IntPtr
    Private _ieSaveDialogHandlePtr As IntPtr
    Private _iePopupTimer As System.Timers.Timer

    'Private _pid As Integer

    'Private _documentComplete As New Threading.AutoResetEvent(False)
    'Private _browserNonReadyCount As Integer = 0

    'NEW browser ready logic
    Private evtBrowserReady As New Threading.AutoResetEvent(False)

    'Fill filebox
    Private Declare Auto Function PostMessage Lib "user32" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As UInt32, ByVal lParam As UInt32) As Boolean
    Private Declare Auto Function FindWindow Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Auto Function FindWindowEx Lib "user32" (ByVal hWnd1 As IntPtr, ByVal hWnd2 As IntPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As IntPtr
    Private Const WM_CHAR = &H102

    'Close IE popups
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

    Public Sub New(ByVal logger As BMS_Log)
        Try
            _logger = logger

            _iePopupTimer = New System.Timers.Timer(10000)
            AddHandler _iePopupTimer.Elapsed, AddressOf popup_TimerFired
            _iePopupTimer.Enabled = True

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Enumerate the Windows handles of all Windows sited on the specified parent...
    ''' <summary>
    ''' Popup handling function.  Enumerate the Windows handles of all Windows sited on the specified parent.
    ''' </summary>
    ''' <param name="ParentWindowHandle">Integer-pointer of parent window handle.</param>
    ''' <returns>Array list of child window handles.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Function popup_GetChildWindowHandles(ByVal ParentWindowHandle As IntPtr) As ArrayList
        Try

            Dim ptrChild As IntPtr
            Dim clsRet As New ArrayList

            'get first child handle...
            ptrChild = popup_GetChildWindowHandle(ParentWindowHandle)

            Do Until ptrChild.Equals(IntPtr.Zero)
                'add to collection of handles...
                clsRet.Add(ptrChild)
                'get next child...
                ptrChild = popup_GetNextWindowHandle(ptrChild)
            Loop

            'return...
            Return clsRet
        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Enumerate the Windows handles of all Windows sited on the specified parent...
    ''' <summary>
    ''' Popup handling function.  Enumerate the Windows handles of all Windows sited on the specified parent.
    ''' </summary>
    ''' <param name="ParentWindowHandle">Integer-pointer of parent window handle.</param>
    ''' <returns>Integer-pointer of child window handle.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Function popup_GetChildWindowHandle(ByVal ParentWindowHandle As IntPtr) As IntPtr
        Return GetWindow(ParentWindowHandle, GW_CHILD)
    End Function

    'Enumerate the Windows handles of all Windows sited on the specified parent...
    ''' <summary>
    ''' Popup handling function.  Enumerate the Windows handles of all Windows sited on the specified parent.
    ''' </summary>
    ''' <param name="CurrentWindowhandle">Integer-pointer of current window handle.</param>
    ''' <returns>Integer-pointer of next window handle.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Function popup_GetNextWindowHandle(ByVal CurrentWindowhandle As IntPtr) As IntPtr
        Return GetWindow(CurrentWindowhandle, GW_HWNDNEXT)
    End Function

    'Returns the text of the window, used so that we can confirm that we have the right dialog window...
    ''' <summary>
    ''' Popup handling function.  Returns the text of the window, used so that we can confirm that we have the right dialog window.
    ''' </summary>
    ''' <param name="WindowHandle">Integer-pointer of window handle for which to get text.</param>
    ''' <returns>Window text of specified window.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Function popup_GetWindowText(ByVal WindowHandle As IntPtr) As String
        Try
            Dim ptrRet As IntPtr
            Dim ptrLength As IntPtr

            'get length for buffer...
            ptrLength = SendMessageA(WindowHandle, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero)

            'create buffer for return value...
            Dim sb As New System.Text.StringBuilder(ptrLength.ToInt32 + 1)

            'get window text...
            ptrRet = SendMessageString(WindowHandle, WM_GETTEXT, ptrLength.ToInt32 + 1, sb)

            'get return value...
            Return sb.ToString

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Sends a button click message to the specified window...
    ''' <summary>
    ''' Popup handling function.  Sends a button click message to the specified window.
    ''' </summary>
    ''' <param name="WindowHandle">Integer-pointer of window handle for which to press button.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Sub popup_ClickButton(ByVal WindowHandle As IntPtr)
        Try
            SendMessage(WindowHandle, BM_CLICK, 0, IntPtr.Zero)

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Handle Timer Fired event for timer
    ''' <summary>
    ''' Popup handling function.  Handles TimerFired event.  Checks for popup message in the browser on a specifed interval.
    ''' </summary>
    ''' <param name="sender">Timer sender object.</param>
    ''' <param name="e">Timer event args.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Sub popup_TimerFired(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        Try
            If Not _ie Is Nothing Then
                Dim ptrDialogWindow As IntPtr = GetWindow(_ieHandlePtr, GW_ENABLEDPOPUP)

                'If the popup window is one displayed by the browser, then send the close message to the window...
                If popup_GetWindowText(ptrDialogWindow) = "Microsoft Internet Explorer" Then
                    Dim clsChildHandles As ArrayList = popup_GetChildWindowHandles(ptrDialogWindow)

                    Dim buildMessage As String = ""
                    'look through all of the child handles of the window for an "OK" button (this method 
                    'can also be used to gather more specific information about the dialog itself, such as 
                    'the message being displayed)...
                    For Each ptrHandle As IntPtr In clsChildHandles
                        buildMessage = "[" & popup_GetWindowText(ptrHandle) & "] " & buildMessage
                    Next

                    _logger.Log("IE dialog message encountered [popup_TimerFired]: " & buildMessage)

                    For Each ptrHandle As IntPtr In clsChildHandles
                        'if the OK button is found, click it...
                        If popup_GetWindowText(ptrHandle) = "OK" Then
                            _logger.Log("Clicking [OK] on dialog message")
                            popup_ClickButton(ptrHandle)
                            Exit For
                        ElseIf popup_GetWindowText(ptrHandle) = "&OK" Then
                            _logger.Log("Clicking [OK] on dialog message")
                            popup_ClickButton(ptrHandle)
                            Exit For
                        End If
                    Next
                ElseIf popup_GetWindowText(ptrDialogWindow) = "Security Alert" Then
                    Dim clsChildHandles As ArrayList = popup_GetChildWindowHandles(ptrDialogWindow)

                    Dim buildMessage As String = ""
                    'look through all of the child handles of the window for an "OK" button (this method 
                    'can also be used to gather more specific information about the dialog itself, such as 
                    'the message being displayed)...
                    For Each ptrHandle As IntPtr In clsChildHandles
                        buildMessage = "[" & popup_GetWindowText(ptrHandle) & "] " & buildMessage
                    Next

                    _logger.Log("IE dialog message encountered [popup_TimerFired]: " & buildMessage)

                    For Each ptrHandle As IntPtr In clsChildHandles
                        'if the OK button is found, click it...
                        If popup_GetWindowText(ptrHandle) = "Yes" Then
                            _logger.Log("Clicking [Yes] on dialog message")
                            popup_ClickButton(ptrHandle)
                            Exit For
                        ElseIf popup_GetWindowText(ptrHandle) = "&Yes" Then
                            _logger.Log("Clicking [Yes] on dialog message")
                            popup_ClickButton(ptrHandle)
                            Exit For
                        End If
                    Next

                End If

            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    ''' <summary>
    ''' Popup handling function.  Finds save dialog box that contains the specified window text.
    ''' </summary>
    ''' <param name="windowText">Window text that is contained on window to find.</param>
    ''' <returns>Integer-pointer of dialog box to find.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function popup_FindSaveDialogHWND(ByVal windowText As String) As IntPtr
        Dim zeroPtr As IntPtr = 0
        Dim windowHWND As IntPtr = FindWindowByCaption(zeroPtr, windowText)

        Return windowHWND
    End Function

    ''' <summary>
    ''' Popup handling function.  Clicks button on save dialog box parent.
    ''' </summary>
    ''' <param name="buttonText">Text on button to click.</param>
    ''' <param name="ptrParentWindow">Integer-pointer of parent window.</param>
    ''' <returns>True if button was clicked; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function popup_ClickSaveDialogParent(ByVal buttonText As String, ByVal ptrParentWindow As IntPtr) As Boolean
        Dim clsChildHandles As ArrayList = popup_GetChildWindowHandles(ptrParentWindow)

        Dim buildMessage As String = ""
        'look through all of the child handles of the window for an "OK" button (this method 
        'can also be used to gather more specific information about the dialog itself, such as 
        'the message being displayed)...
        For Each ptrHandle As IntPtr In clsChildHandles
            buildMessage = "[" & popup_GetWindowText(ptrHandle) & "] " & buildMessage
        Next

        _logger.Log("IE dialog message encountered [popup_ClickSaveDialogParent]: " & buildMessage)

        For Each ptrHandle As IntPtr In clsChildHandles
            'if the OK button is found, click it...
            If popup_GetWindowText(ptrHandle) = buttonText Then
                _logger.Log("Clicking [" & buttonText & "] on dialog message")
                popup_ClickButton(ptrHandle)
                Return True
            ElseIf popup_GetWindowText(ptrHandle) = "&" & buttonText Then
                _logger.Log("Clicking [" & buttonText & "] on dialog message")
                popup_ClickButton(ptrHandle)
                Return True
            End If
        Next
        Return False
    End Function

    Public Function popup_ClickSaveDialogParent_GetFilename(ByVal buttonText As String, ByVal ptrParentWindow As IntPtr) As String

        Dim clsChildHandles As ArrayList = popup_GetChildWindowHandles(ptrParentWindow)

        Dim buildMessage As String = ""

        Dim filename As String = ""
        Dim getNextMessage As Boolean = False

        'look through all of the child handles of the window for an "OK" button (this method 
        'can also be used to gather more specific information about the dialog itself, such as 
        'the message being displayed)...
        For Each ptrHandle As IntPtr In clsChildHandles
            buildMessage = buildMessage & "[" & popup_GetWindowText(ptrHandle) & "] "

            If getNextMessage = True Then
                filename = popup_GetWindowText(ptrHandle)
                _logger.Log("Retrieved filename: " & filename)
                getNextMessage = False
            End If
            If popup_GetWindowText(ptrHandle) = "Name:" Then getNextMessage = True
        Next

        _logger.Log("IE dialog message encountered [popup_ClickSaveDialogParent]: " & buildMessage)

        For Each ptrHandle As IntPtr In clsChildHandles
            'if the OK button is found, click it...
            If popup_GetWindowText(ptrHandle) = buttonText Then
                _logger.Log("Clicking [" & buttonText & "] on dialog message")
                popup_ClickButton(ptrHandle)
                Return filename
            ElseIf popup_GetWindowText(ptrHandle) = "&" & buttonText Then
                _logger.Log("Clicking [" & buttonText & "] on dialog message")
                popup_ClickButton(ptrHandle)
                Return filename
            End If
        Next

        Return filename

    End Function

    ''' <summary>
    ''' Popup handling function.  Clicks button on save dialog box child.
    ''' </summary>
    ''' <param name="windowText">Window text of dialog.</param>
    ''' <param name="buttonText">Text on button to click.</param>
    ''' <param name="ptrParentWindow">Integer-pointer of parent window.</param>
    ''' <returns>True if button was clicked; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function popup_ClickSaveDialogChild(ByVal windowText As String, ByVal buttonText As String, ByVal ptrParentWindow As IntPtr)

        Dim ptrDialogWindow As IntPtr = GetWindow(ptrParentWindow, GW_ENABLEDPOPUP)

        'If the popup window is one displayed by the browser, then send the close message to the window...
        If popup_GetWindowText(ptrDialogWindow) = windowText Then
            Dim clsChildHandles As ArrayList = popup_GetChildWindowHandles(ptrDialogWindow)

            Dim buildMessage As String = ""
            'look through all of the child handles of the window for an "OK" button (this method 
            'can also be used to gather more specific information about the dialog itself, such as 
            'the message being displayed)...
            For Each ptrHandle As IntPtr In clsChildHandles
                buildMessage = "[" & popup_GetWindowText(ptrHandle) & "] " & buildMessage
            Next

            _logger.Log("IE dialog message encountered [popup_ClickSaveDialogChild]: " & buildMessage)

            For Each ptrHandle As IntPtr In clsChildHandles
                'if the OK button is found, click it...
                If popup_GetWindowText(ptrHandle) = buttonText Then
                    _logger.Log("Clicking [" & buttonText & "] on dialog message")
                    popup_ClickButton(ptrHandle)
                    Return True
                ElseIf popup_GetWindowText(ptrHandle) = "&" & buttonText Then
                    _logger.Log("Clicking [" & buttonText & "] on dialog message")
                    popup_ClickButton(ptrHandle)
                    Return True
                End If
            Next

        End If

        Return False
    End Function

    Public Function popup_ClickSaveDialogChild_GetFilename(ByVal windowText As String, ByVal buttonText As String, ByVal ptrParentWindow As IntPtr) As String

        Dim ptrDialogWindow As IntPtr = GetWindow(ptrParentWindow, GW_ENABLEDPOPUP)

        Dim filename As String = ""
        Dim getNextMessage As Boolean = False

        'If the popup window is one displayed by the browser, then send the close message to the window...
        If popup_GetWindowText(ptrDialogWindow) = windowText Then
            Dim clsChildHandles As ArrayList = popup_GetChildWindowHandles(ptrDialogWindow)

            Dim buildMessage As String = ""

            'look through all of the child handles of the window for an "OK" button (this method 
            'can also be used to gather more specific information about the dialog itself, such as 
            'the message being displayed)...
            For Each ptrHandle As IntPtr In clsChildHandles
                buildMessage = buildMessage & "[" & popup_GetWindowText(ptrHandle) & "] "

                If getNextMessage = True Then
                    filename = popup_GetWindowText(ptrHandle)
                    _logger.Log("Retrieved filename: " & filename)
                    getNextMessage = False
                End If
                If popup_GetWindowText(ptrHandle) = "Name:" Then getNextMessage = True
            Next

            _logger.Log("IE dialog message encountered [popup_ClickSaveDialogChild]: " & buildMessage)

            For Each ptrHandle As IntPtr In clsChildHandles
                'if the OK button is found, click it...
                If popup_GetWindowText(ptrHandle) = buttonText Then
                    _logger.Log("Clicking [" & buttonText & "] on dialog message")
                    popup_ClickButton(ptrHandle)
                    Return filename
                ElseIf popup_GetWindowText(ptrHandle) = "&" & buttonText Then
                    _logger.Log("Clicking [" & buttonText & "] on dialog message")
                    popup_ClickButton(ptrHandle)
                    Return filename
                End If
            Next

        End If

        Return filename
    End Function

    ''' <summary>
    ''' Popup handling function.  Clicks button containing text specified on message box specified.
    ''' </summary>
    ''' <param name="windowText">Window text of message to click button on.</param>
    ''' <param name="buttonText">Text of button to click.</param>
    ''' <returns>True if message box was found and button clicked; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function popup_ClearMessage(ByVal windowText As String, ByVal buttonText As String) As Boolean
        Try
            If Not _ie Is Nothing Then
                Dim ptrDialogWindow As IntPtr = GetWindow(_ieHandlePtr, GW_ENABLEDPOPUP)

                'If the popup window is one displayed by the browser, then send the close message to the window...
                If popup_GetWindowText(ptrDialogWindow) = windowText Then
                    Dim clsChildHandles As ArrayList = popup_GetChildWindowHandles(ptrDialogWindow)

                    Dim buildMessage As String = ""
                    'look through all of the child handles of the window for an "OK" button (this method 
                    'can also be used to gather more specific information about the dialog itself, such as 
                    'the message being displayed)...
                    For Each ptrHandle As IntPtr In clsChildHandles
                        buildMessage = "[" & popup_GetWindowText(ptrHandle) & "] " & buildMessage
                    Next

                    _logger.Log("IE dialog message encountered [popup_ClearMessage]: " & buildMessage)

                    For Each ptrHandle As IntPtr In clsChildHandles
                        'if the OK button is found, click it...
                        If popup_GetWindowText(ptrHandle) = buttonText Then
                            _logger.Log("Clicking [" & buttonText & "] on dialog message")
                            popup_ClickButton(ptrHandle)
                            Return True
                        ElseIf popup_GetWindowText(ptrHandle) = "&" & buttonText Then
                            _logger.Log("Clicking [" & buttonText & "] on dialog message")
                            popup_ClickButton(ptrHandle)
                            Return True
                        End If
                    Next

                End If

            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try
        Return False

    End Function

    ''' <summary>
    ''' Finds and kills processes containing partial name 'IEXPLORE'.
    ''' </summary>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Sub KillBrowserWindows()
        Dim processData() As Process = Process.GetProcesses

        For Each p As Process In processData
            If p.ProcessName.ToUpper.Contains("IEXPLORE") Then
                Try
                    _logger.Log("[IE] Process killed [" & p.MainWindowTitle & "]")
                    p.Kill()
                Catch ex As Exception
                    _logger.Log("[IE] Failed to kill browser session [" & p.MainWindowTitle & "]: " & ex.Message)
                End Try

            End If
        Next p

    End Sub

    'Initialize browser
    ''' <summary>
    ''' Initializes and attaches to a new Internet Explorer process Threading.Thread.
    ''' </summary>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Sub InitializeBrowser()

        Dim p As Process = Nothing

        Try
            Dim retryOnce As Boolean = False
            Dim skipRemainingSections As Boolean = False
            Dim retryCount As Integer = 0

            Do
                skipRemainingSections = False

                _logger.Log("[IE] Initializing browser")

                'Initialize browser
                p = Process.Start("iexplore.exe", "about:blank")
                '_pid = p.Id

                If p Is Nothing Then
                    If retryOnce = False Then
                        'Retry this loop once
                        retryOnce = True
                        skipRemainingSections = True
                    Else
                        KillBrowserWindows()
                        Throw New HTMLInteractionException("Error initializing browser - cannot launch iexplore.exe")
                    End If
                End If

                If skipRemainingSections = False Then
                    'Don't run if retryOnce is set

                    Threading.Thread.Sleep(2000)
                    'Console.WriteLine("Launched browser (" & p.MainWindowHandle.ToInt32 & ")")

                    'Attach _ie object to browser with the same handle
                    Dim allBrowsers As New SHDocVw.ShellWindows
                    If allBrowsers.Count = 0 Then
                        If retryOnce = False Then
                            'Retry this loop once
                            retryOnce = True
                            skipRemainingSections = True
                        Else
                            p.Kill()
                            KillBrowserWindows()
                            Throw New HTMLInteractionException("Error initializing browser - no browser sessions can be found")
                        End If
                    End If

                    If skipRemainingSections = False Then
                        'Don't run if retryOnce is set

                        For i As Integer = 0 To (allBrowsers.Count)
                            Dim e As SHDocVw.InternetExplorer 'WebBrowser 
                            e = allBrowsers.Item(i)
                            If Not e Is Nothing Then
                                If e.HWND = p.MainWindowHandle.ToInt32 Then
                                    _ie = e
                                    _ie.Silent = True
                                    _ieHandlePtr = p.MainWindowHandle
                                    AddHandler _ie.DocumentComplete, AddressOf ie_DocumentComplete

                                    Exit For
                                End If
                            End If

                        Next i

                        If _ie Is Nothing Then
                            retryCount = retryCount + 1
                            If retryCount >= 5 Then
                                If retryOnce = False Then
                                    'Retry this loop once
                                    retryOnce = True
                                    skipRemainingSections = True
                                Else
                                    p.Kill()
                                    KillBrowserWindows()
                                    Throw New HTMLInteractionException("Error initializing browser - could not attach to iexplore.exe session")
                                End If
                            End If

                            _logger.Log("Could not attach to iexplore.exe session")
                            Threading.Thread.Sleep(2000)
                            'Throw New Exception("Error initializing browser - could not attach to iexplore.exe session")
                        End If
                    End If
                End If

            Loop Until Not _ie Is Nothing

        Catch ex As Exception
            Try
                If Not p Is Nothing Then p.Kill()
                KillBrowserWindows()
            Catch ex1 As Exception
                _logger.Log("Kill Browser Windows failed: " & ex1.Message & " [" & Trim(ex1.StackTrace) & "]")
            End Try

            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try


    End Sub

    'Initialize browser
    ''' <summary>
    ''' Initializes and connects to a new instance of Internet Explorer.
    ''' </summary>
    ''' <remarks>Author: Nick Washburn</remarks>
    Private Sub InitializeBrowser2()

        Try

            _logger.Log("[IE] Initializing browser")

            _ie = New SHDocVw.InternetExplorer
            '_ie.Silent = True
            _ie.Visible = True
            _ieHandlePtr = _ie.HWND
            AddHandler _ie.DocumentComplete, AddressOf ie_DocumentComplete

            Threading.Thread.Sleep(2000)

        Catch ex As Exception

            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try


    End Sub

    'Navigate to a new URL in Browser window
    ''' <summary>
    ''' Navigate to a new URL in Browser window.
    ''' </summary>
    ''' <param name="navigateURL">URL to navigate to.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub NavigateBrowser(ByVal navigateURL As String)
        Try
            _logger.Log("[IE] Navigating browser")

            If _ie Is Nothing Then
                InitializeBrowser()
            End If

            '_browserNonReadyCount = 1

            _ie.Navigate(navigateURL)

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Navigate to a new URL in Browser window
    ''' <summary>
    ''' Navigate to a new URL in Browser window.
    ''' </summary>
    ''' <param name="navigateURL">URL to navigate to.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub NavigateBrowser2(ByVal navigateURL As String)
        Try
            _logger.Log("[IE] Navigating browser")

            If _ie Is Nothing Then
                InitializeBrowser2()
            End If

            '_browserNonReadyCount = 1

            _ie.Navigate(navigateURL, Nothing, Nothing, Nothing, Nothing)

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Navigate to a new URL in Browser window
    ''' <summary>
    ''' Navigates browser window to previous URL.
    ''' </summary>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub NavigateBrowserBack()
        Try
            _logger.Log("[IE] Navigating browser back")

            If _ie Is Nothing Then
                Throw New System.Exception("Browser is not initialized.")
            End If

            '_browserNonReadyCount = 1

            _ie.GoBack()

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Public Sub IncreaseNonReadyCount()
    '    Try
    '        _logger.Log("Increasing browser wait count by 1")
    '        _browserNonReadyCount = Threading.Interlocked.Increment(_browserNonReadyCount)

    '    Catch ex As Exception
    '        Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

    '    End Try

    'End Sub

    Public Function BrowserReadyWait(Optional ByVal timeout As Integer = 60000) As Boolean

        _logger.Log("[IE] Waiting for browser ready state")
        Dim success As Boolean = Me.evtBrowserReady.WaitOne(timeout, False)
        If success = False Then _logger.Log("[IE] The wait request timed out")

        Return success

    End Function

    'Public Sub BrowserWait(Optional ByVal timeout As Integer = 60000)
    '    Try
    '        _logger.Log("[IE] Waiting for browser ready state")
    '        Dim success As Boolean = False

    '        If _browserNonReadyCount <= 0 Then
    '            _logger.Log("[IE] Browser is already ready")
    '            success = _documentComplete.WaitOne(1000, False)

    '        Else
    '            success = _documentComplete.WaitOne(timeout, False)
    '            If success = False Then
    '                'Throw New Exception("Error waiting for browser - the browse request timed out.")
    '                _logger.Log("WARNING - The wait request timed out, browser may not be ready")
    '            End If
    '        End If

    '    Catch ex As Exception
    '        Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

    '    End Try

    'End Sub

    'Public Sub BrowserClearWaitQueue(Optional ByVal numberToClear As Integer = 1)
    '     Try
    '         Dim x As Integer = 0
    '         Do Until x >= numberToClear
    '             _documentComplete.WaitOne(100, False)
    '             x = x + 1
    '         Loop

    '     Catch ex As Exception
    '         Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

    '     End Try

    ' End Sub

    'Close browser window
    ''' <summary>
    ''' Closes Internet Explorer browser.
    ''' </summary>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub CloseBrowser()
        Try
            _logger.Log("[IE] Closing browser")

            _ie.Quit()

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Isolate an HTML element and return to calling function
    ''' <summary>
    ''' Isolates an HTML element and returns to caller.
    ''' </summary>
    ''' <param name="elementName">HTML element to isolate.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>HTML element as IHTMLElement.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    ''' 
    Public Function IsolateHTMLElement(ByVal elementName As String, ByVal myDocument As WebDocument) As mshtml.IHTMLElement ' ByVal frameName As String, Optional ByVal subframeName As String = "") As IHTMLElement
        Try
            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing

            document = myDocument.pHTMLDocument

            Dim elementCollection As mshtml.IHTMLElementCollection = document.getElementsByName(elementName)
            Dim performSecondaryMatch As Boolean = False

            If elementCollection.length = 0 Then
                performSecondaryMatch = True
            ElseIf elementCollection.length > 1 Then
                Throw New System.Exception("More than one element (" & elementCollection.length & ") were found with the name [" & elementName & "]")
            End If

            'Do secondary matching on input elements
            If performSecondaryMatch = True Then
                Dim elementCollectionInputElements As mshtml.IHTMLElementCollection = document.getElementsByTagName("INPUT")
                For Each element As mshtml.IHTMLElement In elementCollectionInputElements
                    If Trim(element.innerText) = elementName Then
                        Return element
                    End If
                    If Trim(element.value) = elementName Then
                        Return element
                    End If
                Next
                Throw New System.Exception("No elements were found with the name [" & elementName & "]")
            End If

            Return elementCollection.item(0, Nothing)

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    Public Function IsolateHTMLElementByID(ByVal elementID As String, ByVal myDocument As WebDocument) As mshtml.IHTMLElement ' ByVal frameName As String, Optional ByVal subframeName As String = "") As IHTMLElement
        Try
            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing

            document = myDocument.pHTMLDocument

            Dim element As mshtml.IHTMLElement = document.getElementById(elementID)

            Return element

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Search for an HTML element and return true or false
    ''' <summary>
    ''' Searches for an HTML element on a page.
    ''' </summary>
    ''' <param name="elementName">HTML element to search for.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>True if the HTML element is found; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function SearchHTMLElement(ByVal elementName As String, ByVal myDocument As WebDocument) As Boolean ', ByVal frameName As String, Optional ByVal subframeName As String = "") As Boolean
        Try
            Dim itemFound As Boolean = False

            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing

            document = myDocument.pHTMLDocument

            Dim elementCollection As mshtml.IHTMLElementCollection = document.getElementsByName(elementName)
            Dim performSecondaryMatch As Boolean = False

            If elementCollection.length = 0 Then
                performSecondaryMatch = True
            ElseIf elementCollection.length > 1 Then
                'Throw New System.Exception("More than one element (" & elementCollection.length & ") were found with the name [" & elementName & "]")
                itemFound = True
            Else
                itemFound = True
            End If

            'Do secondary matching on input & anchor elements
            If performSecondaryMatch = True Then
                Dim elementCollectionInputElements As mshtml.IHTMLElementCollection = document.getElementsByTagName("INPUT")
                For Each element As mshtml.IHTMLElement In elementCollectionInputElements
                    If Trim(element.innerText) = elementName Then
                        itemFound = True
                        'Return element
                    End If
                    If Trim(element.value) = elementName Then
                        'Return element
                        itemFound = True
                    End If
                Next

                Dim elementCollectionAnchors As mshtml.IHTMLElementCollection = document.getElementsByTagName("A")
                For Each element As mshtml.IHTMLElement In elementCollectionAnchors
                    If Trim(element.innerText) = elementName Then
                        itemFound = True
                        'Return element
                    End If
                Next
                'Throw New System.Exception("No elements were found with the name [" & elementName & "]")
            End If

            Return itemFound

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Fill IE text box
    ''' <summary>
    ''' Fills an Internet Explorer text box with specified text.
    ''' </summary>
    ''' <param name="textBoxName">Text box to fill.</param>
    ''' <param name="textBoxValue">Value to put in text box.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub FillTextBox(ByVal textBoxName As String, ByVal textBoxValue As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Entering text into [" & textBoxName & "]")

            Dim element As mshtml.IHTMLElement = Nothing

            element = IsolateHTMLElement(textBoxName, myDocument)

            Dim textBox As mshtml.HTMLInputElement = element
            textBox.value = textBoxValue

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Fill IE text box
    ''' <summary>
    ''' Fills an Internet Explorer text box with specified text.
    ''' </summary>
    ''' <param name="textBoxName">Text box to fill.</param>
    ''' <param name="textBoxValue">Value to put in text box.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub FillTextBox2(ByVal textBoxName As String, ByVal textBoxValue As String, ByVal myDocument As WebDocument)  ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Entering text into [" & textBoxName & "]")

            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing

            document = myDocument.pHTMLDocument

            Dim iHTMLCol As mshtml.IHTMLElementCollection
            Dim iHTMLEle As mshtml.IHTMLElement
            Dim str As String

            iHTMLCol = document.getElementsByTagName("input")

            Dim foundTextBox As Boolean = False
            For Each iHTMLEle In iHTMLCol
                If Not iHTMLEle.getAttribute("name") Is Nothing Then
                    str = iHTMLEle.getAttribute("name").ToString
                    If str = textBoxName Then
                        iHTMLEle.setAttribute("value", textBoxValue)
                        foundTextBox = True
                        Exit For
                    End If
                End If
            Next

            If foundTextBox = False Then Throw New System.Exception("No textbox was found with name [" & textBoxName & "]")

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Fill IE file box
    ''' <summary>
    ''' Fills an Internet Explorer file text box with specified filename.
    ''' </summary>
    ''' <param name="textFilename">Filename to insert into file text box.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub FillFileTextBox(ByVal textFilename As String) 'ByVal fileBoxName As String, ByVal textFilename As String, ByVal frameName As String)
        Try
            _logger.Log("[IE] Entering filename")

            'Dim ieFrame As IntPtr
            Dim shellDocObjectView As IntPtr
            Dim InternetExplorerServer As IntPtr

            'ieFrame = FindWindow("ieframe", vbNullString)
            'shellDocObjectView = FindWindowEx(ieFrame, Nothing, "shell docobject view", vbNullString)
            shellDocObjectView = FindWindowEx(_ieHandlePtr, Nothing, "shell docobject view", vbNullString)
            InternetExplorerServer = FindWindowEx(shellDocObjectView, Nothing, "internet explorer_server", vbNullString)

            Dim messageResult As Boolean = Nothing

            If Len(textFilename) > 0 Then
                Dim i As Integer = 1
                While i <= Len(textFilename)
                    Dim c As String = ""
                    c = Mid(textFilename, i, 1)
                    messageResult = PostMessage(InternetExplorerServer, WM_CHAR, Asc(c), 0)
                    i = i + 1
                End While
            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    ''' <summary>
    ''' Fires javascript OnChange event for specified text box.
    ''' </summary>
    ''' <param name="textBoxName">Text box name to fire event for.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub FireOnChange_TextBox(ByVal textBoxName As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        _logger.Log("[IE] Firing OnChange event for text box [" & textBoxName & "]")

        Dim element As mshtml.IHTMLElement = Nothing
        'If Len(subframeName) > 0 Then
        '    element = isolateHTMLElement(textBoxName, frameName, subframeName)
        'Else
        '    element = isolateHTMLElement(textBoxName, frameName)
        'End If
        element = IsolateHTMLElement(textBoxName, myDocument)

        Dim textBox As mshtml.HTMLInputElement = element

        textBox.focus()
        textBox.FireEvent("onchange", textBox)

    End Sub

    Public Sub FireOnKeyDown_TextBox(ByVal textBoxName As String, ByVal myDocument As WebDocument)

        'NW NEW CODE
        _logger.Log("[IE] Firing OnKeyDown event for text box [" & textBoxName & "]")
        Dim element As mshtml.IHTMLElement = IsolateHTMLElement(textBoxName, myDocument)
        Dim textBox As mshtml.HTMLInputElement = element
        textBox.focus()
        textBox.FireEvent("onkeydown", textBox)

    End Sub

    Public Sub FireOnClick_TableRow(ByVal rowID As String, ByVal myDocument As WebDocument)

        'NW NEW CODE
        _logger.Log("[IE] Firing OnClick event for table row [" & rowID & "]")
        Dim element As mshtml.IHTMLElement = IsolateHTMLElementByID(rowID, myDocument)
        Dim tableRow As mshtml.HTMLTableRowClass = element
        tableRow.focus()
        tableRow.FireEvent("onclick")

    End Sub

    Public Sub ClickEnter()
        Try
            _logger.Log("[IE] Click Enter")
            'Dim ieFrame As IntPtr
            Dim shellDocObjectView As IntPtr
            Dim internetExplorerServer As IntPtr

            'ieFrame = FindWindow("ieframe", vbNullString)
            'shellDocObjectView = FindWindowEx(ieFrame, Nothing, "shell docobject view", vbNullString)
            shellDocObjectView = FindWindowEx(_ieHandlePtr, Nothing, "shell docobject view", vbNullString)
            internetExplorerServer = FindWindowEx(shellDocObjectView, Nothing, "internet explorer_server", vbNullString)
            PostMessage(internetExplorerServer, WM_KEYDOWN, 13, 0)



        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try
    End Sub

    ''' <summary>
    ''' Fires javascript OnChange event for specified text box.
    ''' </summary>
    ''' <param name="textBoxName">Text box name to fire event for.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub FireOnBlur_TextBox(ByVal textBoxName As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        _logger.Log("[IE] Firing Blur event for text box [" & textBoxName & "]")

        Dim element As mshtml.IHTMLElement = Nothing
        'If Len(subframeName) > 0 Then
        '    element = isolateHTMLElement(textBoxName, frameName, subframeName)
        'Else
        '    element = isolateHTMLElement(textBoxName, frameName)
        'End If
        element = IsolateHTMLElement(textBoxName, myDocument)

        Dim textBox As mshtml.HTMLInputElement = element

        textBox.focus()
        textBox.FireEvent("onblur", textBox)

    End Sub

    'Click IE button
    ''' <summary>
    ''' Clicks specified button on web page.
    ''' </summary>
    ''' <param name="buttonName">Name of button to click.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub ClickButton(ByVal buttonName As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Clicking button [" & buttonName & "]")

            'Dim element As IHTMLElement = isolateHTMLElement(buttonName, frameName)
            Dim element As mshtml.IHTMLElement = Nothing

            element = IsolateHTMLElement(buttonName, myDocument)

            Dim button As mshtml.HTMLInputElement = element
            '_browserNonReadyCount = 1

            button.click()

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try
    End Sub


    'Click button
    ''' <summary>
    ''' Clicks specified button on web page.
    ''' </summary>
    ''' <param name="buttonOuterHTML">Outer HTML of button to click.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub ClickButton2(ByVal buttonOuterHTML As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Clicking button [" & buttonOuterHTML & "]")


            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing

            document = myDocument.pHTMLDocument

            Dim iHTMLCol As mshtml.IHTMLElementCollection
            Dim iHTMLEle As mshtml.IHTMLElement

            iHTMLCol = document.getElementsByTagName("input")


            Dim foundButton As Boolean = False
            ' Press the submit button
            For Each iHTMLEle In iHTMLCol
                'If Not iHTMLEle.getAttribute("name") Is Nothing Then
                'MsgBox(iHTMLEle.outerHTML)
                If iHTMLEle.outerHTML = buttonOuterHTML Then
                    '"<INPUT tabIndex=4 type=submit value=""Verify Login"">"
                    '<INPUT type=submit  value='Verify Login' tabindex=4>
                    iHTMLEle.click()
                    foundButton = True
                    Exit For
                End If
                'End If
            Next

            If foundButton = False Then Throw New System.Exception("No button was found with outer HTML [" & buttonOuterHTML & "]")

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Click IE radio button
    ''' <summary>
    ''' Clicks specified radio button on web page.
    ''' </summary>
    ''' <param name="radioButtonName">Name of radio button group.</param>
    ''' <param name="radioButtonValue">Value of specific radio button to click.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub ClickRadioButton(ByVal radioButtonName As String, ByVal radioButtonValue As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Clicking radio button [" & radioButtonName & ":" & radioButtonValue & "]")


            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing

            document = myDocument.pHTMLDocument

            Dim elementCollection As mshtml.IHTMLElementCollection = document.getElementsByName(radioButtonName)
            Dim element As mshtml.IHTMLElement
            Dim matchFound As Boolean = False

            For Each element In elementCollection
                If element.value = radioButtonValue Then
                    Dim radioElement As mshtml.HTMLInputElement = Nothing
                    radioElement = element
                    radioElement.checked = True
                    matchFound = True
                    Exit For
                End If
            Next element

            If matchFound = False Then
                Throw New System.Exception("No elements were found with the name [" & radioButtonName & ":" & radioButtonValue & "]")
            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try
    End Sub

    'Click check box if only one found containing keyword
    ''' <summary>
    ''' Clicks a check box containing keyword specified.
    ''' </summary>
    ''' <param name="keyword">Keyword or phrase contained description of check box to click.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>Number of check boxes found with keywords specified.</returns>
    ''' <remarks>Will only click check box if exactly one match found containing specified keyword.
    ''' <para>Author: Nick Washburn</para></remarks>
    Public Function ClickCheckBoxContains(ByVal keyword As String, ByVal myDocument As WebDocument) As Integer ', ByVal frameName As String, Optional ByVal subframeName As String = "") As Integer
        Try
            _logger.Log("[IE] Finding check box containing [" & keyword & "]")

            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing

            document = myDocument.pHTMLDocument

            'Get HTML
            Dim htmlBody As String = ""
            Dim inputElements As New Collection
            Dim checkBoxNames As New Collection

            htmlBody = document.body.parentElement.innerHTML

            '--------------------------------------------
            'Break out checkbox elements from raw HTML
            '--------------------------------------------
            Dim finishedSearching As Boolean = False
            Dim tempHTMLBody As String = ""

            tempHTMLBody = htmlBody.ToString
            Do
                Try
                    Dim tempSearchValue As String = ""
                    Dim searchLocationInput As Integer = 0
                    Dim searchLocationBR As Integer = 0

                    searchLocationInput = InStr(tempHTMLBody.ToUpper, "<INPUT TYPE=")
                    If searchLocationInput = 0 Then
                        finishedSearching = True
                    Else
                        searchLocationBR = InStr(searchLocationInput, tempHTMLBody.ToUpper, "<BR>")
                        tempSearchValue = Mid(tempHTMLBody, searchLocationInput, searchLocationBR - searchLocationInput)
                        _logger.Log("Added [" & tempSearchValue & "] HTML block to collection")
                        inputElements.Add(tempSearchValue)
                        tempHTMLBody = Mid(tempHTMLBody, searchLocationBR)
                    End If
                Catch ex As Exception
                    _logger.Log("Can't add any more HTML blocks: " & ex.Message)
                    Exit Do
                End Try

            Loop Until finishedSearching


            '--------------------------------------------
            'Find matching Name(s)
            '--------------------------------------------
            For Each inpElement As String In inputElements
                Try
                    If inpElement.Contains(keyword) Then
                        _logger.Log("Analyzing [" & inpElement & "] HTML block")
                        Dim searchLocationName As Integer = 0
                        Dim searchLocationNameEnd As Integer = 0
                        searchLocationName = InStr(inpElement.ToUpper, "NAME=")
                        If searchLocationName > 0 Then
                            searchLocationNameEnd = InStr(searchLocationName + 5, inpElement, ">")
                            Dim nameToAdd As String = ""
                            nameToAdd = Mid(inpElement, searchLocationName + 5, searchLocationNameEnd - searchLocationName - 7)
                            checkBoxNames.Add(nameToAdd)

                            _logger.Log("Added [" & nameToAdd & "] name to collection")

                        End If

                    End If
                Catch ex As Exception
                    _logger.Log("Can't add any more names: " & ex.Message)
                    Exit For
                End Try
            Next inpElement

            Dim elementCollection As mshtml.IHTMLElementCollection = document.getElementsByTagName("input")
            Dim element As mshtml.IHTMLElement
            Dim numberOfMatches As Integer = 0

            '---------------------------------
            'Count the number of elements
            '---------------------------------
            For Each element In elementCollection

                For Each checkName As String In checkBoxNames
                    If element.outerHTML.Contains(checkName) Then
                        numberOfMatches = numberOfMatches + 1
                    End If
                Next checkName

            Next element

            _logger.Log(numberOfMatches & " matches found")

            '---------------------------------
            'Check appropriate element
            '---------------------------------
            If numberOfMatches = 1 Then
                For Each element In elementCollection

                    For Each checkName As String In checkBoxNames
                        If element.outerHTML.Contains(checkName) Then
                            Dim checkIElement As mshtml.HTMLInputElement = Nothing
                            checkIElement = element
                            checkIElement.checked = True
                            _logger.Log("Element [" & checkName & "] checked")

                            Return numberOfMatches

                        End If
                    Next checkName

                Next element

            End If

            'If numberOfMatches = 0 Then
            '            Throw New System.Exception("No elements were found with the name [" & checkBoxName & ":" & checkBoxValue & "]")
            'End If

            Return numberOfMatches

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try
    End Function

    'Click check box if only one found containing keyword
    ''' <summary>
    ''' Click radio button containing specified keyword.
    ''' </summary>
    ''' <param name="keyword">Keyword in radio button to click.</param>
    ''' <param name="radioButtonGroup">Group name of radio button.</param>
    ''' <param name="signature_Start">Signature to use when searching for keyword.</param>
    ''' <param name="signature_End">Signature to use when searching for keyword.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <param name="required">If true, throws exception if radio button not found.</param>
    ''' <returns>Returns number of radio button matches found.</returns>
    ''' <remarks>Only clicks radio button if exactly one match found with specified text.
    ''' <para>Author: Nick Washburn</para></remarks>
    Public Function ClickRadioButtonContains(ByVal keyword As String, ByVal radioButtonGroup As String, _
        ByVal signature_Start As String, ByVal signature_End As String, ByVal signature_Exclude_YN As Boolean, ByVal signature_Exclude As String, _
        ByVal myDocument As WebDocument, ByVal required As Boolean) As Integer ', ByVal frameName As String, Optional ByVal subframeName As String = "") As Integer

        Try
            _logger.Log("[IE] Finding check box containing [" & keyword & "]")

            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing
            document = myDocument.pHTMLDocument

            'Get HTML
            Dim htmlBody As String = ""
            Dim inputElements As New Collection
            Dim radioButtonNames As New Collection

            htmlBody = document.body.parentElement.innerHTML

            '--------------------------------------------
            'Break out radio button elements from raw HTML
            '--------------------------------------------
            Dim finishedSearching As Boolean = False
            Dim tempHTMLBody As String = ""

            tempHTMLBody = htmlBody.ToString
            Do
                Try
                    Dim tempSearchValue As String = ""
                    Dim searchLocationInput As Integer = 0
                    Dim searchLocationBR As Integer = 0

                    searchLocationInput = InStr(tempHTMLBody.ToUpper, signature_Start.ToUpper)
                    If searchLocationInput = 0 Then
                        finishedSearching = True
                    Else
                        searchLocationBR = InStr(searchLocationInput, tempHTMLBody.ToUpper, signature_End.ToUpper)
                        tempSearchValue = Mid(tempHTMLBody, searchLocationInput, searchLocationBR - searchLocationInput)
                        _logger.Log("Added [" & tempSearchValue & "] HTML block to collection")
                        inputElements.Add(tempSearchValue)
                        tempHTMLBody = Mid(tempHTMLBody, searchLocationBR)
                    End If
                Catch ex As Exception
                    _logger.Log("Can't add any more HTML blocks: " & ex.Message)
                    Exit Do
                End Try

            Loop Until finishedSearching


            '--------------------------------------------
            'Find matching Name(s)
            '--------------------------------------------
            For Each inpElement As String In inputElements
                Try
                    If inpElement.Contains(keyword) Then
                        If signature_Exclude_YN = False Or (signature_Exclude_YN = True And inpElement.Contains(signature_Exclude) = False) Then

                            _logger.Log("Analyzing [" & inpElement & "] HTML block")
                            Dim searchLocationName As Integer = 0
                            Dim searchLocationNameEnd As Integer = 0
                            searchLocationName = InStr(inpElement.ToUpper, "VALUE=")
                            If searchLocationName > 0 Then
                                searchLocationNameEnd = InStr(searchLocationName + 6, inpElement, "name=")
                                Dim nameToAdd As String = ""
                                nameToAdd = Mid(inpElement, searchLocationName + 6, searchLocationNameEnd - searchLocationName - 8)
                                radioButtonNames.Add(nameToAdd)

                                _logger.Log("Added [" & nameToAdd & "] name to collection")

                            End If

                        End If


                    End If
                Catch ex As Exception
                    _logger.Log("Can't add any more names: " & ex.Message)
                    Exit For
                End Try
            Next inpElement

            Dim elementCollection As mshtml.IHTMLElementCollection = document.getElementsByTagName("input")
            Dim element As mshtml.IHTMLElement
            Dim numberOfMatches As Integer = 0

            '---------------------------------
            'Count the number of elements
            '---------------------------------
            For Each element In elementCollection

                For Each radioButtonName As String In radioButtonNames
                    If element.outerHTML.Contains(radioButtonName) Then
                        If signature_Exclude_YN = False Or (signature_Exclude_YN = True And element.outerHTML.Contains(signature_Exclude) = False) Then

                            numberOfMatches = numberOfMatches + 1

                        End If

                    End If
                Next radioButtonName

            Next element

            _logger.Log(numberOfMatches & " matches found")

            '---------------------------------
            'Check appropriate element
            '---------------------------------
            If numberOfMatches = 1 Then

                Dim elementCollection2 As mshtml.IHTMLElementCollection = document.getElementsByName(radioButtonGroup)
                Dim element2 As mshtml.IHTMLElement

                For Each element2 In elementCollection2
                    For Each radioButtonName As String In radioButtonNames
                        If element2.outerHTML.ToUpper.Contains("VALUE=" & radioButtonName.ToUpper) Then
                            Dim radioElement As mshtml.HTMLInputElement = Nothing
                            radioElement = element2
                            radioElement.checked = True
                            radioElement.click()
                            radioElement.focus()
                            radioElement.FireEvent("onchange", radioElement)

                            Return numberOfMatches

                        End If
                    Next radioButtonName

                Next element2

                If required Then Throw New System.Exception("No elements were found with the name [" & radioButtonGroup & "] containing keyword " & keyword & "]")
            Else
                If required Then Throw New System.Exception("Multiple elements [" & numberOfMatches & "] were found with the name [" & radioButtonGroup & "] containing keyword " & keyword & "]")
            End If

            If numberOfMatches = 0 Then
                If required Then Throw New System.Exception("No elements were found with the name [" & radioButtonGroup & "] containing keyword " & keyword & "]")
            End If

            Return numberOfMatches

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try
    End Function

    'Select option element by text
    ''' <summary>
    ''' Select item specified from option list box on web page.
    ''' </summary>
    ''' <param name="selectElementName">Element name of option list box to select in.</param>
    ''' <param name="selectText">Text of item to select in list box.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>True if found match for specified text; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Overloads Function SelectOption(ByVal selectElementName As String, ByVal selectText As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Selecting value (" & selectText & ") from option group [" & selectElementName & "]")

            'Dim element As IHTMLElement = isolateHTMLElement(selectElementName, frameName)
            Dim element As mshtml.IHTMLElement = Nothing

            element = IsolateHTMLElement(selectElementName, myDocument)

            Dim foundMatch As Boolean = False

            Dim selectElement As mshtml.HTMLSelectElementClass
            selectElement = element

            For i As Integer = 0 To selectElement.length - 1
                If selectElement.options(i).innertext = Nothing Then
                    'Option is emptry string
                    If Len(selectText) = 0 Then
                        foundMatch = True
                        selectElement.selectedIndex = i
                        Exit For
                    End If
                Else
                    'Option contains string
                    If selectElement.options(i).innertext.ToString = selectText Then
                        foundMatch = True
                        selectElement.selectedIndex = i
                        Exit For
                    End If
                End If
            Next i

            Return foundMatch

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Select option element by index
    ''' <summary>
    ''' Select item specified from option list box on web page.
    ''' </summary>
    ''' <param name="selectElementName">Element name of option list box to select in.</param>
    ''' <param name="selectIndex">Index of item to select.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>Index of selected item.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Overloads Function SelectOption(ByVal selectElementName As String, ByVal selectIndex As Integer, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Selecting index (" & selectIndex & ") from option group [" & selectElementName & "]")

            'Dim element As IHTMLElement = isolateHTMLElement(selectElementName, frameName)
            Dim element As mshtml.IHTMLElement = Nothing

            element = IsolateHTMLElement(selectElementName, myDocument)

            Dim selectElement As mshtml.HTMLSelectElementClass = Nothing
            selectElement = element

            selectElement.selectedIndex = selectIndex

            Return selectElement.selectedIndex
        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Search a select option group
    ''' <summary>
    ''' Searches an option list for an item with specified text.
    ''' </summary>
    ''' <param name="selectElementName">Element name of option list box to select in.</param>
    ''' <param name="selectText">Text of item to select in list box.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>True if the item was found with the specified text; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function SearchSelect(ByVal selectElementName As String, ByVal selectText As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            Dim matchIndex As Integer = -1
            'Dim element As IHTMLElement = isolateHTMLElement(selectElementName, frameName)
            Dim element As mshtml.IHTMLElement = Nothing

            element = IsolateHTMLElement(selectElementName, myDocument)

            If element Is Nothing Then
                Throw New System.Exception("Could not isolate HTML element [" & selectElementName & "]")
            End If

            Dim selectElement As mshtml.HTMLSelectElementClass = Nothing
            selectElement = element

            For i As Integer = 0 To selectElement.length - 1
                Dim textToSearch As String = ""
                If selectElement.options(i).innerText Is Nothing Then
                    textToSearch = ""
                Else
                    textToSearch = selectElement.options(i).innertext.ToString
                End If

                If InStr(textToSearch, selectText) > 0 Then
                    matchIndex = i
                    Exit For
                End If
            Next i

            Return matchIndex


        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try
    End Function

    'Click IE anchor
    ''' <summary>
    ''' Clicks an HTML anchor with specified text on the web page.
    ''' </summary>
    ''' <param name="anchorName">Name of anchor to click.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub ClickAnchor(ByVal anchorName As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Clicking anchor [" & anchorName & "]")

            'Get document (parent or frame)
            Dim document As mshtml.HTMLDocument = Nothing

            document = myDocument.pHTMLDocument

            'Search element collection for anchor
            Dim elementCollection As mshtml.IHTMLElementCollection = document.getElementsByTagName("A")
            Dim element As mshtml.IHTMLElement = Nothing
            Dim matchFound As Boolean = False

            For Each element In elementCollection
                If element.innerText = anchorName Then
                    Dim anchor As mshtml.HTMLAnchorElement = Nothing
                    anchor = element

                    '_browserNonReadyCount = 1

                    anchor.click()
                    matchFound = True
                    Exit For
                End If
            Next element

            If matchFound = False Then
                Throw New System.Exception("No anchors were found with the name [" & anchorName & "]")
            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Verify browser contains text string
    ''' <summary>
    ''' Verifies if web page contains specified text.
    ''' </summary>
    ''' <param name="textVerify">Text to search for on web page.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>True if specified text found on web page; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function PageContainsText(ByVal textVerify As String, ByVal myDocument As WebDocument) As Boolean ', ByVal frameName As String, Optional ByVal subframeName As String = "") As Boolean
        Try
            If Len(textVerify) > 30 Then
                _logger.Log("[IE] Verifying browser contents (" & Left(textVerify, 27) & "...)")
            Else
                _logger.Log("[IE] Verifying browser contents (" & textVerify & ")")
            End If

            Dim browserText As String = ""

            browserText = GetBrowserText(myDocument)

            If InStr(browserText, Trim(textVerify), CompareMethod.Text) = 0 Then
                'Does not contain text
                'Try again after 5 seconds
                Threading.Thread.Sleep(5000)

                If Len(textVerify) > 22 Then
                    _logger.Log("[IE] Retry - verifying browser contents (" & Left(textVerify, 19) & "...)")
                Else
                    _logger.Log("[IE] Retry - verifying browser contents (" & textVerify & ")")
                End If


                browserText = GetBrowserText(myDocument)

                If InStr(browserText, Trim(textVerify), CompareMethod.Text) = 0 Then
                    Return False
                Else
                    'Does contain text after retry
                    Return True
                End If

            Else
                'Does contain text
                Return True
            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Verify browser contains text string, no log, no retry
    ''' <summary>
    ''' Verifies if web page contains specified text.
    ''' </summary>
    ''' <param name="textVerify">Text to search for on web page.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>True if specified text found on web page; otherwise, false.</returns>
    ''' <remarks>Does not insert log message or retry search if text not found.
    ''' <para>Author: Nick Washburn</para></remarks>
    Public Function PageContainsText2(ByVal textVerify As String, ByVal myDocument As WebDocument) As Boolean ', ByVal frameName As String, Optional ByVal subframeName As String = "") As Boolean
        Try
            'Dim browserText As String = getBrowserText(frameName)
            Dim browserText As String = ""
            Dim browserHTML As String = ""

            browserText = GetBrowserText(myDocument)

            browserHTML = GetBrowserHTML(myDocument)

            If InStr(browserText, Trim(textVerify), CompareMethod.Text) = 0 Then
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Get HTML contents of browser
    ''' <summary>
    ''' Gets HTML contents of browser web page.
    ''' </summary>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>Browser HTML (string).</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function GetBrowserHTML(ByVal myDocument As WebDocument) As String 'ByVal frameName As String) As String
        Try


            Return myDocument.pHTMLDocument.body.parentElement.innerHTML

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Get text contents of browser
    ''' <summary>
    ''' Gets text contents of browser web page.
    ''' </summary>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <returns>Browser text (string).</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function GetBrowserText(ByVal myDocument As WebDocument) As String 'ByVal frameName As String) As String
        Try

            Return myDocument.pHTMLDocument.body.parentElement.innerText

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    'Save current page to HTML file
    ''' <summary>
    ''' Saves current page HTML to a file.
    ''' </summary>
    ''' <param name="filename">Filename to save to.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub SaveBrowserHTML(ByVal filename As String, ByVal myDocument As WebDocument) 'ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Saving browser html to file")

            Threading.Thread.Sleep(500)

            Dim objStreamWriter As System.IO.StreamWriter = Nothing
            objStreamWriter = New System.IO.StreamWriter(filename, False)

            objStreamWriter.WriteLine(GetBrowserHTML(myDocument))
            objStreamWriter.Close()

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    'Save current page text contents to file
    ''' <summary>
    ''' Saves current page text to a file.
    ''' </summary>
    ''' <param name="filename">Filename to save to.</param>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Sub SaveBrowserText(ByVal filename As String, ByVal myDocument As WebDocument) ', ByVal frameName As String, Optional ByVal subframeName As String = "")
        Try
            _logger.Log("[IE] Saving browser text to file")

            Threading.Thread.Sleep(500)

            Dim objStreamWriter As System.IO.StreamWriter = Nothing
            objStreamWriter = New System.IO.StreamWriter(filename, False)

            objStreamWriter.WriteLine(GetBrowserText(myDocument))
            objStreamWriter.Close()

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub

    ''' <summary>
    ''' Verifies if a web page contains an iFrame.
    ''' </summary>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <param name="containsText">Text that appears in iFrame.</param>
    ''' <returns>True if the specified text is contained within the iFrame; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function PageContainsIFrame(ByVal myDocument As WebDocument, ByVal containsText As String) As Boolean

        Try
            Dim currentDocument As mshtml.HTMLDocument = myDocument.pHTMLDocument  '_ie.Document
            Dim windowFrame As mshtml.IHTMLWindow2 = currentDocument.frames
            Dim framesCollection As mshtml.IHTMLFramesCollection2 = windowFrame.frames
            Dim nbf As Integer = framesCollection.length

            Dim i As Integer = 0
            While i < nbf
                Dim index As Object = i
                Dim objFrame As Object = framesCollection.item(index)
                Dim frame As mshtml.IHTMLWindow2 = objFrame

                Dim frameDocument As mshtml.HTMLDocument = frame.document
                Dim searchText As String = frameDocument.body.innerText()

                If searchText.ToLower.Contains(containsText.ToLower) Then
                    Return True
                End If
                i = i + 1
            End While

            Return False

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    ''' <summary>
    ''' Gets HTMLDocument from IFrame.
    ''' </summary>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <param name="containsText">Text that is present in iFrame.</param>
    ''' <returns>HTMLDocument</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function GetIFrameDocument(ByVal myDocument As WebDocument, ByVal containsText As String) As WebDocument

        Try
            Dim currentDocument As mshtml.HTMLDocument = myDocument.pHTMLDocument  '_ie.Document
            Dim windowFrame As mshtml.IHTMLWindow2 = currentDocument.frames
            Dim framesCollection As mshtml.IHTMLFramesCollection2 = windowFrame.frames
            Dim nbf As Integer = framesCollection.length

            Dim i As Integer = 0
            While i < nbf
                Dim index As Object = i
                Dim objFrame As Object = framesCollection.item(index)
                Dim frame As mshtml.IHTMLWindow2 = objFrame

                Dim frameDocument As mshtml.HTMLDocument = frame.document
                Dim searchText As String = frameDocument.body.innerText()

                If searchText.ToLower.Contains(containsText.ToLower) Then
                    Return New WebDocument(frameDocument, Nothing)
                End If

                i = i + 1
            End While

            Throw New Exception("Could not find iFrame with text [" & containsText & "]")

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    ''' <summary>
    ''' Checks if HTML page contains specified frame.
    ''' </summary>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <param name="frameName">Name of frame.</param>
    ''' <returns>True if the HTML page contains the specified frame; otherwise, false.</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function PageContainsFrame(ByVal myDocument As WebDocument, ByVal frameName As String) As Boolean

        Try
            If Len(frameName) > 0 Then
                'Master Frame
                Dim currentDocument As Object = Nothing

                If myDocument Is Nothing Then
                    currentDocument = _ie.Document
                Else
                    currentDocument = myDocument.pHTMLDocument
                End If

                Dim i As Integer = 0

                For i = 0 To currentDocument.frames.length - 1
                    Dim currentWindow As mshtml.HTMLWindow2 = Nothing
                    currentWindow = currentDocument.frames(i)
                    If currentWindow.name = frameName Then
                        Return True

                    End If
                Next i
                Return False

            Else
                Return False

            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    ''' <summary>
    ''' Returns an BMS_IEDocument from specified frame.
    ''' </summary>
    ''' <param name="myDocument">BMS_IEDocument passed in by caller.</param>
    ''' <param name="frameName">Name of frame.</param>
    ''' <returns>BMS_IEDocument</returns>
    ''' <remarks>Author: Nick Washburn</remarks>
    Public Function GetFrameDocument(ByVal myDocument As WebDocument, ByVal frameName As String) As WebDocument

        Try
            If Len(frameName) > 0 Then
                'Master Frame
                Dim currentDocument As Object = Nothing

                If myDocument Is Nothing Then
                    currentDocument = _ie.Document
                Else
                    currentDocument = myDocument.pHTMLDocument
                End If

                Dim i As Integer = 0

                For i = 0 To currentDocument.frames.length - 1
                    Dim currentWindow As mshtml.HTMLWindow2 = Nothing
                    currentWindow = currentDocument.frames(i)
                    If currentWindow.name = frameName Then
                        Return New WebDocument(currentWindow.document, Nothing)

                    End If
                Next i
                Throw New System.Exception("Frame " & frameName & " was not found.")

            Else
                Return New WebDocument(_ie.Document, Nothing)

            End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Function

    Private Sub ie_DocumentComplete(ByVal pDisp As Object, ByRef URL As Object)
        Try
            'If _browserNonReadyCount > 0 Then
            '    _browserNonReadyCount = Threading.Interlocked.Decrement(_browserNonReadyCount)
            '    _logger.Log("[IE] Browser ready")
            Me.evtBrowserReady.Set()
            'Else
            '    _browserNonReadyCount = Threading.Interlocked.Decrement(_browserNonReadyCount)
            '    _logger.Log("[IE] Browser ready [*]")
            'End If

        Catch ex As Exception
            Throw New System.Exception("** DLL EXCEPTION ** " & ex.Message & " [" & Trim(ex.StackTrace) & "]", ex.InnerException)

        End Try

    End Sub


End Class
