Public Class WebECFInterface
    Inherits WebAbstractInterface
    Implements IDisposable

    Private _webBrowserOpen As Boolean = False
    Private _databaseControl As DatabaseControl = Nothing

    Public Sub New(ByVal mLogger As LogControl, ByVal mDatabaseControl As DatabaseControl)
        Try
            Dim tWebSession As New WebSession(mLogger)
            Initialize(tWebSession, mLogger)

            _databaseControl = mDatabaseControl

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public ReadOnly Property pWebBrowserOpen() As Boolean
        Get
            Return _webBrowserOpen
        End Get
    End Property

    Public Function IE_GetHTMLControl_ByKeyword(ByVal keyword As String) As WebElement
        Try
            Throw New BMSL.NotImplementedException("GetHTMLControlByKeyword not implemented.")

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function GetBlankWebElement(ByVal found As Boolean, ByVal id As String) As WebElement
        Try
            Return New WebElement(Nothing, found, id, "", "", "", "", "")

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetHTMLControl_ByID(ByVal id As String) As WebElement

        Try
            'Dim element As mshtml.IHTMLElement = _webSession.pIEWindow.IsolateHTMLElementByID(id, _webSession.pIEDocument)
            Dim element As mshtml.IHTMLElement = _webSession.pIEDocument.IsolateHTMLElementByID(id)

            If element Is Nothing Then
                Return New WebElement(Nothing, False, id, "", "", "", "", "")
            Else
                Return New WebElement(element, True, id, element.id, element.innerText, element.innerHTML, element.outerText, element.outerHTML)
            End If

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetHTMLControl_ByName(ByVal name As String) As WebElement()
        Try
            'If _webSession.pIEWindow.SearchHTMLElement(name, _webSession.pIEDocument) = True Then

            '    Dim element As mshtml.IHTMLElement = _webSession.pIEWindow.IsolateHTMLElement(name, _webSession.pIEDocument)
            '    Return New WebElement(element, True, name, element.id, element.innerText, element.innerHTML)
            'Else
            '    Return New WebElement(Nothing, False, name, "", "", "")
            'End If

            Dim elementCollection As mshtml.IHTMLElementCollection = _webSession.pIEDocument.pHTMLDocument.getElementsByName(name)

            If elementCollection.length = 0 Then

                Dim webElementArray(0) As WebElement
                webElementArray(0) = New WebElement(Nothing, False, name, "", "", "", "", "")
                Return webElementArray

            Else
                Dim webElementArray(elementCollection.length - 1) As WebElement
                Dim element As mshtml.IHTMLElement = Nothing

                Dim ix As Integer = 0
                For Each element In elementCollection
                    webElementArray(ix) = New WebElement(element, True, name, element.id, element.innerText, element.innerHTML, element.outerText, element.outerHTML)
                    ix += 1
                Next element

                Return webElementArray
            End If


        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetHTMLControl_ByIDPartial(ByVal id As String) As WebElement()
        Try
            Dim elementCollection As mshtml.IHTMLElementCollection = _webSession.pIEDocument.pHTMLDocument.all

            If elementCollection.length = 0 Then

                Dim webElementArray(0) As WebElement
                webElementArray(0) = New WebElement(Nothing, False, id, "", "", "", "", "")
                Return webElementArray

            Else
                Dim webElementArray(elementCollection.length - 1) As WebElement
                Dim element As mshtml.IHTMLElement = Nothing

                Dim ix As Integer = 0
                For Each element In elementCollection
                    If element.id.Trim.ToLower.Contains(id.Trim.ToLower) Then
                        webElementArray(ix) = New WebElement(element, True, id, element.id, element.innerText, element.innerHTML, element.outerText, element.outerHTML)
                    End If
                    ix += 1
                Next element

                Return webElementArray
            End If

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetHTMLControl_ByNameInArray(ByVal name As String, ByVal arrayNumber As Integer) As WebElement
        Try
            Dim webElements() As WebElement = IE_GetHTMLControl_ByName(name)

            If webElements.Length <= 1 Then
                'If webElements(0) is nothing Then Throw New HTMLElementNotFoundException("No elements were found with name [" & name & "].")

                Return webElements(0)
            Else
                'Throw New HTMLElementNotFoundException("Invalid number of elements found.  [" & webElements.Length & "] names found.  [" & name & "]")
                Return webElements(arrayNumber)

            End If

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetHTMLControl_ByNameUnique(ByVal name As String) As WebElement
        Try
            Dim webElements() As WebElement = IE_GetHTMLControl_ByName(name)

            If webElements.Length <= 1 Then
                'If webElements(0) is nothing Then Throw New HTMLElementNotFoundException("No elements were found with name [" & name & "].")

                'if webelements(0).pHTMLElement.
                If webElements.Length = 1 Then
                    If webElements(0).pOuterHTML.Trim.ToLower.Contains("display: none") And webElements(0).pOuterHTML.Trim.ToLower.Contains("type=button") Then
                        _logger.LogDetail("Suppressing element output (element contains [display:none] and [type=button]).  [" & webElements.Length & "] items found.  [" & name & "]")
                        '  Throw New HTMLElementNotFoundException("Invalid number of elements found (some elements contain [display:none] and [type=button]).  [" & webElements.Length & "] items found.  [" & name & "]")
                        Return New WebElement(Nothing, False, name, "", "", "", "", "")
                    End If
                End If

                Return webElements(0)
            Else
                Throw New HTMLElementNotFoundException("Invalid number of elements found.  [" & webElements.Length & "] items found.  [" & name & "]")

            End If

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetHTMLControl_ByNameAndTagType(ByVal name As String, ByVal tagType As String) As WebElement
        Try
            Dim webElementsTagType() As WebElement = IE_GetHTMLControl_ByTagType(tagType)
            Dim webElementsName() As WebElement = IE_GetHTMLControl_ByName(name)

            Dim foundElement As WebElement = Nothing
            Dim foundCount As Integer = 0

            If webElementsTagType.Length = 1 Then
                'If webElementsTagType(0) Is Nothing Then Throw New HTMLElementNotFoundException("No elements were found with tag type [" & tagType & "].")

                Return webElementsTagType(0)
            ElseIf webElementsName.Length = 1 Then
                'If webElementsName(0) Is Nothing Then Throw New HTMLElementNotFoundException("No elements were found with name [" & name & "].")

                Return webElementsName(0)
            Else

                'Cross reference tag type and name collections, see how many match
                For Each element_i As WebElement In webElementsTagType
                    For Each element_j As WebElement In webElementsName
                        If Not element_i.pInnerHTML Is Nothing And Not element_j.pInnerHTML Is Nothing Then
                            If element_i.pInnerHTML = element_j.pInnerHTML And element_i.pInnerHTML.Trim <> "" Then
                                'Found match
                                foundElement = element_j
                                foundCount += 1
                            End If
                        End If
                    Next element_j
                Next element_i

                'If foundCount = 0 Then Throw New HTMLElementNotFoundException("Invalid number of elements found.  [" & foundCount & "] elements found.  Name [" & name & "] Tag Type [" & tagType & "]")
                If foundCount > 1 Then Throw New HTMLElementNotFoundException("Invalid number of elements found.  [" & foundCount & "] elements found.  Name [" & name & "] Tag Type [" & tagType & "]")

                'If found then return foundelement
                If Not foundElement Is Nothing Then Return foundElement

                Return New WebElement(Nothing, False, name, "", "", "", "", "")

            End If
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetHTMLControl_ByTagType(ByVal tagType As String) As WebElement()
        Try
            Dim elementCollection As mshtml.IHTMLElementCollection = _webSession.pIEDocument.pHTMLDocument.getElementsByTagName(tagType)

            If elementCollection.length = 0 Then

                Dim webElementArray(0) As WebElement
                webElementArray(0) = New WebElement(Nothing, False, tagType, "", "", "", "", "")
                Return webElementArray

            Else
                Dim webElementArray(elementCollection.length - 1) As WebElement
                Dim element As mshtml.IHTMLElement = Nothing

                Dim ix As Integer = 0
                For Each element In elementCollection
                    webElementArray(ix) = New WebElement(element, True, tagType, element.id, element.innerText, element.innerHTML, element.outerText, element.outerHTML)
                    ix += 1
                Next element

                Return webElementArray

            End If

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_Button_Click(ByVal button As WebElement) As Boolean
        Try
            _logger.LogInfo("[WebECFInterface] Clicking button [" & button.pDescription & "]")

            Dim mButton As mshtml.HTMLButtonElement = button.pHTMLElement

            Me._webSession.pIEWindow.ClearReadyDocuments()

            mButton.click()
            mButton.FireEvent("onclick")

            Return True

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_Hyperlink_Click(ByVal hyperlink As WebElement) As Boolean
        Try
            _logger.LogInfo("[WebECFInterface] Clicking hyperlink [" & hyperlink.pDescription & "]")

            Dim mHyperlink As mshtml.HTMLAnchorElement = hyperlink.pHTMLElement

            Me._webSession.pIEWindow.ClearReadyDocuments()

            mHyperlink.click()
            mHyperlink.FireEvent("onclick")

            Return True

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_SelectList_SelectByIndex(ByVal selectList As WebElement, ByVal index As Integer) As Boolean

        Try
            _logger.LogInfo("[WebECFInterface] Selecting index [" & index & "] on select list [" & selectList.pDescription & "]")

            Dim selectElement As mshtml.HTMLSelectElementClass = selectList.pHTMLElement

            selectElement.selectedIndex = index
            selectElement.FireEvent("onclick")
            selectElement.FireEvent("onchange")

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_SelectList_SelectByValue(ByVal selectList As WebElement, ByVal value As String) As Boolean
        Try
            _logger.LogInfo("[WebECFInterface] Selecting value [" & value & "] on select list [" & selectList.pDescription & "]")

            Dim selectElement As mshtml.HTMLSelectElementClass = selectList.pHTMLElement

            For i As Integer = 0 To selectElement.length - 1
                If selectElement.options(i).innertext = Nothing Then
                    'Option is emptry string
                    If Len(value) = 0 Then
                        selectElement.selectedIndex = i
                        selectElement.FireEvent("onclick")
                        selectElement.FireEvent("onchange")

                        Return True
                    End If
                Else
                    'Option contains string
                    If selectElement.options(i).innertext.ToString.Trim.ToLower = value.Trim.ToLower Then
                        selectElement.selectedIndex = i
                        selectElement.FireEvent("onclick")
                        selectElement.FireEvent("onchange")

                        Return True
                    End If
                End If
            Next i

            Return False

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_SelectList_Search(ByVal selectList As WebElement, ByVal value As String, ByVal partialMatch As Boolean, Optional ByVal startingIx As Integer = 0) As Integer
        Try
            _logger.LogInfo("[WebECFInterface] Search for value [" & value & "] on select list [" & selectList.pDescription & "]")

            Dim selectElement As mshtml.HTMLSelectElementClass = selectList.pHTMLElement

            If startingIx > selectElement.length - 1 Then
                'Throw New HTMLInteractionException("Starting index specified [" & startingIx & "] is larger than select element items [" & selectElement.length - 1 & "].")
                Return -1
            End If

            For i As Integer = startingIx To selectElement.length - 1
                If selectElement.options(i).innertext = Nothing Then
                    'Option is emptry string
                    If Len(value) = 0 Then
                        'selectElement.selectedIndex = i
                        Return i
                    End If
                Else
                    'Option contains string
                    If partialMatch Then
                        If selectElement.options(i).innertext.ToString.Trim.ToLower.Contains(value.Trim.ToLower) Then
                            'selectElement.selectedIndex = i
                            Return i
                        End If
                    Else
                        If selectElement.options(i).innertext.ToString.Trim.ToLower = value.Trim.ToLower Then
                            'selectElement.selectedIndex = i
                            Return i
                        End If
                    End If

                End If
            Next i

            Return -1

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_CheckBox_Click(ByVal checkBox As WebElement) As Boolean
        Try
            _logger.LogInfo("[WebECFInterface] Clicking check box[" & checkBox.pDescription & "]")

            Dim mCheckBox As mshtml.HTMLInputElement = checkBox.pHTMLElement
            mCheckBox.click()
            mCheckBox.FireEvent("onclick")

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_RadioButton_Click(ByVal radioButton As WebElement) As Boolean
        Try
            _logger.LogInfo("[WebECFInterface] Clicking radio button [" & radioButton.pDescription & "]")

            Dim mRadioButton As mshtml.HTMLInputElement = radioButton.pHTMLElement
            mRadioButton.click()
            mRadioButton.FireEvent("onclick")

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_TextBox_Fill(ByVal textBox As WebElement, ByVal value As String) As Boolean
        Try
            _logger.LogInfo("[WebECFInterface] Filling text box [" & textBox.pDescription & "]")

            Dim mTextBox As mshtml.HTMLInputElement = textBox.pHTMLElement

            mTextBox.value = value
            mTextBox.FireEvent("onkeydown")
            mTextBox.FireEvent("onchange")
            mTextBox.FireEvent("onblur")

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_TextBox_Focus(ByVal textBox As WebElement) As Boolean

        Try
            _logger.LogInfo("[WebECFInterface] Focusing on textbox [" & textBox.pDescription & "]")

            Dim mTextBox As mshtml.HTMLInputElement = textBox.pHTMLElement

            mTextBox.focus()
            mTextBox.FireEvent("onfocus")

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_TextBox_FireJavascript(ByVal textBox As WebElement) As Boolean

        Try
            _logger.LogInfo("[WebECFInterface] Firing javascript on textbox [" & textBox.pDescription & "]")

            Dim mTextBox As mshtml.HTMLInputElement = textBox.pHTMLElement

            mTextBox.FireEvent("onkeydown")
            mTextBox.FireEvent("onchange")
            mTextBox.FireEvent("onblur")

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_HWND_DownKey_Click() As Boolean

        Try
            _logger.LogInfo("[WebECFInterface] Clicking down key")

            _webSession.pIEWindow.HWND_KeyDown()

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Sub IE_HWND_TextBox_Fill(ByVal text As String)
        Try
            _logger.LogInfo("[WebECFInterface] Entering text into box")

            _webSession.pIEWindow.HWND_TypeText(text)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub
    Public Sub IE_HWND_EnterKey_Click()
        Try
            _logger.LogInfo("[WebECFInterface] Clicking enter key")

            _webSession.pIEWindow.HWND_KeyEnter()

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Sub

    Public Function IE_PageText_Find(ByVal text As String) As Boolean
        Try
            _logger.LogInfo("[WebECFInterface] Finding text [" & text & "]")

            Dim pageText As String = ""
            pageText = _webSession.pIEDocument.GetText
            If pageText.ToLower.Contains(text.ToLower) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Private Function IE_IsHTMLRendered(ByVal pageHTML As String, ByVal htmlFragment As Boolean) As Boolean

        'This is until we figure out how to verify html fragments, ie no end tag
        If htmlFragment Then Return True

        Dim htmlTags As String() = SelectHTMLTags(BMSL.Constants.HTMLTagType.VerifyHTMLEndTag)

        For i As Integer = 0 To htmlTags.Length - 1

            If pageHTML.ToLower.Contains(htmlTags(i).ToLower.Trim) Then
                Return True

            End If
        Next i

        Return False

    End Function

    Public Function SelectHTMLTags(ByVal htmlTagType As Integer) As String()
        Try
            Dim sqlData1 As BMSL.DatabaseSession = _databaseControl.CreateNewSession("usp_Define_HTML_Tags__Select")
            sqlData1.AddParameter("@HTML_Tag_Type", SqlDbType.Int, 4, htmlTagType)

            Dim dsDataset1 As DataSet = sqlData1.SelectDataset()

            If dsDataset1.Tables(0).Rows.Count = 0 Then Throw New DataIntegrityException("HTML Tag Type lookup failed.  No rows returned.  [" & htmlTagType & "]")

            'Return original value set
            Dim ix1 As Integer = 0
            Dim returnString1(dsDataset1.Tables(0).Rows.Count - 1) As String

            For Each dsRow1 As DataRow In dsDataset1.Tables(0).Rows

                If Not dsRow1.IsNull("HTML_Tag") Then
                    returnString1(ix1) = dsRow1.Item("HTML_Tag")
                    ix1 += 1
                End If

            Next dsRow1

            Return returnString1

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetPageHTML(ByVal htmlFragment As Boolean) As String
        Try
            _logger.LogInfo("[WebECFInterface] Getting page HTML")
            '  Return _webSession.pIEDocument.GetHTML
            Dim pageHTML As String = ""

            Dim pageCompleteTimeout As Integer = 0
            Threading.Thread.Sleep(100)
            Do
                pageHTML = _webSession.pIEDocument.GetHTML
                If IE_IsHTMLRendered(pageHTML, htmlFragment) = True Then Return pageHTML

                _logger.LogInfo("Waiting for HTML to complete rendering")
                Threading.Thread.Sleep(1500)
                pageCompleteTimeout += 1

            Loop Until pageCompleteTimeout > 5

            _logger.LogWarning("HTML may not have completed rendering")

            Return pageHTML

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_GetPageText(ByVal htmlFragment As Boolean) As String
        Try
            _logger.LogInfo("[WebECFInterface] Getting page text")

            'Return _webSession.pIEDocument.GetText

            Dim pageHTML As String = ""

            Dim pageCompleteTimeout As Integer = 0
            Threading.Thread.Sleep(100)
            Do
                pageHTML = _webSession.pIEDocument.GetHTML
                If IE_IsHTMLRendered(pageHTML, htmlFragment) = True Then
                    Return _webSession.pIEDocument.GetText
                End If

                _logger.LogInfo("Waiting for HTML to complete rendering")
                Threading.Thread.Sleep(1500)
                pageCompleteTimeout += 1
            Loop Until pageCompleteTimeout > 5

            _logger.LogWarning("HTML may not have completed rendering")

            Return _webSession.pIEDocument.GetText

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IE_AcquirePopup(ByVal urlparts As String()) As Boolean
        Return _webSession.AcquirePopup(urlparts)

    End Function

    Public Function IE_DiscardPopup() As Boolean
        _webSession.DiscardPopup()
        Return True

    End Function

    Public Overrides Function IE_Navigate(ByVal url As String) As Boolean
        Try
            Me._webSession.pIEWindow.ClearReadyDocuments()

            _logger.LogInfo("[WebECFInterface] Navigating url")
            _webBrowserOpen = True
            _webSession.Navigate(url)
            Threading.Thread.Sleep(1500)

            Return True

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Overrides Function IE_Quit() As Boolean
        Try
            Me._webSession.pIEWindow.ClearReadyDocuments()

            _logger.LogInfo("[WebECFInterface] Closing browser")
            _webBrowserOpen = False
            _webSession.Quit()
            Return True

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function GetFormattedCheckpoint(ByVal description As String) As String
        Try
            Return description & "[" & Format(Now(), "MMddhhmmsffff") & "]"
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    'Public Sub SetBrowserReadyCheckpoint(ByVal checkpoint As String)
    '    Try
    '        _webSession.pIEWindow.AddBrowserReadyCheckpoint(checkpoint)

    '    Catch ex As Exception
    '        _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '        Throw ex

    '    End Try

    'End Sub

    Public Function IE_WaitForBrowserIdle()
        Return IE_WaitForBrowserIdle(1, 120000)
    End Function

    Public Function IE_WaitForBrowserIdle(ByVal expectedDocuments As Integer)
        Return IE_WaitForBrowserIdle(expectedDocuments, 120000)

    End Function
    Public Function IE_WaitForBrowserIdle(ByVal expectedDocuments As Integer, ByVal timeout As Integer) 'Optional ByVal expectedDocuments As Integer = 1, Optional ByVal timeout As Integer = 120000) As Boolean
        If expectedDocuments > 50 Then Throw New BMSL.HTMLInteractionException("Expected documents count is too high.  [" & expectedDocuments & "]")
        If timeout > 12000000 Then Throw New BMSL.HTMLInteractionException("Timeout is too high.  [" & timeout & "]")

        Dim endTime As Date = DateAdd(DateInterval.Second, timeout / 1000, Now())
        Dim wait1 As Integer = (timeout + 2) / 2
        Dim wait2 As Integer = (timeout + 2) / 8

        Dim status As String = ""
        Dim lastStatus As String = ""
        Dim idleCount As Integer = 0

        Dim currentDocument As Integer = expectedDocuments
        Dim ix As Integer = 0
        Do

            ix += 1

            Dim waitOK As Boolean = False
            If currentDocument > 0 Then
                If Me._webSession.pIEWindow.pReadyDocuments >= expectedDocuments Then
                    'Expected documents are already ready
                    _logger.LogInfo("[IE] document(s) are already ready. [" & _webSession.pIEWindow.pReadyDocuments & "]")
                    Return True
                Else
                    If currentDocument = expectedDocuments Then
                        _logger.LogInfo("[IE] Waiting for first document")
                        waitOK = _webSession.pIEWindow.BrowserReadyWait(wait1)
                    Else
                        _logger.LogInfo("[IE] Waiting for additional document")
                        waitOK = _webSession.pIEWindow.BrowserReadyWait(wait2)
                    End If
                    currentDocument -= 1
                End If
            Else
                waitOK = False
            End If

            Dim busy As Boolean = _webSession.pIEWindow.BrowserBusy
            lastStatus = status

            If Not waitOK And busy Then
                'Wait timed out and busy.  
                'The page is still loading.
                '
                If lastStatus <> status Or ix Mod 5 = 0 Then _logger.LogInfo("[IE] Browser busy [doc:complete]")
                status = "tb"
                idleCount = 0

            ElseIf waitOK And busy Then
                'Wait OK and busy.
                'The next page is probably loading.
                '
                If lastStatus <> status Or ix Mod 5 = 0 Then _logger.LogInfo("[IE] Browser busy [doc:ready]")
                status = "rb"
                idleCount = 0

            ElseIf Not waitOK And Not busy Then
                'Wait timed out and not busy.
                'The pages are probably done loading.
                '
                If lastStatus <> status Or ix Mod 5 = 0 Then _logger.LogInfo("[IE] Browser idle [doc:complete]")
                status = "ti"
                idleCount += 1

            ElseIf waitOK And Not busy Then
                'Wait OK and not busy.
                'Need to check for the next page.
                '
                If lastStatus <> status Or ix Mod 5 = 0 Then _logger.LogInfo("[IE] Browser idle [doc:ready]")
                status = "ri"
                idleCount += 1

            End If

            If lastStatus = "ri" And status = "ti" And idleCount >= 3 And currentDocument <= 0 Then
                _logger.LogInfo("[IE] Browser ready")
                Return True
            ElseIf lastStatus = "ti" And status = "ti" And idleCount >= 3 And currentDocument <= 0 Then
                _logger.LogInfo("[IE] Browser ready")
                Return True
            End If

            Threading.Thread.Sleep(500)

        Loop Until Now() > endTime

        _logger.LogWarning("[IE] Browser idle timed out")
        Return False

    End Function

    'Public Function IE_WaitForBrowserReady(ByVal checkpoint As String, Optional ByVal timeout As Integer = 60000) As Boolean
    '    Try
    '        Dim chk As WebCheckpoint = _webSession.pIEWindow.GetBrowserReadyCheckpoint(checkpoint)

    '        If chk.Checked = True Then
    '            'Already done
    '            _logger.LogInfo("Browser is ready [" & checkpoint & "] already complete [IE_WaitForBrowserReady]")
    '            Threading.Thread.Sleep(500)
    '        Else
    '            _webSession.pIEWindow.BrowserReadyWait(timeout)
    '            _logger.LogInfo("Browser is ready [" & checkpoint & "] [IE_WaitForBrowserReady]")
    '            Threading.Thread.Sleep(500)
    '        End If

    '        _webSession.pIEWindow.RemoveBrowserReadyCheckpoint(checkpoint)
    '        Return True

    '    Catch ex As Exception
    '        _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '        Throw ex

    '    End Try


    'End Function

    'Public Function IE_WaitForBrowserReadyOrDynamicText(ByVal checkpoint As String, ByVal dynamicText() As String, Optional ByVal timeout As Integer = 60000) As Boolean
    '    Try
    '        Dim chk As WebCheckpoint = _webSession.pIEWindow.GetBrowserReadyCheckpoint(checkpoint)

    '        Dim endTime As Date = DateAdd(DateInterval.Second, timeout / 1000, Now())

    '        Do
    '            'Check for browser ready 
    '            If chk.Checked = True Then
    '                'Checkpoint reached
    '                _logger.LogInfo("Browser is ready [" & checkpoint & "] complete [IE_WaitForBrowserReadyOrDynamicText]")

    '                _webSession.pIEWindow.RemoveBrowserReadyCheckpoint(checkpoint)
    '                Threading.Thread.Sleep(500)
    '                Return True
    '            End If

    '            'Check for dynamic text
    '            Try
    '                For Each text As String In dynamicText
    '                    Dim pageText As String = Me.IE_GetPageText(True)
    '                    If pageText.Trim.ToLower.Contains(text.Trim.ToLower) Then
    '                        'Dynamic text reached
    '                        _logger.LogInfo("Browser is ready [" & text & "] reached [IE_WaitForBrowserReadyOrDynamicText]")

    '                        _webSession.pIEWindow.RemoveBrowserReadyCheckpoint(checkpoint)
    '                        Threading.Thread.Sleep(500)
    '                        Return True

    '                    End If
    '                Next text
    '          Catch ex As Exception

    '            End Try
    '            'Wait one cycle
    '            Threading.Thread.Sleep(1000)
    '        Loop Until Now() > endTime

    '        _logger.LogInfo("[IE] The wait request timed out [IE_WaitForBrowserReadyOrDynamicText]")

    '        _webSession.pIEWindow.RemoveBrowserReadyCheckpoint(checkpoint)
    '        Return False

    '    Catch ex As Exception
    '        _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '        Throw ex

    '    End Try

    'End Function

    'Public Function IE_WaitForDynamicText(ByVal dynamicText() As String, Optional ByVal timeout As Integer = 60000) As Boolean
    '    Try
    '        Dim endTime As Date = DateAdd(DateInterval.Second, timeout / 1000, Now())

    '        Do

    '            'Check for browser busy
    '            If _webSession.pIEWindow.Busy Then
    '                'Threading.Thread.Sleep(1000)
    '                'Don't do anthing
    '            Else
    '                'Check for dynamic text
    '                Try
    '                    For Each text As String In dynamicText
    '                        Dim pageText As String = Me.IE_GetPageText(True)
    '                        If pageText.Trim.ToLower.Contains(text.Trim.ToLower) Then
    '                            'Dynamic text reached
    '                            _logger.LogInfo("Browser is ready [" & text & "] reached [IE_WaitForDynamicText]")

    '                            Threading.Thread.Sleep(500)
    '                            Return True

    '                        End If
    '                    Next text
    '                Catch ex As Exception

    '                End Try
    '            End If

    '            'Wait one cycle
    '            Threading.Thread.Sleep(1000)
    '        Loop Until Now() > endTime

    '        _logger.LogInfo("[IE] The wait request timed out [IE_WaitForDynamicText]")

    '        Return False

    '    Catch ex As Exception
    '        _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '        Throw ex

    '    End Try


    'End Function

    'Public Sub IE_WaitForTimeElapsed(ByVal milliseconds As Integer)
    '    Try
    '        Throw New BMSL.NotImplementedException("This method is not implemented.")
    '    Catch ex As Exception
    '        _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
    '        Throw ex

    '    End Try

    'End Sub

    Private _dismissDialogRetryMax As Integer = 0
    Public Sub IE_HWND_DismissDialog(Optional ByVal retryMax As Integer = 10)

        _dismissDialogRetryMax = retryMax
        Dim t1 As New Threading.Thread(AddressOf IE_HWND_DismissIEDialogThread)
        t1.Start()

    End Sub

    Private Sub IE_HWND_DismissIEDialogThread()

        Dim ix As Integer = 0

        Do
            Try
                If Not _webSession Is Nothing Then
                    If Not _webSession.pIEWindow Is Nothing Then
                        _webSession.pIEWindow.HWND_DismissIEDialog()
                        Threading.Thread.Sleep(1000)
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
                ix += 1
            Catch ex As Exception
                _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
                Exit Do
            End Try

        Loop While ix < _dismissDialogRetryMax

    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free managed resources when explicitly called
            End If

            ' TODO: free shared unmanaged resources
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
