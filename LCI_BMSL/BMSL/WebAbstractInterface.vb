Public MustInherit Class WebAbstractInterface

    Protected _webSession As WebSession = Nothing
    Protected _logger As LogControl = Nothing

    Protected Function pWebSession() As WebSession
        Return _webSession
    End Function

    Protected Sub Initialize(ByVal mWebSession As WebSession, ByVal mLogger As LogControl)
        Try
            _webSession = mWebSession
            _logger = mLogger

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public MustOverride Function IE_Navigate(ByVal url As String) As Boolean
    Public MustOverride Function IE_Quit() As Boolean

    'Public MustOverride Sub SetBrowserReadyCheckpoint(ByVal checkpoint As String)
    'Public MustOverride Function IE_WaitForBrowserReady(ByVal checkpoint As String, Optional ByVal timeout As Integer = 60000) As Boolean
    'Public MustOverride Function IE_WaitForDynamicText(ByVal dynamicText As String(), Optional ByVal timeout As Integer = 60000) As Boolean
    'Public MustOverride Function IE_WaitForBrowserReadyOrDynamicText(ByVal checkpoint As String, ByVal dynamicText() As String, Optional ByVal timeout As Integer = 60000) As Boolean
    'Public MustOverride Sub IE_WaitForTimeElapsed(ByVal milliseconds As Integer)

    'Public MustOverride Function IE_Button_Click(ByVal button As WebElement) As Boolean
    'Public MustOverride Function IE_Hyperlink_Click(ByVal button As WebElement) As Boolean
    'Public MustOverride Function IE_TextBox_Fill(ByVal textBox As WebElement, ByVal value As String) As Boolean
    'Public MustOverride Function IE_CheckBox_Click(ByVal checkBox As WebElement) As Boolean
    'Public MustOverride Function IE_RadioButton_Click(ByVal radioButton As WebElement) As Boolean
    'Public MustOverride Function IE_PageText_Find(ByVal text As String) As Boolean
    'Public MustOverride Function IE_GetPageText(ByVal htmlFragment As Boolean) As String
    'Public MustOverride Function IE_GetPageHTML(ByVal htmlFragment As Boolean) As String

    'Public MustOverride Function IE_GetHTMLControl_ByNameAndTagType(ByVal name As String, ByVal tagType As String) As WebElement
    'Public MustOverride Function IE_GetHTMLControl_ByName(ByVal name As String) As WebElement()
    'Public MustOverride Function IE_GetHTMLControl_ByNameUnique(ByVal name As String) As WebElement
    'Public MustOverride Function IE_GetHTMLControl_ByID(ByVal id As String) As WebElement
    'Public MustOverride Function IE_GetHTMLControl_ByTagType(ByVal tagType As String) As WebElement()
    'Public MustOverride Function IE_GetHTMLControl_ByKeyword(ByVal keyword As String) As WebElement
    'Public MustOverride Function IE_GetHTMLControl_ByTagKeywords(ByVal beginTag As String, ByVal endTag As String) As WebElement

End Class
