Public Class WebDocument

    Private _HTMLDocument As mshtml.HTMLDocument = Nothing
    Private _logger As LogControl = Nothing

    Public Sub New(ByVal tHTMLDocument As mshtml.HTMLDocument, ByVal mLogger As LogControl)
        Try
            _HTMLDocument = tHTMLDocument
            _logger = mLogger

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Function GetText() As String
        Try
            Return _HTMLDocument.body.parentElement.innerText

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            'Throw ex
            Return ""

        End Try

    End Function

    Public Function GetHTML() As String
        Try
            Return _HTMLDocument.body.parentElement.innerHTML

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            'Throw ex
            Return ""

        End Try

    End Function

    Public Function IsolateHTMLElementByID(ByVal elementID As String) As mshtml.IHTMLElement
        Try
            Dim element As mshtml.IHTMLElement = _HTMLDocument.getElementById(elementID)

            Return element

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IsolateHTMLElementByName(ByVal elementName As String) As mshtml.IHTMLElementCollection
        Try
            Dim elementCollection As mshtml.IHTMLElementCollection = _HTMLDocument.getElementsByName(elementName)

            Return elementCollection

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function IsolateHTMLElementByTagName(ByVal tagName As String) As mshtml.IHTMLElementCollection

        Try
            Dim elementCollection As mshtml.IHTMLElementCollection = _HTMLDocument.getElementsByTagName(tagName)

            Return elementCollection

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public ReadOnly Property pHTMLDocument() As mshtml.HTMLDocument
        Get
            Return _HTMLDocument
        End Get
    End Property

End Class
