Public Class WebElement

    Private _element As mshtml.IHTMLElement = Nothing
    Private _isFound As Boolean = False
    Private _description As String = ""
    Private _id As String = ""
    Private _innerHTML As String = ""
    Private _innerText As String = ""
    Private _outerHTML As String = ""
    Private _outerText As String = ""

    Public ReadOnly Property pHTMLElement() As mshtml.IHTMLElement
        Get
            Return _element
        End Get
    End Property

    Public ReadOnly Property pIsFound() As Boolean
        Get
            Return _isFound
        End Get
    End Property

    Public ReadOnly Property pDescription() As String
        Get
            Return _description
        End Get
    End Property

    Public ReadOnly Property pID() As String
        Get
            Return _id
        End Get
    End Property

    Public ReadOnly Property pInnerText() As String
        Get
            Return _innerText
        End Get
    End Property

    Public ReadOnly Property pInnerHTML() As String
        Get
            Return _innerHTML
        End Get
    End Property

    Public ReadOnly Property pOuterText() As String
        Get
            Return _outerText
        End Get
    End Property

    Public ReadOnly Property pOuterHTML() As String
        Get
            Return _outerHTML
        End Get
    End Property

    Public Sub ResetDescription(ByVal mDescription As String)
        _description = mDescription
    End Sub

    Public Sub New(ByVal element As mshtml.IHTMLElement, ByVal isFound As Boolean, ByVal description As String, ByVal id As String, ByVal innerText As String, ByVal innerHTML As String, ByVal outerText As String, ByVal outerHTML As String)
        _element = element
        _isFound = isFound
        _description = description
        _id = id
        _innerText = innerText
        _innerHTML = innerHTML
        _outerText = outerText
        _outerHTML = outerHTML

    End Sub

End Class
