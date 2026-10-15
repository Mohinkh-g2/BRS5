Public Class BMS_IEHTMLTag

    Private _parentHTML As String = ""
    Private _parentTag As BMS_IEHTMLTag = Nothing

    Private _tagHTML As String = ""
    Private _tagHTMLLocation As Integer = 0
    Private _tagOpen As String = ""
    Private _tagClose As String = ""
    Private _keywordNearOpen As String = ""
    Private _keywordNearClose As String = ""

    Private _foundTag As Boolean = False

    Private _comments As String = ""

    Public Sub New(ByVal parentHTML As String, ByVal tagOpen As String, ByVal tagClose As String, ByVal keywordNearOpen As String, ByVal keywordNearClose As String)

        _parentHTML = parentHTML
        _tagOpen = tagOpen
        _tagClose = tagClose
        _keywordNearOpen = keywordNearOpen
        _keywordNearClose = keywordNearClose

    End Sub

    Public Property pFoundTag() As Boolean
        Get
            Return _foundTag
        End Get
        Set(ByVal value As Boolean)
            _foundTag = True
        End Set
    End Property

    Public Property pParentHTML() As String
        Get
            Return _parentHTML
        End Get
        Set(ByVal value As String)
            _parentHTML = value
        End Set
    End Property

    Public Property pTagHTMLLocation() As Integer
        Get
            Return _tagHTMLLocation
        End Get
        Set(ByVal value As Integer)
            _tagHTMLLocation = value
        End Set
    End Property

    Public Property pParentTag() As BMS_IEHTMLTag
        Get
            Return _parentTag
        End Get
        Set(ByVal value As BMS_IEHTMLTag)
            _parentTag = value
        End Set
    End Property

    Public Property pTagHTML() As String
        Get
            Return _tagHTML
        End Get
        Set(ByVal value As String)
            _tagHTML = value
        End Set
    End Property

    Public Property pTagOpen() As String
        Get
            Return _tagOpen
        End Get
        Set(ByVal value As String)
            _tagOpen = value
        End Set
    End Property

    Public Property pTagClose() As String
        Get
            Return _tagClose
        End Get
        Set(ByVal value As String)
            _tagClose = value
        End Set
    End Property

    Public Property pKeywordNearOpen() As String
        Get
            Return _keywordNearOpen
        End Get
        Set(ByVal value As String)
            _keywordNearOpen = value
        End Set
    End Property

    Public Property pKeywordNearClose() As String
        Get
            Return _keywordNearClose
        End Get
        Set(ByVal value As String)
            _keywordNearClose = value
        End Set
    End Property

    Public Property pComments() As String
        Get
            Return _comments
        End Get
        Set(ByVal value As String)
            _comments = value
        End Set
    End Property

End Class
