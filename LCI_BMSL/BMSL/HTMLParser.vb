Public Class HTMLParser

    Dim _logger As LogControl = Nothing

    Public Sub New(ByVal mLogger As LogControl)
        _logger = mLogger

    End Sub

    Public Function ParseNextTagFromHTML(ByVal parentHTML As String, ByVal tagOpen As String, ByVal tagClose As String, ByVal keyword As String, _
       ByVal nextTagOpen As String, ByVal nextTagClose As String, ByVal nextTagKeyword As String) As HTMLTag

        Try
            Dim startPos As Integer = 1
            Dim tag As HTMLTag = ParseTag(parentHTML, tagOpen, tagClose, keyword, startPos)

            If tag.pComments.Trim.Length > 0 Then Return tag

            If tag.pTagHTMLLocation + Len(tag.pTagHTML) = 0 Then tag.pComments = tag.pComments & "; Tag HTML Location + length is invalid" : Return tag

            Dim nextTag As HTMLTag = ParseTag(parentHTML, nextTagOpen, nextTagClose, nextTagKeyword, tag.pTagHTMLLocation + Len(tag.pTagHTML))

            Return nextTag

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try


    End Function

    Public Function ParseTagFromHTML(ByVal parentHTML As String, ByVal tagOpen As String, ByVal tagClose As String, ByVal keyword As String) As HTMLTag

        Try
            Dim startPos As Integer = 1

            Return ParseTag(parentHTML, tagOpen, tagClose, keyword, startPos)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try


    End Function

    Private Function ParseTag(ByVal parentHTML As String, ByVal tagOpen As String, ByVal tagClose As String, ByVal keyword As String, ByVal startPos As Integer) As HTMLTag
        Try

            Dim tag As HTMLTag = New HTMLTag(parentHTML, tagOpen, tagClose, keyword, "")

            Dim comments As String = ""
            Dim parseError As Boolean = False
            If Not tagOpen.Contains("<") Then tag.pComments = "Open tag must contain bracket <" : Return tag
            'If Not tagOpen.Contains(">") Then tag.p_Comments = "Open tag must contain brackets < >" : Return tag
            'If Not tagClose.Contains("<") Then tag.p_Comments = "Close tag must contain brackets < >" : Return tag
            'If Not tagClose.Contains(">") Then tag.p_Comments = "Close tag must contain brackets < >" : Return tag
            If Not tagClose.Contains("</") Then tag.pComments = "Close tag must contain a closing bracket </" : Return tag

            If keyword.Trim.Length = 0 Then tag.pComments = "Keyword length must be greater than zero" : Return tag
            If parentHTML.Trim.Length = 0 Then tag.pComments = "Parent HTML length must be greater than zero" : Return tag


            Dim pos As Integer = startPos

            Dim foundTagOpen As Boolean = False
            Dim foundTagClose As Boolean = False
            Dim foundKeyword As Boolean = False

            Dim tagOpenPos As Integer = 0
            Dim tagClosePos As Integer = 0
            Dim tagLen As Integer = 0

            While pos < parentHTML.Length

                'Try to find open tag
                tagOpenPos = InStr(pos, parentHTML.ToLower, tagOpen.ToLower)
                If tagOpenPos > 0 Then
                    foundTagOpen = True

                    'Try to find close tag
                    tagClosePos = InStr(tagOpenPos, parentHTML.ToLower, tagClose.ToLower)
                    If tagClosePos > 0 Then
                        foundTagClose = True

                        'Try to find keyword
                        Dim possibleKeyword As String = Mid(parentHTML, tagOpenPos, tagClosePos - tagOpenPos + Len(tagClose))

                        If possibleKeyword.ToLower.Contains(keyword.ToLower) Then
                            tagLen = tagClosePos - tagOpenPos + Len(tagClose)
                            'Found keyword 
                            foundKeyword = True
                            Exit While
                        Else
                            'Move on to next part of html
                            pos = tagClosePos + 1
                        End If
                    End If
                End If

                'open/cloase tag not found in html (or next part of html)
                If tagOpenPos = 0 Then Exit While
                If tagClosePos = 0 Then Exit While

            End While

            If foundTagOpen = False Then tag.pComments = "Tag " & tagOpen & " not found." : Return tag
            If foundTagClose = False Then tag.pComments = "Tag " & tagClose & " not found." : Return tag
            If foundKeyword = False Then tag.pComments = "Keyword [" & keyword & "] not found." : Return tag

            'Return tag contents
            tag.pFoundTag = True
            tag.pTagHTML = Mid(parentHTML, tagOpenPos, tagClosePos - tagOpenPos + Len(tagClose))
            tag.pComments = ""
            tag.pTagHTMLLocation = tagOpenPos

            Return tag

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

End Class
