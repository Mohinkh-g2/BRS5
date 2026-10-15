Public Class ASPPDFOps

    Private _logger As BMSL.LogControl = Nothing

    Public Sub New(ByVal mLogger As BMSL.LogControl)
        Try
            _logger = mLogger

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Sub

    Public Function PrintToPDF(ByVal importHTMLFilename As String, ByVal exportPDFFilename As String) As String
        Try
            _logger.LogInfo("Printing [" & importHTMLFilename & "] to pdf [" & exportPDFFilename & "]")

            Dim Pdf As ASPPDFLib.IPdfManager
            Pdf = New ASPPDFLib.PdfManager
            Dim Doc As ASPPDFLib.IPdfDocument

            Pdf = New ASPPDFLib.PdfManager
            Doc = Pdf.CreateDocument
            Doc.ImportFromUrl(importHTMLFilename)

            If IO.File.Exists(exportPDFFilename) Then IO.File.Move(exportPDFFilename, exportPDFFilename & "[" & Format(Now(), "yyyyMMddHHmmss") & "].pdf")

            Dim result As String = Doc.Save(exportPDFFilename, False)

            Return result

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try
    End Function

    Public Function MergeExecute(ByVal importPDFDocsFolder As String, ByVal exportPDFFilename As String) As Integer


        Try
            _logger.LogInfo("Merging PDF's")
            _logger.LogInfo("Opening new pdf document")

            Dim Pdf As ASPPDFLib.PdfManager
            Pdf = New ASPPDFLib.PdfManager
            'Dim NewPdf As New ASPPDFLib.PdfManager()

            Dim Doc As ASPPDFLib.IPdfDocument = Nothing
            Dim NewDoc As ASPPDFLib.IPdfDocument = Nothing


            If IO.File.Exists(exportPDFFilename) Then IO.File.Move(exportPDFFilename, exportPDFFilename & "[" & Format(Now(), "yyyyMMddHHmmss") & "].pdf")

            Dim fileCollection() As String = System.IO.Directory.GetFiles(importPDFDocsFolder)

            'Scan the folder for files
            Dim fileIx As Integer = 0

            Dim primaryDocument As String = ""

            For Each fullFilename As String In fileCollection

                '---------------------------------
                'Get file information
                Dim documentFilename As String = ""
                Dim documentPath As String = ""
                Dim documentExtension As String = ""

                'Split filename by \
                Dim filenameArray1 As String() = fullFilename.Split("\")
                Dim filePart As String = filenameArray1(filenameArray1.Length - 1)
                Dim filePath As String = Left(fullFilename, Len(fullFilename) - Len(filePart)) 'keep last slash

                documentFilename = filePart
                documentPath = filePath

                'Split filename by .
                Dim filenameArray2 As String() = documentFilename.Split(".")
                Dim filePartExt As String = filenameArray2(filenameArray2.Length - 1)
                Dim filePartNoExt As String = Left(documentFilename, Len(documentFilename) - Len(filePartExt) - 1) 'remove period

                documentExtension = filePartExt

                _logger.LogInfo("Found document [" & fullFilename & "]")

                'Check if this is the first or 2-n file
                If fileIx = 0 Then
                    primaryDocument = fullFilename
                ElseIf fileIx = 1 Then
                    MergeDocuments(primaryDocument, fullFilename, exportPDFFilename)
                Else
                    IO.File.Move(exportPDFFilename, exportPDFFilename & ".pending[" & fileIx & "].pdf")
                    MergeDocuments(exportPDFFilename & ".pending[" & fileIx & "].pdf", fullFilename, exportPDFFilename)
                    Do
                        Try
                            IO.File.Delete(exportPDFFilename & ".pending[" & fileIx & "].pdf")
                            Exit Do
                        Catch ex As Exception
                            _logger.LogInfo("Waiting for file [" & exportPDFFilename & ".pending[" & fileIx & "].pdf" & "]")
                            Threading.Thread.Sleep(500)
                        End Try
                    Loop
                End If

                fileIx += 1

            Next fullFilename


        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Private Function MergeDocuments(ByVal primaryDocument As String, ByVal appendDocument As String, ByVal outputDocument As String) As Boolean

        Dim Pdf As ASPPDFLib.IPdfManager = Nothing
        Dim Pri_Doc As ASPPDFLib.IPdfDocument = Nothing
        Dim Append_Doc As ASPPDFLib.IPdfDocument = Nothing

        Try
            Pdf = New ASPPDFLib.PdfManager

            Dim Path As String
            Dim DocName As String

            If IO.File.Exists(outputDocument) Then IO.File.Delete(outputDocument)

            _logger.LogInfo("Merging documents [" & primaryDocument & " ] and [" & appendDocument & "] to [" & outputDocument & "]")

            DocName = outputDocument
            Path = DocName
            'Doc = Pdf.OpenDocument(Path)

            'Doc.Save("C:\A.pdf")
            'Doc.Close()

            Pri_Doc = Pdf.OpenDocument(primaryDocument)

            Append_Doc = Pdf.OpenDocument(appendDocument)

            Pri_Doc.AppendDocument(Append_Doc)

            'TODO modified this value
            Dim pPage As ASPPDFLib.IPdfPage = Pri_Doc.Pages(1)
            Dim pParam As ASPPDFLib.IPdfParam = Pdf.CreateParam()
            pParam("FitH").Value = 1

            Dim pDest As ASPPDFLib.IPdfDest = pPage.CreateDest(pParam)

            Pri_Doc.OpenAction = pDest

            Pri_Doc.Save(Path, True)
        
            Path = Nothing
            DocName = Nothing

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex
        Finally

            Try
                Pri_Doc.Close()
            Catch ex As Exception
                _logger.LogError("Could not close primary PDF: " & ex.Message)
            End Try

            Try
                Append_Doc.Close()
            Catch ex As Exception
                _logger.LogError("Could not close append PDF: " & ex.Message)
            End Try

            Pdf = Nothing
            Pri_Doc = Nothing
            Append_Doc = Nothing

        End Try

    End Function

    Public Function Merge(ByVal importPDFDocsFolder As String, ByVal exportPDFFilename As String) As Integer

        Try
            _logger.LogInfo("Opening new pdf document")

            Dim Pdf As ASPPDFLib.PdfManager
            Pdf = New ASPPDFLib.PdfManager
            'Dim NewPdf As New ASPPDFLib.PdfManager()

            Dim Doc As ASPPDFLib.IPdfDocument = Nothing
            Dim NewDoc As ASPPDFLib.IPdfDocument = Nothing


            If IO.File.Exists(exportPDFFilename) Then IO.File.Move(exportPDFFilename, exportPDFFilename & "[" & Format(Now(), "yyyyMMddHHmmss") & "].pdf")

            ' NewDoc = Pdf.CreateDocument()
            'NewDoc.Save(exportPDFFilename)
            'NewDoc.Close()
            'NewDoc = Nothing
            Threading.Thread.Sleep(1000)

            ' NewDoc = Pdf.OpenDocument(exportPDFFilename)

            Dim fileCollection() As String = System.IO.Directory.GetFiles(importPDFDocsFolder)

            'Scan the folder for files
            Dim fileIx As Integer = 0

            For Each fullFilename As String In fileCollection

                '---------------------------------
                'Get file information
                Dim documentFilename As String = ""
                Dim documentPath As String = ""
                Dim documentExtension As String = ""

                'Split filename by \
                Dim filenameArray1 As String() = fullFilename.Split("\")
                Dim filePart As String = filenameArray1(filenameArray1.Length - 1)
                Dim filePath As String = Left(fullFilename, Len(fullFilename) - Len(filePart)) 'keep last slash

                documentFilename = filePart
                documentPath = filePath

                'Split filename by .
                Dim filenameArray2 As String() = documentFilename.Split(".")
                Dim filePartExt As String = filenameArray2(filenameArray2.Length - 1)
                Dim filePartNoExt As String = Left(documentFilename, Len(documentFilename) - Len(filePartExt) - 1) 'remove period

                documentExtension = filePartExt

                _logger.LogInfo("Found document [" & fullFilename & "]")

                'Check if this is the first or 2-n file
                _logger.LogInfo("Opening [" & fullFilename & "]")
                Doc = Pdf.OpenDocument(fullFilename)
                If fileIx = 0 Then
                    _logger.LogInfo("Saving as [" & exportPDFFilename & "]")
                    Doc.Save(exportPDFFilename)
                ElseIf fileIx = 1 Then
                    _logger.LogInfo("Opening [" & exportPDFFilename & "]")
                    NewDoc = Pdf.OpenDocument(exportPDFFilename)
                    _logger.LogInfo("Appending [" & fullFilename & "] to [" & exportPDFFilename & "]")
                    NewDoc.AppendDocument(Doc)
                Else
                    _logger.LogInfo("Appending [" & fullFilename & "] to [" & exportPDFFilename & "]")
                    NewDoc.AppendDocument(Doc)
                    'NewDoc.Save(exportPDFFilename)
                    'NewDoc.Close()
                End If
                _logger.LogInfo("Closing [" & fullFilename & "]")
                Doc.Close()

                fileIx += 1

            Next fullFilename

            _logger.LogInfo("Saving [" & exportPDFFilename & ".new.pdf" & "]")
            NewDoc.Save(exportPDFFilename & ".new.pdf")
            _logger.LogInfo("Closing [" & exportPDFFilename & ".new.pdf" & "]")
            NewDoc.Close()

            '_logger.LogInfo("Closing pdf document")

            Threading.Thread.Sleep(1000)

            'NewDoc.Close()
            'NewDoc.Save(exportPDFFilename, True)

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function
    'Public Function Merge(ByVal importPDFDocsFolder As String, ByVal exportPDFFilename As String) As Integer

    '    Dim Pdf As ASPPDFLib.IPdfManager
    '    Pdf = New ASPPDFLib.PdfManager
    '    Dim Doc As ASPPDFLib.IPdfDocument
    '    'Dim Doc_Temp As ASPPDFLib.IPdfDocument

    '    'Dim Path As String
    '    'Dim DocName As String

    '    'Dim myDataSet As DataSet = Nothing
    '    'Dim myDataTable As DataTable = Nothing
    '    'Dim myDataRow As DataRow = Nothing

    '    'Dim myDataTable2 As DataTable = Nothing
    '    'Dim myDataRow2 As DataRow = Nothing


    '    'Dim Generate_PDF As Generate_PDF_Documents
    '    'Generate_PDF = New Generate_PDF_Documents
    '    'Generate_PDF.GetConnectString = ConnectString


    '    Try

    '        Dim fileCollection() As String = System.IO.Directory.GetFiles(importPDFDocsFolder)

    '        'Scan the folder for files and update their information

    '        For Each fullFilename As String In fileCollection
    '            Dim fullPath As String = String.Empty
    '            Dim filename As String = String.Empty

    '            'Set full path 
    '            fullPath = fullFilename
    '            Dim filename1 As String() = fullFilename.Split("\")

    '            'Get filename only
    '            If filename1.Length > 0 Then
    '                filename = filename1(filename1.Length - 1)
    '            Else
    '                filename = fullFilename
    '            End If

    '            Trace.TraceInformation("[Ifw]Found [" & filename & "]")

    '            '   DoFileOperations(inputFolder, destfolder, fullFilename, filename)

    '        Next fullFilename

    '        'myDataSet = Generate_PDF.Return_Merge_POCs(DateString(Today()))
    '        'myDataTable = myDataSet.Tables(0)

    '        Dim RootGeneratedDocs As String = ""
    '        Dim RootAcceptedTemp As String = ""
    '        Dim RootAcceptedDocs As String = ""

    '        '    If myDataTable.Rows.Count > 0 Then
    '        'For Each myDataRow In myDataTable.Rows

    '        DocName = Mid(myDataRow("Unique_Identifier").ToString, 1, 20) & "_POC.pdf"
    '        Path = RootGeneratedDocs & DocName
    '        Doc = Pdf.OpenDocument(Path)
    '        Doc.Save(RootAcceptedTemp & DocName, False)
    '        Doc.Close()

    '        Doc = Pdf.OpenDocument(RootAcceptedTemp & DocName, False)

    '        ' myDataTable2 = Generate_PDF.Return_Supporting_Documents(myDataRow("Unique_Identifier").ToString)



    '        '   If myDataTable2.Rows.Count > 0 Then
    '        'For Each myDataRow2 In myDataTable2.Rows

    '        Doc_Temp = Pdf.OpenDocument(RootAcceptedDocs & myDataRow2("Document_Name").ToString)
    '        Doc.AppendDocument(Doc_Temp)

    '        '  Next

    '        '  End If

    '        Doc.Save(Path, True)
    '        Doc.Close()

    '        'delete the temp B-10

    '        If System.IO.File.Exists(RootAcceptedTemp & DocName) = True Then

    '            System.IO.File.Delete(RootAcceptedTemp & DocName)
    '        End If



    '        ' Next



    '        '  End If

    '        'messagetxt.Text = "Number of Mergings: " & myDataTable.Rows.Count

    '        Pdf = Nothing
    '        Doc = Nothing
    '        'Doc_Temp = Nothing

    '        'Path = Nothing
    '        'DocName = Nothing

    '        'myDataSet = Nothing
    '        'myDataTable = Nothing
    '        'myDataRow = Nothing

    '        'myDataTable2 = Nothing
    '        'myDataRow2 = Nothing

    '        ' Generate_PDF = Nothing

    '    Catch ex As Exception

    '        '      messagetxt.Text = "Error Appending B10 for OR or CAE " & ex.ToString
    '    End Try
    '    Return Nothing

    'End Function

End Class
