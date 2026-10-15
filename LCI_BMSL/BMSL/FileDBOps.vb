Public Class FileDBOps

    Private _logger As BMSL.LogControl = Nothing
    Private _databaseControl As BMSL.DatabaseControl = Nothing

    Public Sub New(ByVal logger As BMSL.LogControl, ByVal mDatabaseControl As BMSL.DatabaseControl)

        _logger = logger
        _databaseControl = mDatabaseControl

    End Sub

    Public Function InsertBinaryDocumentsInFolder(ByVal importFolder As String, ByVal foreignIDType As String, ByVal documentType As Integer) As Integer
        Try
            Dim fileCollection() As String = System.IO.Directory.GetFiles(importFolder)

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


                'Split filename by _
                Dim filenameArray3 As String() = filePartNoExt.Split("_")
                Dim foreignID As String = filenameArray3(0)
                Dim documentSource As Integer = BMSL.Constants.DocumentSource.ExternalSource
                Dim documentStatus As Integer = BMSL.Constants.DocumentStatus.UploadPending

                If Not IsNumeric(foreignID) Then
                    _logger.LogWarning("WARNING: Foreign ID is not numeric. [" & foreignID & "]")
                    foreignID = 0
                End If

                'Insert document 
                InsertBinaryDocument(foreignID, foreignIDType, "", "", documentType, documentSource, documentStatus, _
                    fullFilename, documentExtension, documentFilename)

                fileIx += 1

            Next fullFilename

            Return fileIx

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function InsertAsciiDocument(ByVal foreignID As Integer, ByVal foreignIDType As String, _
     ByVal lciCourtNumber As String, ByVal petitionSequenceNumber As String, _
    ByVal documentType As Integer, ByVal documentSource As Integer, ByVal documentStatus As Integer, _
    ByVal ascii As String, ByVal documentExtension As String, ByVal documentDescription As String) As Boolean

        Try

            _logger.LogInfo("Inserting ascii document to database")

            Dim sqlString As String = "usp_Documents__Insert_Ascii"
            Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)
            sqlData.AddParameter("@LCI_Court_Number", SqlDbType.VarChar, 5, lciCourtNumber)
            sqlData.AddParameter("@Petition_Sequence_Number", SqlDbType.VarChar, 8, petitionSequenceNumber)
            sqlData.AddParameter("@Foreign_ID", SqlDbType.Int, 4, foreignID)
            sqlData.AddParameter("@Foreign_ID_Type", SqlDbType.VarChar, 125, foreignIDType)
            sqlData.AddParameter("@Document_Type", SqlDbType.Int, 4, documentType)
            sqlData.AddParameter("@Document_Source", SqlDbType.Int, 4, documentSource)
            sqlData.AddParameter("@Document_Status", SqlDbType.Int, 4, documentStatus)
            sqlData.AddParameter("@Document_Ascii", SqlDbType.VarChar, Len(ascii), ascii)
            sqlData.AddParameter("@Document_Extension", SqlDbType.VarChar, 10, documentExtension)
            sqlData.AddParameter("@Document_Description", SqlDbType.VarChar, 250, documentDescription)

            Dim rowsAffected As Integer = sqlData.ExecuteNonQuery
            _logger.LogInfo(rowsAffected & " row(s) affected")

            If rowsAffected = 0 Then Throw New BMSL.DataIntegrityException("SQL Insert [" & sqlString & "] failed.  No rows were inserted.")

            Return True
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        End Try

    End Function

    Public Function InsertBinaryDocument(ByVal foreignID As Integer, ByVal foreignIDType As String, _
        ByVal lciCourtNumber As String, ByVal petitionSequenceNumber As String, _
        ByVal documentType As Integer, ByVal documentSource As Integer, ByVal documentStatus As Integer, _
        ByVal binaryFilename As String, ByVal documentExtension As String, ByVal documentDescription As String) As Boolean

        Dim binaryStream As IO.FileStream = Nothing
        Dim binaryReader As IO.BinaryReader = Nothing

        Try

            _logger.LogInfo("Inserting binary document to database")

            'Get binary document from file
            binaryStream = New IO.FileStream(binaryFilename, IO.FileMode.Open, IO.FileAccess.Read)
            binaryReader = New IO.BinaryReader(binaryStream)
            Dim fileLength As Integer = binaryReader.BaseStream.Length
            Dim binaryBuffer(fileLength) As Byte
            binaryReader.Read(binaryBuffer, 0, fileLength)

            binaryReader.Close()
            binaryStream.Close()

            'Insert binary document
            Dim sqlString As String = "usp_Documents__Insert_Binary"
            Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)
            sqlData.AddParameter("@LCI_Court_Number", SqlDbType.VarChar, 5, lciCourtNumber)
            sqlData.AddParameter("@Petition_Sequence_Number", SqlDbType.VarChar, 8, petitionSequenceNumber)
            sqlData.AddParameter("@Foreign_ID", SqlDbType.Int, 4, foreignID)
            sqlData.AddParameter("@Foreign_ID_Type", SqlDbType.VarChar, 125, foreignIDType)
            sqlData.AddParameter("@Document_Type", SqlDbType.Int, 4, documentType)
            sqlData.AddParameter("@Document_Source", SqlDbType.Int, 4, documentSource)
            sqlData.AddParameter("@Document_Status", SqlDbType.Int, 4, documentStatus)
            sqlData.AddParameter("@Document_Binary", SqlDbType.VarBinary, fileLength, binaryBuffer)
            sqlData.AddParameter("@Document_Extension", SqlDbType.VarChar, 10, documentExtension)
            sqlData.AddParameter("@Document_Description", SqlDbType.VarChar, 250, documentDescription)

            Dim rowsAffected As Integer = sqlData.ExecuteNonQuery
            _logger.LogInfo(rowsAffected & " row(s) affected")

            If rowsAffected = 0 Then Throw New BMSL.DataIntegrityException("SQL Insert [" & sqlString & "] failed.  No rows were inserted.")


            Return True
        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        Finally
            binaryReader.Close()
            binaryStream.Close()

            binaryReader = Nothing
            binaryStream = Nothing

        End Try

    End Function

    Public Function SelectAsciiDocument(ByVal filePath As String, ByVal documentID As Integer) As String

        Dim textWriter As IO.StreamWriter = Nothing

        Try

            _logger.LogInfo("Selecting ascii document")

            Dim filename As String = ""
            Dim documentExtension As String = ""

            'Get binary from database
            Dim sqlString As String = "usp_Documents__Select"
            Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)
            sqlData.AddParameter("@Document_ID", SqlDbType.Int, 4, documentID)
            sqlData.AddParameter("@Foreign_ID", SqlDbType.Int, 4, -1)
            sqlData.AddParameter("@Foreign_ID_Type", SqlDbType.VarChar, 125, "")
            sqlData.AddParameter("@Document_Type", SqlDbType.Int, 4, -1)
            'sqlData.AddParameter("@LCI_Court_Number", SqlDbType.VarChar, 5, "")
            'sqlData.AddParameter("@Petition_Sequence_Number", SqlDbType.VarChar, 8, "")

            Dim dsDataset As DataSet = sqlData.SelectDataset()
            _logger.LogInfo(dsDataset.Tables(0).Rows.Count & " row(s) returned")

            If dsDataset.Tables(0).Rows.Count = 0 Then Throw New BMSL.DataIntegrityException("Select Ascii document failed.  No rows returned.  Document ID [" & documentID & "]")

            Dim lciCourtNumber As String = ""
            Dim petitionSequenceNumber As String = ""
            Dim documentType As String = ""

            If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Extension") Then documentExtension = dsDataset.Tables(0).Rows(0).Item("Document_Extension")
            If Not dsDataset.Tables(0).Rows(0).IsNull("LCI_Court_Number") Then lciCourtNumber = dsDataset.Tables(0).Rows(0).Item("LCI_Court_Number")
            If Not dsDataset.Tables(0).Rows(0).IsNull("Petition_Sequence_Number") Then petitionSequenceNumber = dsDataset.Tables(0).Rows(0).Item("Petition_Sequence_Number")
            If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Type") Then documentType = dsDataset.Tables(0).Rows(0).Item("Document_Type")

            If Len(Trim(documentExtension)) > 0 Then
                If Right(filePath, 1) = "\" Then filename = filePath & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType & "." & documentExtension _
                Else filename = filePath & "\" & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType & "." & documentExtension
            Else
                If Right(filePath, 1) = "\" Then filename = filePath & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType _
                Else filename = filePath & "\" & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType
            End If

            If IO.File.Exists(filename) Then IO.File.Move(filename, filename & "[" & Format(Now(), "yyyyMMddHmmssff") & "]")

            If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Ascii") Then
                'Dim binaryLength = Len(dsDataset.Tables(0).Rows(0).Item("Document_Binary"))
                'Dim binaryBuffer() As Byte = dsDataset.Tables(0).Rows(0).Item("Document_Binary")

                'Stream to file
                textWriter = New IO.StreamWriter(filename, False)

                textWriter.Write(dsDataset.Tables(0).Rows(0).Item("Document_Ascii"))

                textWriter.Close()

            Else
                Throw New BMSL.DataIntegrityException("Select Ascii document failed.  Row returned contains null value.  Document ID [" & documentID & "]")
            End If


            'Return filename
            Return filename

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        Finally

            textWriter.Close()

            textWriter = Nothing

        End Try

    End Function

    Public Function SelectAsciiDocuments(ByVal filePath As String, ByVal foreignID As Integer, ByVal foreignIDType As String, ByVal documentType As String) As String()

        Dim textWriter As IO.StreamWriter = Nothing

        Try

            _logger.LogInfo("Selecting ascii documents")

            Dim filename As String = ""
            Dim documentExtension As String = ""

            'Get binary from database
            Dim sqlString As String = "usp_Documents__Select"
            Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)
            sqlData.AddParameter("@Document_ID", SqlDbType.Int, 4, -1)
            sqlData.AddParameter("@Foreign_ID", SqlDbType.Int, 4, foreignID)
            sqlData.AddParameter("@Foreign_ID_Type", SqlDbType.VarChar, 125, foreignIDType)
            sqlData.AddParameter("@Document_Type", SqlDbType.Int, 4, documentType)
            'sqlData.AddParameter("@LCI_Court_Number", SqlDbType.VarChar, 5, "")
            'sqlData.AddParameter("@Petition_Sequence_Number", SqlDbType.VarChar, 8, "")

            Dim dsDataset As DataSet = sqlData.SelectDataset()
            _logger.LogInfo(dsDataset.Tables(0).Rows.Count & " row(s) returned")

            If dsDataset.Tables(0).Rows.Count = 0 Then Throw New BMSL.DataIntegrityException("Select Ascii documents failed.  No rows returned.  Foreign ID [" & foreignID & "] Type [" & foreignIDType & "]")


            Dim filenames(dsDataset.Tables(0).Rows.Count - 1) As String
            Dim ix As Integer = 0

            For Each dsRow As DataRow In dsDataset.Tables(0).Rows

                Dim lciCourtNumber As String = ""
                Dim petitionSequenceNumber As String = ""
                'Dim documentID As Integer = 0

                If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Extension") Then documentExtension = dsDataset.Tables(0).Rows(0).Item("Document_Extension")
                If Not dsDataset.Tables(0).Rows(0).IsNull("LCI_Court_Number") Then lciCourtNumber = dsDataset.Tables(0).Rows(0).Item("LCI_Court_Number")
                If Not dsDataset.Tables(0).Rows(0).IsNull("Petition_Sequence_Number") Then petitionSequenceNumber = dsDataset.Tables(0).Rows(0).Item("Petition_Sequence_Number")
                If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Type") Then documentType = dsDataset.Tables(0).Rows(0).Item("Document_Type")

                If Len(Trim(documentExtension)) > 0 Then
                    If Right(filePath, 1) = "\" Then filename = filePath & foreignID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType & "." & documentExtension _
                    Else filename = filePath & "\" & foreignID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType & "." & documentExtension
                Else
                    If Right(filePath, 1) = "\" Then filename = filePath & foreignID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType _
                    Else filename = filePath & "\" & foreignID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType
                End If

                If IO.File.Exists(filename) Then IO.File.Move(filename, filename & "[" & Format(Now(), "yyyyMMddHmmssff") & "]")

                If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Ascii") Then
                    'Dim binaryLength = Len(dsDataset.Tables(0).Rows(0).Item("Document_Binary"))
                    'Dim binaryBuffer() As Byte = dsDataset.Tables(0).Rows(0).Item("Document_Binary")

                    'Stream to file
                    textWriter = New IO.StreamWriter(filename, False)

                    textWriter.Write(dsDataset.Tables(0).Rows(0).Item("Document_Ascii"))

                    textWriter.Close()

                Else
                    Throw New BMSL.DataIntegrityException("Select Ascii document failed.  Row returned contains null value.  Foreign ID [" & foreignID & "]")
                End If


                filenames(ix) = filename
                ix += 1

            Next dsRow 'Return filename

            'Return filename array
            Return filenames

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        Finally

            textWriter.Close()

            textWriter = Nothing

        End Try

    End Function

    Public Function SelectBinaryDocument(ByVal filePath As String, ByVal documentID As Integer) As String

        Dim binaryStream As IO.FileStream = Nothing
        Dim binaryWriter As IO.BinaryWriter = Nothing

        Try

            _logger.LogInfo("Selecting binary document")

            Dim filename As String = ""
            Dim documentExtension As String = ""

            'Get binary from database
            Dim sqlString As String = "usp_Documents__Select"
            Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)
            sqlData.AddParameter("@Document_ID", SqlDbType.Int, 4, documentID)
            sqlData.AddParameter("@Foreign_ID", SqlDbType.Int, 4, -1)
            sqlData.AddParameter("@Foreign_ID_Type", SqlDbType.VarChar, 125, "")
            sqlData.AddParameter("@Document_Type", SqlDbType.Int, 4, -1)
            'sqlData.AddParameter("@LCI_Court_Number", SqlDbType.VarChar, 5, "")
            'sqlData.AddParameter("@Petition_Sequence_Number", SqlDbType.VarChar, 8, "")

            Dim dsDataset As DataSet = sqlData.SelectDataset()
            _logger.LogInfo(dsDataset.Tables(0).Rows.Count & " row(s) returned")

            If dsDataset.Tables(0).Rows.Count = 0 Then Throw New BMSL.DataIntegrityException("Select Binary document failed.  No rows returned.  Document ID [" & documentID & "]")

            Dim lciCourtNumber As String = ""
            Dim petitionSequenceNumber As String = ""
            Dim documentType As String = ""

            If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Extension") Then documentExtension = dsDataset.Tables(0).Rows(0).Item("Document_Extension")
            If Not dsDataset.Tables(0).Rows(0).IsNull("LCI_Court_Number") Then lciCourtNumber = dsDataset.Tables(0).Rows(0).Item("LCI_Court_Number")
            If Not dsDataset.Tables(0).Rows(0).IsNull("Petition_Sequence_Number") Then petitionSequenceNumber = dsDataset.Tables(0).Rows(0).Item("Petition_Sequence_Number")
            If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Type") Then documentType = dsDataset.Tables(0).Rows(0).Item("Document_Type")

            If Len(Trim(documentExtension)) > 0 Then
                If Right(filePath, 1) = "\" Then filename = filePath & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType & "." & documentExtension _
                Else filename = filePath & "\" & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType & "." & documentExtension
            Else
                If Right(filePath, 1) = "\" Then filename = filePath & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType _
                Else filename = filePath & "\" & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType
            End If


            If IO.File.Exists(filename) Then IO.File.Move(filename, filename & "[" & Format(Now(), "yyyyMMddHmmssff") & "]")

            If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Binary") Then
                ' Dim binaryLength = Len(dsDataset.Tables(0).Rows(0).Item("Document_Binary"))
                Dim binaryBuffer() As Byte = dsDataset.Tables(0).Rows(0).Item("Document_Binary")
                Dim binaryLength As Integer = binaryBuffer.Length

                'Stream to file
                binaryStream = New IO.FileStream(filename, IO.FileMode.CreateNew, IO.FileAccess.Write)
                binaryWriter = New IO.BinaryWriter(binaryStream)
                'Dim fileLength As Integer = binaryReader.BaseStream.Length
                'Dim binaryBuffer(fileLength) As Byte
                'binaryReader.Read(binaryBuffer, 0, fileLength)

                binaryWriter.Write(binaryBuffer, 0, binaryLength)

                binaryWriter.Close()
                binaryStream.Close()

            Else
                Throw New BMSL.DataIntegrityException("Select Binary document failed.  Row returned contains null value.  Document ID [" & documentID & "]")
            End If


            'Return filename
            Return filename

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        Finally
            binaryWriter.Close()
            binaryStream.Close()

            binaryWriter = Nothing
            binaryStream = Nothing

        End Try

    End Function

    Public Function SelectBinaryDocuments(ByVal filePath As String, ByVal foreignID As Integer, ByVal foreignIDType As String, ByVal documentType As String) As String()

        Dim binaryStream As IO.FileStream = Nothing
        Dim binaryWriter As IO.BinaryWriter = Nothing

        Try

            _logger.LogInfo("Selecting binary documents")

            Dim documentExtension As String = ""

            'Get binary from database
            Dim sqlString As String = "usp_Documents__Select"
            Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)
            sqlData.AddParameter("@Document_ID", SqlDbType.Int, 4, -1)
            sqlData.AddParameter("@Foreign_ID", SqlDbType.Int, 4, foreignID)
            sqlData.AddParameter("@Foreign_ID_Type", SqlDbType.VarChar, 125, foreignIDType)
            sqlData.AddParameter("@Document_Type", SqlDbType.Int, 4, documentType)
            'sqlData.AddParameter("@LCI_Court_Number", SqlDbType.VarChar, 5, "")
            'sqlData.AddParameter("@Petition_Sequence_Number", SqlDbType.VarChar, 8, "")

            Dim dsDataset As DataSet = sqlData.SelectDataset()
            _logger.LogInfo(dsDataset.Tables(0).Rows.Count & " row(s) returned")

            If dsDataset.Tables(0).Rows.Count = 0 Then Throw New BMSL.DataIntegrityException("Select Binary documents failed.  No rows returned.  Foreign ID [" & foreignID & "] Type [" & foreignIDType & "]")

            Dim filenames(dsDataset.Tables(0).Rows.Count - 1) As String
            Dim ix As Integer = 0

            For Each dsRow As DataRow In dsDataset.Tables(0).Rows

                Dim lciCourtNumber As String = ""
                Dim petitionSequenceNumber As String = ""
                Dim documentID As Integer = 0

                Dim filename As String = ""

                If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Extension") Then documentExtension = dsDataset.Tables(0).Rows(0).Item("Document_Extension")
                If Not dsDataset.Tables(0).Rows(0).IsNull("LCI_Court_Number") Then lciCourtNumber = dsDataset.Tables(0).Rows(0).Item("LCI_Court_Number")
                If Not dsDataset.Tables(0).Rows(0).IsNull("Petition_Sequence_Number") Then petitionSequenceNumber = dsDataset.Tables(0).Rows(0).Item("Petition_Sequence_Number")
                If Not dsDataset.Tables(0).Rows(0).IsNull("Document_ID") Then documentID = dsDataset.Tables(0).Rows(0).Item("Document_ID")

                If Len(Trim(documentExtension)) > 0 Then
                    If Right(filePath, 1) = "\" Then filename = filePath & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType & "." & documentExtension _
                    Else filename = filePath & "\" & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType & "." & documentExtension
                Else
                    If Right(filePath, 1) = "\" Then filename = filePath & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType _
                    Else filename = filePath & "\" & documentID & "_" & lciCourtNumber & "_" & petitionSequenceNumber & "_" & documentType
                End If


                Threading.Thread.Sleep(100)
                If IO.File.Exists(filename) Then IO.File.Move(filename, filename & "[" & Format(Now(), "yyyyMMddHmmssff") & "]")

                If Not dsDataset.Tables(0).Rows(0).IsNull("Document_Binary") Then
                    ' Dim binaryLength = Len(dsDataset.Tables(0).Rows(0).Item("Document_Binary"))
                    Dim binaryBuffer() As Byte = dsDataset.Tables(0).Rows(0).Item("Document_Binary")
                    Dim binaryLength As Integer = binaryBuffer.Length

                    'Stream to file
                    binaryStream = New IO.FileStream(filename, IO.FileMode.CreateNew, IO.FileAccess.Write)
                    binaryWriter = New IO.BinaryWriter(binaryStream)
                    'Dim fileLength As Integer = binaryReader.BaseStream.Length
                    'Dim binaryBuffer(fileLength) As Byte
                    'binaryReader.Read(binaryBuffer, 0, fileLength)

                    binaryWriter.Write(binaryBuffer, 0, binaryLength)

                    binaryWriter.Close()
                    binaryStream.Close()

                Else
                    Throw New BMSL.DataIntegrityException("Select Binary document failed.  Row returned contains null value.  Document ID [" & documentID & "]")
                End If

                filenames(ix) = filename
                ix += 1

            Next dsRow

            'Return filename array
            Return filenames

        Catch ex As Exception
            _logger.LogMethodException(Me.GetType.Name.ToString() & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name.ToString(), ex.Message, ex.StackTrace)
            Throw ex

        Finally
            binaryWriter.Close()
            binaryStream.Close()

            binaryWriter = Nothing
            binaryStream = Nothing

        End Try

    End Function

    Public Sub ZipFolderContents(ByVal zipFilename As String, ByVal inputFolder As String, ByVal pkZipExecutable As String, ByVal unzipWaitTimeout As Integer)

        'Extract zip into temp file
        Dim processExe As String = pkZipExecutable
        Dim args As String = " -add " & zipFilename & " " & inputFolder & "*.*"
        Dim proc As Diagnostics.Process = Nothing
        proc = Diagnostics.Process.Start(processExe, args)

        Dim success As Boolean = False
        success = proc.WaitForExit(unzipWaitTimeout)

        If success = False Then Throw New BMSL.FileIOException("Unzip wait timeout exceeded.")

    End Sub

    Public Sub ImportZipContentsAsPDF(ByVal zipFilename As String, ByVal foreignIDType As String, ByVal documentType As Integer, _
        ByVal inputFolder As String, ByVal pkZipExecutable As String, ByVal unzipWaitTimeout As Integer)

        Dim filenameNoPath As String = ""
        Dim filepath As String = ""
        Dim filenameNoExtension As String = ""
        Dim filenameExtension As String = ""

        'Split filename by \
        Dim filenameArray1 As String() = zipFilename.Split("\")
        filenameNoPath = filenameArray1(filenameArray1.Length - 1)
        filepath = Left(zipFilename, Len(zipFilename) - Len(filenameNoPath)) 'keep last slash

        'Split filename by .
        Dim filenameArray2 As String() = filenameNoPath.Split(".")
        filenameExtension = filenameArray2(filenameArray2.Length - 1)
        filenameNoExtension = Left(filenameNoPath, Len(filenameNoPath) - Len(filenameExtension) - 1) 'remove period

        'Create extract folder
        Dim extractFolder As String = inputFolder & "UNPACKED_" & filenameNoExtension.Replace(".", "_")

        If IO.Directory.Exists(extractFolder) Then
            IO.Directory.Move(extractFolder, extractFolder & "[" & Format(Now(), "yyyyMMddHHmmss") & "]")
        End If

        IO.Directory.CreateDirectory(extractFolder)

        'Extract zip into temp file
        Dim processExe As String = pkZipExecutable
        Dim args As String = " -extract " & zipFilename & " " & extractFolder
        Dim proc As Diagnostics.Process = Nothing
        proc = Diagnostics.Process.Start(processExe, args)

        Dim success As Boolean = False
        success = proc.WaitForExit(unzipWaitTimeout)

        If success = False Then Throw New BMSL.FileIOException("Unzip wait timeout exceeded.")

        'Insert documents extracted
        InsertBinaryDocumentsInFolder(extractFolder, foreignIDType, documentType)

        'Delete unpack folder
        'TODO
    End Sub

End Class
