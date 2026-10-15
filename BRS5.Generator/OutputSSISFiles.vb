Imports Microsoft.SqlServer.Dts.Runtime
Imports System.Text

Public Class OutputSSISFiles

    Const STATUS_COMPLETE As Integer = 10
    Const STATUS_ERROR As Integer = 99

    Dim _logger As BMSL.LogControl = Nothing
    Dim _packageFolder As String = String.Empty
    Dim _outputFolder As String = String.Empty
    Dim _deliveryFolder As String = String.Empty
    Dim _OnHoldFolder As String = String.Empty

    ' Added - CJ
    Dim _databaseControl As BMSL.DatabaseControl = Nothing

    Public Sub New(ByVal databaseControl As BMSL.DatabaseControl, ByVal logger As BMSL.LogControl, ByVal packageFolder As String, ByVal outputFolder As String, ByVal deliveryFolder As String, ByVal onHoldFolder As String)
        _logger = logger

        ' Added - CJ
        _databaseControl = databaseControl 'New BMSL.DatabaseControl(_logger, LocalConstants.SQL_CONNECTION_STRING) 'REMOVED NW 20120924

        'REMOVED NW 20120924
        '' Added - CJ
        'packageFolder = LocalConstants.PACKAGE_FOLDER
        'outputFolder = LocalConstants.OUTPUT_FOLDER
        'deliveryFolder = LocalConstants.DELIVERY_FOLDER '"\\Jenkins\Altura\Output\BRS5\"
        'onHoldFolder = LocalConstants.ON_HOLD_FOLDER

        _packageFolder = packageFolder
        _outputFolder = outputFolder
        _deliveryFolder = deliveryFolder
        _OnHoldFolder = onHoldFolder

    End Sub


    Public Sub GenerateAllSSIS_exceptManifest()
        GenerateSSIS("Primary_Case.txt", "Primary_Case.dtsx")
        GenerateSSIS("Attorney.txt", "Attorney.dtsx")
        ' moved BVM into Batch Type 4 processing
        ' GenerateSSIS("BVM.txt", "BVM.dtsx")
        GenerateSSIS("Case_Attorney_History.txt", "Case_Attorney_History.dtsx")
        GenerateSSIS("Case_Chapter_History.txt", "Case_Chapter_History.dtsx")
        GenerateSSIS("Case_Date_History.txt", "Case_Date_History.dtsx")
        GenerateSSIS("Case_Meeting_341_History.txt", "Case_Meeting_341_History.dtsx")
        GenerateSSIS("Case_Status_History.txt", "Case_Status_History.dtsx")
        GenerateSSIS("Case_Trustee_History.txt", "Case_Trustee_History.dtsx")
        GenerateSSIS("Court.txt", "Court.dtsx")
        GenerateSSIS("Debtor_Address.txt", "Debtor_Address.dtsx")
        GenerateSSIS("Debtor_Address_History.txt", "Debtor_Address_History.dtsx")
        GenerateSSIS("Debtor_Name.txt", "Debtor_Name.dtsx")
        GenerateSSIS("Debtor_SSN_TIN.txt", "Debtor_SSN_TIN.dtsx")
        GenerateSSIS("Job_Control.txt", "Job_Control.dtsx")
        GenerateSSIS("Staging_Hold.txt", "Staging_Hold.dtsx")
        GenerateSSIS("Trustee.txt", "Trustee.dtsx")
        'except manifest
    End Sub

    Public Sub GenerateSSIS(ByVal filename As String, ByVal packageName As String)

        Dim outputFilename As String = ""

        OutputFile(_outputFolder, filename, packageName)

    End Sub

    Public Sub CopyClientSSIS(ByVal clientCode As String, ByVal blnSavetoOnHOLD As Boolean)

        ' get a SQL dataset with file names/delivery flags
        Dim sqlString2 As String = "client.usp_Delivery_Options__Select_Client_Delivery"   ' add batch type for the filter (or usercode)
        Dim sqlData2 As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString2)
        sqlData2.AddParameter("@UserCode", SqlDbType.VarChar, 50, clientCode)

        Dim dsDataset2 As DataSet = sqlData2.SelectDataset

        If dsDataset2.Tables(0).Rows.Count > 0 Then

            For Each dsRow2 As DataRow In dsDataset2.Tables(0).Rows

                'copy everything from SOURCE folder to DEST folder named after clientCode
                'for this client when delivery option = 1

                ' by Nick
                'loop for files
                Dim ssisFile As String = ""

                Dim deliveryFlag As Integer = 0

                If Not dsRow2.IsNull("Delivery_Option_Description") Then ssisFile = dsRow2.Item("Delivery_Option_Description")
                If Not dsRow2.IsNull("Delivery_Flag") Then deliveryFlag = dsRow2.Item("Delivery_Flag")

                ' by Nick
                'if file should be delivered is true:
                If deliveryFlag = 1 Then
                    CopySSIS(clientCode, ssisFile, blnSavetoOnHOLD)
                End If

            Next dsRow2

        End If

    End Sub

    Private Sub CopySSIS(ByVal clientCode As String, ByVal SSISname As String, ByVal blnSavetoOnHOLD As Boolean)

        Dim outputFolder As String = ""
        Dim deliveryFolder As String = ""

        'check if folder name needs slash
        If Right(_outputFolder, 1) = "\" Then
            'ok
            outputFolder = _outputFolder
        Else
            'add slash
            outputFolder = _outputFolder & "\"
        End If


        If blnSavetoOnHOLD Then    ' save to an alternate ON-Hold folder which does not get picked up by macro automation

            'check if folder name needs slash
            If Right(_OnHoldFolder, 1) = "\" Then
                'ok
                deliveryFolder = _OnHoldFolder
            Else
                'add slash
                deliveryFolder = _OnHoldFolder & "\"
            End If
        Else

            'check if folder name needs slash
            If Right(_deliveryFolder, 1) = "\" Then
                'ok
                deliveryFolder = _deliveryFolder
            Else
                'add slash
                deliveryFolder = _deliveryFolder & "\"
            End If
        End If


        'copy from output to delivery
        Dim sourceFilename As String = outputFolder & SSISname & ".txt"

        Dim destFilename As String = deliveryFolder & clientCode & "\" & "LCI_to_" & clientCode & "_" & SSISname & "_" & Format(Now(), "yyyyMMdd") & ".txt"
        CopyFile(sourceFilename, destFilename)

    End Sub

    Public Sub DeleteAllSSIS()
        'delete everything from SOURCE folder 
        ' DeleteSSIS("Attorney.txt")
        DeleteSSIS("BVM.txt")
        DeleteSSIS("Attorney.txt")
        DeleteSSIS("Case_Attorney_History.txt")
        DeleteSSIS("Case_Chapter_History.txt")
        DeleteSSIS("Case_Date_History.txt")
        DeleteSSIS("Case_Meeting_341_History.txt")
        DeleteSSIS("Case_Status_History.txt")
        DeleteSSIS("Case_Trustee_History.txt")
        DeleteSSIS("Court.txt")
        DeleteSSIS("Debtor_Address.txt")
        DeleteSSIS("Debtor_Address_History.txt")
        DeleteSSIS("Debtor_Name.txt")
        DeleteSSIS("Debtor_SSN_TIN.txt")
        DeleteSSIS("Job_Control.txt")
        DeleteSSIS("Primary_Case.txt")
        DeleteSSIS("Staging_Hold.txt")
        DeleteSSIS("Trustee.txt")
        'includes manifest
        DeleteSSIS("Manifest.txt")


    End Sub

    Public Sub DeleteSSIS(ByVal shortFilename As String)

        DeleteFile(_outputFolder, shortFilename)

    End Sub

    Public Sub resolveFileConflict(ByVal filename As String)

        _logger.LogInfo("Resolving duplicate file conflict: " & filename)

        System.IO.File.Move(filename, filename & ".renamed." & Format(Now(), "yyyyMMddHHmmss"))

    End Sub

    ''' <summary>
    ''' New SSIS output file function
    ''' </summary>
    ''' <param name="folder"></param>
    ''' <param name="shortFilename"></param>
    ''' <param name="ssisPackageName"></param>
    ''' <remarks></remarks>
    Public Sub OutputFile(ByVal folder As String, ByVal shortFilename As String, ByVal ssisPackageName As String)

        'check if file already exists
        If System.IO.File.Exists(folder & shortFilename) Then
            'System.IO.File.Move(folder & shortFilename, folder & shortFilename & ".renamed." & Format(Now(), "yyyyMMddHHmmss"))
            resolveFileConflict(folder & shortFilename)
        End If

        'run SSIS package
        _logger.LogInfo("Running SSIS package: " & _packageFolder & ssisPackageName)


        'SPD Removed hard coded reference, now using property from local constants.
        'Dim p As Process = System.Diagnostics.Process.Start("C:\Program Files\Microsoft SQL Server\130\DTS\Binn\DTEXEC", "/F " & _packageFolder & ssisPackageName)

        Dim pStartProperties As ProcessStartInfo = New ProcessStartInfo

        pStartProperties.Arguments = "/F " & _packageFolder & ssisPackageName & " /Decrypt shade010"
        pStartProperties.FileName = LocalConstants.DTS_EXECUTABLE_PATH


        Dim p As Process = System.Diagnostics.Process.Start(pStartProperties)

        p.WaitForExit(600000)
        'check if file was output
        If System.IO.File.Exists(folder & shortFilename) Then
            'OK, file was output
            '_logger.LogInfo("File Exists")

        Else
            Throw New System.Exception("SSIS Package [" & ssisPackageName & "] did not output file [" & folder & shortFilename & "]")

        End If

    End Sub

    Public Sub OutputFile_TU_BRS5(ByVal ssisPackageName As String)

        'run SSIS package

        'Altura 90 code is commented out
        '_logger.LogInfo("Running SSIS package: " & _packageFolder & ssisPackageName)

        'Dim pStartProperties As ProcessStartInfo = New ProcessStartInfo
        'pStartProperties.Arguments = "/F " & _packageFolder & ssisPackageName
        'pStartProperties.FileName = LocalConstants.DTS_EXECUTABLE_PATH

        'Dim p As Process = System.Diagnostics.Process.Start(pStartProperties)
        'p.WaitForExit(600000)

        'Altura 90 Below
        '--------------------------------------------------------------
        Dim pkgLocation As String = _packageFolder & ssisPackageName
        Dim sbuilder As StringBuilder = New StringBuilder()

        Dim pkg As Package
        Dim app As Application
        Dim pkgResults As DTSExecResult

        app = New Application()
        pkg = app.LoadPackage(pkgLocation, Nothing)

        pkgResults = pkg.Execute()

        If pkgResults = DTSExecResult.Success Then
            _logger.LogInfo(String.Format("Package {0} ran successfully", pkgLocation))
        Else
            _logger.LogInfo(String.Format("Package {0} failed", pkgLocation))

            For Each dtError As DtsError In pkg.Errors
                sbuilder.Append(dtError.Description.ToString())
            Next
            Throw New Exception(sbuilder.ToString())

        End If

    End Sub

    ' ''' <summary>
    ' ''' Original SSIS OutputFile function that renamed file at the same time
    ' ''' </summary>
    ' ''' <param name="clientCode"></param>
    ' ''' <param name="folder"></param>
    ' ''' <param name="shortFilename"></param>
    ' ''' <param name="ssisPackageName"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function OutputFile(ByVal clientCode As String, ByVal folder As String, ByVal shortFilename As String, ByVal ssisPackageName As String) As String

    '    'check if file already exists
    '    If System.IO.File.Exists(folder & shortFilename) Then
    '        'System.IO.File.Move(folder & shortFilename, folder & shortFilename & ".renamed." & Format(Now(), "yyyyMMddHHmmss"))
    '        resolveFileConflict(folder & shortFilename)
    '    End If

    '    'run SSIS package
    '    _logger.LogInfo("Running SSIS package: " & ssisPackageName)

    '    Dim p As Process = System.Diagnostics.Process.Start("C:\Program Files\Microsoft SQL Server\100\DTS\Binn\DTEXEC", "/F " & ssisPackageName)

    '    p.WaitForExit(600000)
    '    'check if file was output
    '    If System.IO.File.Exists(folder & shortFilename) Then
    '        'OK, file was output
    '    Else
    '        Throw New System.Exception("SSIS Package [" & ssisPackageName & "] did not output file [" & folder & shortFilename & "]")

    '    End If

    '    'rename file
    '    Dim renamedShortFilename As String = ""

    '    'check if client code is used
    '    Dim prefix As String = ""
    '    If Len(clientCode) > 0 Then
    '        prefix = "LCI_to_" & clientCode & "_"
    '    End If

    '    'split off extension
    '    Dim shortFilename1 As String() = shortFilename.Split(".")

    '    renamedShortFilename = prefix & shortFilename1(0) & "_" & Format(Now(), "yyyyMMdd") & "." & shortFilename1(1)

    '    _logger.LogInfo("Renaming output file: " & renamedShortFilename)

    '    'check if renamed file already exists
    '    If System.IO.File.Exists(folder & renamedShortFilename) Then
    '        'System.IO.File.Move(folder & renamedShortFilename, folder & renamedShortFilename & ".renamed." & Format(Now(), "yyyyMMddHHmmss"))
    '        resolveFileConflict(folder & renamedShortFilename)
    '    End If

    '    'Rename file
    '    System.IO.File.Move(folder & shortFilename, folder & renamedShortFilename)

    '    'send the name back to the caller
    '    Return renamedShortFilename

    'End Function


    Public Sub CopyFile(ByVal sourceFilename As String, ByVal destFilename As String)

        If System.IO.File.Exists(sourceFilename) Then
            _logger.LogInfo("Copying file " & sourceFilename & " to " & destFilename)

            If System.IO.File.Exists(destFilename) Then
                'dest file already exists - rename
                resolveFileConflict(destFilename)
            Else
                'dest file doesn't exist - OK
            End If

            'copy file 
            System.IO.File.Copy(sourceFilename, destFilename)
        Else
            'file doesn't exist
            Throw New Exception("Could not file copy file [" & sourceFilename & "] - File doesn't exist")

        End If

    End Sub

    Public Sub DeleteFile(ByVal folder As String, ByVal shortFilename As String)

        If Right(folder, 1) = "\" Then
            'ok
        Else
            'add slash
            folder = folder & "\"
        End If

        If System.IO.File.Exists(folder & shortFilename) Then
            _logger.LogInfo("Deleting file: " & folder & shortFilename)
            System.IO.File.Delete(folder & shortFilename)
        Else
            'file doesn't exist
            _logger.LogWarning("File doesn't exist: " & folder & shortFilename)

        End If

    End Sub

    Public Sub MoveFiles(ByVal shortFilename() As String, ByVal sourceFolder As String, ByVal destFolder As String)

        'create archive folder in source folder
        Dim archiveFolderName As String = Format(Now(), "yyyyMMdd") & "\"

        If Not System.IO.Directory.Exists(sourceFolder & archiveFolderName) Then
            System.IO.Directory.CreateDirectory(sourceFolder & archiveFolderName)
        End If

        For Each file As String In shortFilename
            'check if file exists in dest folder
            If System.IO.File.Exists(destFolder & file) Then resolveFileConflict(destFolder & file)

            _logger.LogInfo("Copying file for delivery: " & destFolder & file)

            'copy to dest folder
            System.IO.File.Copy(sourceFolder & file, destFolder & file)

            'check if file exists in archive folder
            If System.IO.File.Exists(sourceFolder & archiveFolderName & file) Then resolveFileConflict(sourceFolder & archiveFolderName & file)

            'move to archive folder
            System.IO.File.Move(sourceFolder & file, sourceFolder & archiveFolderName & file)

        Next file

    End Sub

    Public Sub insertManifesttoBRS5(ByVal batchControlID As Integer, ByVal clientID As Integer, ByVal UserCode As String)

        _logger.LogInfo("populate Manifest Table for client")

        Dim sqlString As String = "brs5.usp_Manifest__Insert_Client"
        Dim sqlData As BMSL.DatabaseSession = _databaseControl.CreateNewSession(sqlString)
        sqlData.AddParameter("@Batch_Control_ID", SqlDbType.Int, 4, batchControlID)
        sqlData.AddParameter("@Client_ID", SqlDbType.Int, 4, clientID)
        sqlData.AddParameter("@UserCode", SqlDbType.VarChar, 50, UserCode)

        Dim rowsAffected As Integer = sqlData.ExecuteNonQuery()

        _logger.LogInfo(rowsAffected & " row(s) affected")

    End Sub


    ''Public Function GenerateBK_ID_Recon(ByVal rptStartDate As DateTime, ByVal rptEndDate As DateTime, ByVal Customercode As String, ByVal BK_ID_Recon_Filename As String, ByVal Package_Name As String) As Integer

    ''    'check if file already exists
    ''    If System.IO.File.Exists(BK_ID_Recon_Filename) Then
    ''        resolveFileConflict(BK_ID_Recon_Filename)
    ''    End If


    ''    'Check if SSIS package file exists
    ''    If VerifyFileExists(Package_Name) = False Then
    ''        _logger.LogInfo("SSIS Package [" & Package_Name & "] was not found.")
    ''        ' Return STATUS_ERROR
    ''    Else
    ''        '  SSIS package was found
    ''    End If

    ''    'run SSIS package
    ''    _logger.LogInfo("Running SSIS package: " & _packageFolder & Package_Name)

    ''    Return WriteSSISBK_ID_RECON(rptStartDate, rptEndDate, Customercode, BK_ID_Recon_Filename, Package_Name)


    ''End Function
    ''Private Function WriteSSISBK_ID_RECON(ByVal rptStartDate As DateTime, ByVal rptEndDate As DateTime, ByVal CustomerCode As String, ByVal BK_ID_Recon_Filename As String, ByVal package_Name As String) As Integer

    ''    Dim eventListener As New SQLServerEventListener()

    ''    'Dim pkgLocation As String
    ''    Dim pkg As New Package
    ''    Dim app As New Application
    ''    Dim pkgResults As DTSExecResult
    ''    'Dim ex As Exception = Nothing

    ''    Try

    ''        'pkgLocation = IS_Package_Location
    ''        'app.PackagePassword = "5p0ng3b0b"
    ''        pkg = app.LoadPackage(package_Name, eventListener)
    ''        pkg.Variables("FileName").Value = BK_ID_Recon_Filename
    ''        pkg.Variables("rptStartDate").Value = rptStartDate
    ''        pkg.Variables("rptEndDate").Value = rptEndDate
    ''        pkg.Variables("CustomerCode").Value = CustomerCode



    ''        pkgResults = pkg.Execute(Nothing, Nothing, eventListener, Nothing, Nothing)

    ''        _logger.LogInfo("SSIS Package Execution Results: " & pkgResults.ToString())
    ''        If (pkgResults = DTSExecResult.Failure) Then
    ''            _logger.LogInfo("SSIS Package Execution Errors: " & pkgResults.ToString())
    ''            Return STATUS_ERROR

    ''        Else
    ''            Return STATUS_COMPLETE
    ''        End If

    ''    Catch ex As Exception
    ''        _logger.LogInfo("WriteSSISBK_ID_Recon generated an exception: " & ex.Message & ex.StackTrace)

    ''        'Throw ex
    ''        Return STATUS_ERROR
    ''    End Try

    ''End Function

    Private Function VerifyFileExists(ByVal filename As String) As Boolean

        If System.IO.File.Exists(filename) Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
