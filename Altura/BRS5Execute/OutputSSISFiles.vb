Public Class OutputSSISFiles

    Dim _logger As BMSL.LogControl = Nothing

    Public Sub New(ByVal databaseControl As BMSL.DatabaseControl, ByVal logger As BMSL.LogControl)
        _logger = logger

    End Sub

    Public Sub Execute(ByVal clientCode As String)

        Dim packageFolder As String = LocalConstants.PACKAGE_FOLDER

        Dim package1Filename As String = packageFolder & "Attorney.dtsx" 'c:\brs5\packages\Attorney.dtsx
        Dim package2Filename As String = packageFolder & "Court.dtsx"
        Dim package3Filename As String = packageFolder & "Debtor_Address.dtsx"
        Dim package4Filename As String = packageFolder & "Debtor_Name.dtsx"
        Dim package5Filename As String = packageFolder & "Debtor_SSN_TIN.dtsx"
        Dim package6Filename As String = packageFolder & "Job_Control.dtsx"
        Dim package7Filename As String = packageFolder & "Primary_Case.dtsx"
        Dim package8Filename As String = packageFolder & "Trustee.dtsx"
        Dim package9Filename As String = packageFolder & "Staging_Hold.dtsx"
        Dim package10Filename As String = packageFolder & "Manifest.dtsx"

        Dim outputFolder As String = LocalConstants.OUTPUT_FOLDER '"C:\BRS5\Output\"
        Dim deliveryFolder As String = LocalConstants.DELIVERY_FOLDER '"\\Jenkins\Altura\Output\BRS5\"

        Dim output1Filename As String = "Attorney.txt"
        Dim output2Filename As String = "Court.txt"
        Dim output3Filename As String = "Debtor_Address.txt"
        Dim output4Filename As String = "Debtor_Name.txt"
        Dim output5Filename As String = "Debtor_SSN_TIN.txt"
        Dim output6Filename As String = "Job_Control.txt"
        Dim output7Filename As String = "Primary_Case.txt"
        Dim output8Filename As String = "Trustee.txt"
        Dim output9Filename As String = "Staging_Hold.txt"
        Dim output10Filename As String = "Manifest.txt"

        Dim finalOutput1Filename As String = ""
        Dim finalOutput2Filename As String = ""
        Dim finalOutput3Filename As String = ""
        Dim finalOutput4Filename As String = ""
        Dim finalOutput5Filename As String = ""
        Dim finalOutput6Filename As String = ""
        Dim finalOutput7Filename As String = ""
        Dim finalOutput8Filename As String = ""
        Dim finalOutput9Filename As String = ""
        Dim finalOutput10Filename As String = ""

        finalOutput1Filename = OutputFile(clientCode, outputFolder, output1Filename, package1Filename)
        finalOutput2Filename = OutputFile(clientCode, outputFolder, output2Filename, package2Filename)
        finalOutput3Filename = OutputFile(clientCode, outputFolder, output3Filename, package3Filename)
        finalOutput4Filename = OutputFile(clientCode, outputFolder, output4Filename, package4Filename)
        finalOutput5Filename = OutputFile(clientCode, outputFolder, output5Filename, package5Filename)
        finalOutput6Filename = OutputFile(clientCode, outputFolder, output6Filename, package6Filename)
        finalOutput7Filename = OutputFile(clientCode, outputFolder, output7Filename, package7Filename)
        finalOutput8Filename = OutputFile(clientCode, outputFolder, output8Filename, package8Filename)
        finalOutput9Filename = OutputFile(clientCode, outputFolder, output9Filename, package9Filename)
        finalOutput10Filename = OutputFile(clientCode, outputFolder, output10Filename, package10Filename)

        Dim moveFilenames(9) As String
        moveFilenames(0) = finalOutput1Filename
        moveFilenames(1) = finalOutput2Filename
        moveFilenames(2) = finalOutput3Filename
        moveFilenames(3) = finalOutput4Filename
        moveFilenames(4) = finalOutput5Filename
        moveFilenames(5) = finalOutput6Filename
        moveFilenames(6) = finalOutput7Filename
        moveFilenames(7) = finalOutput8Filename
        moveFilenames(8) = finalOutput9Filename
        moveFilenames(9) = finalOutput10Filename

        MoveFiles(moveFilenames, outputFolder, deliveryFolder)


    End Sub

    Public Sub resolveFileConflict(ByVal filename As String)
        _logger.LogInfo("Resolving duplicate file conflict: " & filename)

        System.IO.File.Move(filename, filename & ".renamed." & Format(Now(), "yyyyMMddHHmmss"))

    End Sub

    Public Function OutputFile(ByVal clientCode As String, ByVal folder As String, ByVal shortFilename As String, ByVal ssisPackageName As String) As String

        'check if file already exists
        If System.IO.File.Exists(folder & shortFilename) Then
            'System.IO.File.Move(folder & shortFilename, folder & shortFilename & ".renamed." & Format(Now(), "yyyyMMddHHmmss"))
            resolveFileConflict(folder & shortFilename)
        End If

        'run SSIS package
        _logger.LogInfo("Running SSIS package: " & ssisPackageName)

        Dim p As Process = System.Diagnostics.Process.Start("C:\Program Files\Microsoft SQL Server\100\DTS\Binn\DTEXEC", "/F " & ssisPackageName)

        p.WaitForExit(600000)
        'check if file was output
        If System.IO.File.Exists(folder & shortFilename) Then
            'OK, file was output
        Else
            Throw New System.Exception("SSIS Package [" & ssisPackageName & "] did not output file [" & folder & shortFilename & "]")

        End If

        'rename file
        Dim renamedShortFilename As String = ""

        'check if client code is used
        Dim prefix As String = ""
        If Len(clientCode) > 0 Then
            prefix = "LCI_to_" & clientCode & "_"
        End If

        'split off extension
        Dim shortFilename1 As String() = shortFilename.Split(".")

        renamedShortFilename = prefix & shortFilename1(0) & "_" & Format(Now(), "yyyyMMdd") & "." & shortFilename1(1)

        _logger.LogInfo("Renaming output file: " & renamedShortFilename)

        'check if renamed file already exists
        If System.IO.File.Exists(folder & renamedShortFilename) Then
            'System.IO.File.Move(folder & renamedShortFilename, folder & renamedShortFilename & ".renamed." & Format(Now(), "yyyyMMddHHmmss"))
            resolveFileConflict(folder & renamedShortFilename)
        End If

        'Rename file
        System.IO.File.Move(folder & shortFilename, folder & renamedShortFilename)

        'send the name back to the caller
        Return renamedShortFilename

    End Function

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

End Class
