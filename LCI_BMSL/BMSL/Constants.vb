Public Class Constants

    Public Const INI_FILENAME = "C:\BMS\BIN\Settings.ini"
    
    Public Const USER_LOGIN_SP_AUTH = "SDJF*&^D87s6dMSF*&^SD87fr"

    Public Const WEB_INTERACT_PARAM__USER_ENTITY = "User_Entity"
    Public Const WEB_INTERACT_PARAM__COURT_USER = "Court_User"
    Public Const WEB_INTERACT_PARAM__CLIENT_CODE = "Client_Code"
    Public Const WEB_INTERACT_PARAM__CASE_NUMBER = "Case_Number"
    Public Const WEB_INTERACT_PARAM__OFFICE_CODE = "Office_Code"

    Public Const WEB_INTERACT_PARAM__PARTY_NAME = "Party_Name"
    Public Const WEB_INTERACT_PARAM__PARTY_TAXID = "Party_TaxID"
    Public Const WEB_INTERACT_PARAM__PARTY_ADDRESS1 = "Party_Address1"
    Public Const WEB_INTERACT_PARAM__PARTY_ADDRESS2 = "Party_Address2"
    Public Const WEB_INTERACT_PARAM__PARTY_CITY = "Party_City"
    Public Const WEB_INTERACT_PARAM__PARTY_STATE = "Party_State"
    Public Const WEB_INTERACT_PARAM__PARTY_ZIP = "Party_Zip"
    Public Const WEB_INTERACT_PARAM__PARTY_PHONE = "Party_Phone"
    Public Const WEB_INTERACT_PARAM__PARTY_EMAIL = "Party_Email"

    Public Const WEB_INTERACT_DEFAULT_TIMEOUT = 60000

    Public Const ECF_CASE_TYPE_BK = "bk"
    Public Const ECF_CASE_TYPE_AP = "ap"
    Public Const ECF_CASE_TYPE_MP = "mp"

    Public Const FILE_IO_FLAG_SUCCESS = "Success"
    Public Const FILE_IO_FLAG_REJECT = "Reject"
    Public Const FILE_IO_FLAG_PENDING = "Pending"

    Public Const FILE_IO_METHOD_ECF = "ECF"
    Public Const FILE_IO_METHOD_MAIL = "Mail"
    Public Const FILE_IO_METHOD_CERT = "Cert"

    'STATUSES

    Public Enum WebInteractStatus
        Pending = 511001
        Running = 511002
        Paused = 511003
        Canceled = 511004
        Interrupted = 511005
        Complete = 511010
        CaseNotFound = 511011
        MultipleCasesFound = 511012

        CourtUnavailable = 511020

        MultipleNonFatalStatuses = 511030
        MultipleFatalNonFatalStatuses = 511031
        MultipleFatalStatus = 511032
        HTMLLoginFailed = 511033
        DataIntegrityError = 511034
        CaseQueryError = 511035
        HTMLElementNotFoundError = 511036
        HTMLPageTimedOutError = 511037
        DocumentAccessDenied = 511038

        BuildingDatabaseRow = 511065

        ControlJobInvalid = 511090
        HTMLError = 511095
        UnclassifiedError = 511099

    End Enum

    Public Enum DocumentStatus
        UploadPending = 512001
        UploadComplete = 512010
        BuildingDatabaseRow = 512065
        UploadError = 512099
        DownloadComplete = 512110
        LocalDocument = 512501

    End Enum

    Public Enum DocumentUtilityStatus
        Pending = 513001
        Running = 513002
        Complete = 513010
        BuildingDatabaseRow = 513065
        UnclassifiedError = 513099

    End Enum

    Public Enum FileIOStatus
        Pending = 514001
        Running = 514002
        Complete = 514010
        CompleteWithErrors = 514011
        FileNotFound = 514021
        FileOperationsError = 514022
        DataIntegrityError = 514023
        BuildingDatabaseRow = 514065
        UnclassifiedError = 514099

    End Enum

    Public Enum FileBatchStatus
        Pending = 515001
        Running = 515002
        Complete = 515010
        CompleteNonFatalErrors = 515011
        CompleteFatalErrors = 515020
        BuildingDatabaseRow = 515065
        UnclassifiedError = 515099

        Pending_FileImport = 515101
        Running_FileImport = 515102
        Complete_FileImport = 515110
        Error_FileImport = 515199

        Pending_ReportGeneration = 515201
        Running_ReportGeneration = 515202
        Complete_ReportGeneration = 515210
        Error_ReportGeneration = 515299

        Pending_WebInteract = 515301
        Running_WebInteract = 515302
        Complete_WebInteract = 515310
        Error_WebInteract = 515399

        Pending_PaperProcessing = 515401
        Running_PaperProcessing = 515402
        Complete_PaperProcessing = 515410
        Error_PaperProcessing = 515499

        Pending_FileExport = 515501
        Running_FileExport = 515502
        Complete_FileExport = 515510
        Error_FileExport = 515599

        Pending_DocumentMerge = 515601
        Running_DocumentMerge = 515602
        Complete_DocumentMerge = 515610
        Error_DocumentMerge = 515699

    End Enum

    Public Enum FileDataStatus
        Pending = 516001
        Running = 516002
        Complete = 516010

        UnclassifiedDataError = 516199
        DataValidationFailed = 516191
        DataMatchFailed = 516192

        ReportGenerationComplete = 516210
        UnclassifiedReportError = 516299

        CourtInteractionComplete = 516310
        UnclassifiedCourtError = 516399

        PaperProcessingComplete = 516410
        UnclassifiedPaperProcessingError = 516499

        BuildingDatabaseRow = 516065
        UnclassifiedError = 516099

    End Enum

    Public Enum ReportBatchStatus
        Pending = 517001
        Running = 517002
        Complete = 517010
        CompleteNonFatalErrors = 517011
        CompleteFatalErrors = 517012
        BuildingDatabaseRow = 517065
        UnclassifiedError = 517099

    End Enum

    Public Enum ReportDataStatus
        Pending = 518001
        Running = 518002
        Complete = 518010
        MissingRequiredFields = 518021
        DataIntegrityError = 518022
        PDFGenerationError = 518023
        BuildingDatabaseRow = 518065
        UnclassifiedError = 518099

    End Enum

    Public Enum FileReturnStatus
        Pending = 519001
        Running = 519002
        Complete = 519010
        UnclassifiedError = 519099

    End Enum

    Public Enum PaperProcessingStatus
        Pending = 520001
        Running = 520002
        Complete = 520010
        UnclassifiedError = 520099

    End Enum

    Public Enum PaperBatchStatus
        Pending = 521001
        Running = 521002
        Complete = 521010
        UnclassifiedError = 521099

    End Enum

    Public Enum ResubmitBatchStatus
        Pending = 522001
        Running = 522002
        Complete = 522010
        UnclassifiedError = 522099

    End Enum


    'TYPES

    Public Enum CourtInteractionType

        ECFLogin = 11001
        ECFLogout = 11002
        ECFUploadClaim = 11101
        ECFUploadWithdrawalOfClaim = 11102
        ECFUploadTransferOfClaim = 11103
        ECFUploadNoticeOfPaymentChange = 11104
        ECFDownloadCourtFacesheet = 11201
        ECFDownloadDocket = 11202
        ECFDownloadCreditorMatrix = 11203
        ECFDownloadCreditors = 11204
        ECFDownloadAbstractCreditor = 11205
        ECFDownloadClaimsRegister = 11206
        ECFDownloadSubDocket = 11207

    End Enum

    Public Enum CourtType
        ECFLive = 12001
        ECFTrain = 12002
        ECFText = 12003
        PacerLive = 12004

    End Enum

    Public Enum CourtUserLoginType
        ECFPacer = 13001
        ECFFileLimited = 13002
        ECFFileAttorney = 13003
        TrainingECFPacer = 13011
        TrainingECFFileLimited = 13012
        TrainingECFFileAttorney = 13013

    End Enum

    Public Enum DocumentType
        DocketHTML = 14001
        CreditorMatrixHTML = 14002
        CreditorsHTML = 14003
        ClaimsRegisterHTML = 14004
        CourtFaceSheetHTML = 14005
        CourtFilingNEFReceiptHTML = 14006
        SecondCourtFilingNEFReceiptHTML = 14016
        OpenCasesReportHTML = 14007
        DismissedCasesReportHTML = 14008
        DischargedCasesReportHTML = 14009
        TransferredCasesReportHTML = 14010
        ClosedCasesReportHTML = 14011
        CourtTransactionReceiptHTML = 14012
        CourtErrorDocumentHTML = 14013
        CourtAuditDocumentHTML = 14014

        DocketText = 14101
        CreditorMatrixText = 14102
        CreditorsText = 14103
        ClaimsRegisterText = 14104
        CourtFaceSheetText = 14105
        CourtFilingNEFReceiptText = 14106
        SecondCourtFilingNEFReceiptText = 14116
        OpenCasesReportText = 14107
        DismissedCasesReportText = 14108
        DischargedCasesReportText = 14109
        TransferredCasesReportText = 14110
        ClosedCasesReportText = 14111
        CourtTransactionReceiptText = 14112
        CourtErrorDocumentText = 14113
        CourtAuditDocumentText = 14114

        SubDocketPDF = 14201
        CourtFilingPOCB10PDF = 14202
        CourtFilingTOC210PDF = 14203
        CourtFilingAttachmentPDF = 14204
        CourtFilingStatementPDF = 14205
        CourtFilingNPCFormPDF = 14206
        CourtFilingNPCLetterPDF = 14207
        CourtFilingNPCFormAndLetterPDF = 14208
        CourtFilingNPCFormPDF_LAM_1 = 14209
        CourtFilingNPCFormPDF_LAM_2 = 14210

        MailingLabelPDFDebtor = 14301
        MailingLabelPDFDebtor2 = 14311
        MailingLabelPDFAttorney = 14302
        MailingLabelPDFCourt = 14303
        MailingLabelPDFTrustee = 14304


    End Enum

    Public Enum DocumentUtilityType
        ImportDocument = 15001
        ExportDocument = 15002

    End Enum

    Public Enum WebInteractType
        DownloadDocket_SingleCase_New = 16001
        DownloadDocket_MultipleCases_New = 16002
        DownloadDocket_OpenEnded_New = 16003
        DownloadDocket_SingleCase_Update = 16004
        DownloadDocket_MultipleCases_Update = 16005
        DownloadDocket_HotUpdate = 16006
        DownloadDocket_SingleCase_ExternalSystem = 16007
        DownloadDocket_MultipleCases_ExternalSystem = 16008
        DownloadCreditorMatrix = 16009
        DownloadCreditors = 16010
        DownloadAbstractCreditor = 16011
        DownloadClaimsRegister = 16012
        DownloadSummary_OpenCases = 16013
        DownloadSummary_DismissedCases = 16014
        DownloadSummary_DischargedCases = 16015
        DownloadSummary_ConvertedCases = 16016
        DownloadSummary_ClosedCases = 16017
        DownloadSubDocket_SingleDocument = 16018
        DownloadSubDocket_MultipleDocuments = 16019
        DownloadCourtFacesheet = 16020

        UploadProofOfClaim = 16101
        UploadWithdrawalOfClaim = 16102
        UploadTransferofClaim = 16103
        UploadNoticeOfPaymentChange = 16104

        UploadProofOfClaim_Test = 16151
        UploadWithdrawalOfClaim_Test = 16152
        UploadTransferofClaim_Test = 16153
        UploadNoticeOfPaymentChange_Test = 16154

        Download_Train_Docket_SingleCase_New = 16201
        Download_Train_Docket_MultipleCases_New = 16202
        Download_Train_Docket_OpenEnded_New = 16203
        Download_Train_Docket_SingleCase_Update = 16204
        Download_Train_Docket_MultipleCases_Update = 16205
        Download_Train_Docket_HotUpdate = 16206
        Download_Train_Docket_SingleCase_ExternalSystem = 16207
        Download_Train_Docket_MultipleCases_ExternalSystem = 16208
        Download_Train_CreditorMatrix = 16209
        Download_Train_Creditors = 16210
        Download_Train_AbstractCreditor = 16211
        Download_Train_ClaimsRegister = 16212
        Download_Train_Summary_OpenCases = 16213
        Download_Train_Summary_DismissedCases = 16214
        Download_Train_Summary_DischargedCases = 16215
        Download_Train_Summary_ConvertedCases = 16216
        Download_Train_Summary_ClosedCases = 16217
        Download_Train_SubDocket_SingleDocument = 16218
        Download_Train_SubDocket_MultipleDocuments = 16219
        Download_Train_CourtFacesheet = 16220

        Upload_Train_ProofOfClaim = 16301
        Upload_Train_WithdrawalOfClaim = 16302
        Upload_Train_TransferofClaim = 16303
        Upload_Train_NoticeOfPaymentChange = 16304

        Upload_Train_ProofOfClaim_Test = 16351
        Upload_Train_WithdrawalOfClaim_Test = 16352
        Upload_Train_TransferofClaim_Test = 16353
        Upload_Train_NoticeOfPaymentChange_Test = 16354

    End Enum

    Public Enum WebInteractClientGroup
        ECFPuller = 17001
        ECFFiler = 17002
        ECFPullerOrFiler = 17003

    End Enum

    Public Enum WebInteractTypeGroup
        ECFMultipleCourtLoginCapable = 18001
        ECFSingleCourtLoginCapable = 18002

    End Enum

    Public Enum HTMLTagType
        VerifyHTMLEndTag = 19001

    End Enum

    Public Enum DocumentSource
        PACER = 20001
        BMS = 20002
        ExternalSource = 20003

    End Enum

    Public Enum FileIOType
        Define_FilePrimaryRow = 21001
        Define_FileTrailerRow = 21002
        Define_FileHeaderRow = 21003

        'Ver 1.0
        NBS10_Input_NoticeOfPaymentChange = 21101
        NBS10_InputTrailer_NoticeOfPaymentChange = 21102
        NBS10_InputBinaryDocs_NoticeOfPaymentChange = 21151

        NBS10_Output_NoticeOfPaymentChange = 21201
        NBS10_OutputTrailer_NoticeOfPaymentChange = 21202
        NBS10_OutputBinaryDocs_NoticeOfPaymentChange = 21251

        'Ver 1.1
        NBS11_Input_NoticeOfPaymentChange = 21301
        NBS11_InputTrailer_NoticeOfPaymentChange = 21302
        NBS11_InputBinaryDocs_NoticeOfPaymentChange = 21351

        NBS11_Output_NoticeOfPaymentChange = 21401
        NBS11_OutputTrailer_NoticeOfPaymentChange = 21402
        NBS11_OutputBinaryDocs_NoticeOfPaymentChange = 21451

        'Ver 1.2
        NBS12_Input_NoticeOfPaymentChange = 21501
        NBS12_InputTrailer_NoticeOfPaymentChange = 21502
        NBS12_InputBinaryDocs_NoticeOfPaymentChange = 21551

        NBS12_Output_NoticeOfPaymentChange = 21601
        NBS12_OutputTrailer_NoticeOfPaymentChange = 21602
        NBS12_OutputBinaryDocs_NoticeOfPaymentChange = 21651

    End Enum

    Public Enum FileBatchType
        NoticeOfPaymentChange = 22001

    End Enum

    Public Enum FileDataType
        NPC_ECFPacer = 23001
        NPC_Mail = 23002
        NPC_CertifiedMail = 23003
        NPC_Signature = 23004

    End Enum

    Public Enum ReportBatchType
        NBS10_NoticeOfPaymentChange = 24001
        NBS11_NoticeOfPaymentChange = 24002
        NBS12_NoticeOfPaymentChange = 24003

    End Enum

    Public Enum ReportDataType
        NBS10_NoticeOfPaymentChange = 25001
        NBS11_NoticeOfPaymentChange = 25002
        NBS12_NoticeOfPaymentChange = 25003

    End Enum

    Public Enum BatchProductionType
        Live = 26111
        Test = 26222

    End Enum

    Public Enum FileQueueType
        NBS_InputPDF_NoticeOfPaymentChange = 27101

        NBS_OutputPDF_NoticeOfPaymentChange = 27201


    End Enum

    Public Enum FileDataPartyType
        Mortgagor = 28001
        CoMortgagor = 28002
        Debtor = 28003
        CoDebtor = 28004
        DebtorAttorney = 28005
        Trustee = 28006
        Court = 28007
        Creditor = 28008

    End Enum

    Public Enum FileIODirectionType
        Input = 29001
        Output = 29002

    End Enum

    Public Enum FileIODelimiterType
        SpaceDelimited = 30001
        PipeDelimited = 30002

    End Enum

    Public Enum PaperProcessingType
        NPC = 31001
    End Enum

    Public Enum PaperBatchType
        NPC = 32001
    End Enum

    Public Enum MailingDataType
        'Debtor1
        Debtor = 33101
        DebtorCertified = 33102

        'Debtor2        
        Debtor2 = 33111
        Debtor2Certified = 33112

        DebtorAttorney = 33201
        Court = 33301
        CourtSASE = 33302
        CourtBatched = 33303
        CourtBatchedSASE = 33304
        Trustee = 33401
        TrusteeBatched = 33402
        TrusteeCertified = 33403
        TrusteeBatchedCertified = 33404
        OtherParties = 33501

    End Enum

    Public Enum ResubmitBatchType
        ResubmitNPC = 34001

    End Enum
End Class
