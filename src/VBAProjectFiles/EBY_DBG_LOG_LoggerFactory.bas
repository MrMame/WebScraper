Attribute VB_Name = "EBY_DBG_LOG_LoggerFactory"
Option Explicit

Public Function CreateFullMultiLogger() As EBY_DBG_LOG_MultiLogger
    Dim theMultiLogger As New EBY_DBG_LOG_MultiLogger
    With theMultiLogger
        Call .AddLogger(CreateStandardTableLogger)
        Call .AddLogger(CreateWorkbookFolderFileLogger)
    End With
    Set CreateFullMultiLogger = theMultiLogger
End Function


Public Function CreateWorkbookFolderFileLogger() As EBY_DBG_LOG_FileLogger
    Dim strTargetFilename As String
    strTargetFilename = ActiveWorkbook.Path & "\EbayScraper.log"
    Dim theFileLogger As New EBY_DBG_LOG_FileLogger
    theFileLogger.Logfilename = strTargetFilename
    Set CreateWorkbookFolderFileLogger = theFileLogger
End Function


Public Function CreateStandardTableLogger() As EBY_DBG_LOG_TableLogger
    Dim theTableLogger As New EBY_DBG_LOG_TableLogger
    Set CreateStandardTableLogger = theTableLogger
End Function
