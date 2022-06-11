VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_DBG_LOG_FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements EBY_DBG_LOG_ILogger

Private Const STRING_FILENAME_INIT As String = "DebugLog.log"

Dim m_FSO As Object
Dim m_oFile As Object

Private m_Filename As String



Public Property Let Logfilename(value As String)
    m_Filename = value
End Property
Public Property Get Logfilename() As String
    Logfilename = m_Filename
End Property




Private Sub Class_Initialize()
    m_Filename = STRING_FILENAME_INIT
    Set m_FSO = CreateObject("Scripting.FileSystemObject")
End Sub


Public Sub EBY_DBG_LOG_ILogger_Log(Text As String)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
     
    On Error GoTo writeError
    Set m_oFile = m_FSO.OpenTextFile(m_Filename, ForAppending, True, TristateFalse)
    m_oFile.WriteLine Now & vbTab & Text
    m_oFile.Close
    GoTo finish
    
writeError:
    Err.Raise "Error accessing logfile." & Err.Number
finish:
    
        
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    m_oFile.Close
    Set m_FSO = Nothing
    Set m_oFile = Nothing
End Sub
