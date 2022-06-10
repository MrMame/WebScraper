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

Implements ILogger

Private Const STRING_FILENAME_INIT As String = "DebugLog.log"

Dim m_FSO As Object
Dim m_oFile As Object

Private m_Filename As String


Private Sub Class_Initialize()
    m_Filename = STRING_FILENAME_INIT
    Set m_FSO = CreateObject("Scripting.FileSystemObject")
End Sub


Public Sub ILogger_Log(text As String)
    Set oFile = FSO.CreateTextFile(m_Filename)
    m_oFile.WriteLine text
    m_oFile.Close
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    m_oFile.Close
   
    Set m_FSO = Nothing
    Set m_oFile = Nothing
    
End Sub
