VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_DBG_LOG_MultiLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements EBY_DBG_LOG_ILogger

Private m_colLogger As New Collection


Public Sub AddLogger(Logger As EBY_DBG_LOG_ILogger)
    Call m_colLogger.Add(Logger)
End Sub

Public Sub EBY_DBG_LOG_ILogger_Log(Text As String)
    Dim lg As EBY_DBG_LOG_ILogger
    For Each lg In m_colLogger
        Call lg.Log(Text)
    Next
End Sub

