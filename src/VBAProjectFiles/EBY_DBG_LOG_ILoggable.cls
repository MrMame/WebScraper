VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_DBG_LOG_ILoggable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   Interface
'
'       This Interface needs to be implemented by all Objects with logging capabilities.
'


Option Explicit

Public Sub SetLogger(Logger As EBY_DBG_LOG_ILogger): End Sub

Public Sub WriteToLogger(Text As String): End Sub
