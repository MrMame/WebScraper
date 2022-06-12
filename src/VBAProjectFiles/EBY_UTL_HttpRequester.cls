VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_UTL_HttpRequester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements EBY_DBG_LOG_ILoggable

Dim m_logger As EBY_DBG_LOG_ILogger


Public Function GetHTMLDocument(strAdress As String) As HTMLDocument
 
    Dim Html As HTMLDocument, hTable As HTMLTable
    Set Html = New HTMLDocument
        
    Dim strGetRequest As String
    strGetRequest = strAdress
    
    On Error GoTo RequestError
    Dim oXMLHTTP As New MSXML2.XMLHTTP
    With oXMLHTTP
        .Open "GET", strGetRequest, False
        .send
    Html.body.innerHTML = .responseText
    End With
    GoTo finish
    
RequestError:
    EBY_DBG_LOG_ILoggable_WriteToLogger ("[EBY_UTL_HttpRequester] Error Get Request: err.number(" & Err.Number & ") err.description(" & Err.Description & ")")
finish:
    
    EBY_DBG_LOG_ILoggable_WriteToLogger ("[EBY_UTL_HttpRequester] Requested HTMLDocument from address (" & strAdress & "). Length of Responsetext was (" & Len(Html.body.innerHTML) & ") Characters")

    Set GetHTMLDocument = Html
End Function



'   Interface: EBY_DBG_LOG_ILoggable
'
Public Sub EBY_DBG_LOG_ILoggable_SetLogger(Logger As EBY_DBG_LOG_ILogger)
    Set m_logger = Logger
End Sub
Private Sub EBY_DBG_LOG_ILoggable_WriteToLogger(Text As String)
    If (Not m_logger Is Nothing) Then m_logger.Log (Text)
End Sub

