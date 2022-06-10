VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_DBG_LOG_TableLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILogger

Private Const STRING_SHEETNAME_LOG_INIT As String = "Log"


Dim m_shtTarget As Worksheet

Private Sub Class_Initialize()
    Set m_shtTarget = GetSheet(STRING_SHEETNAME_LOG_INIT, False)
End Sub


Public Sub ILogger_Log(text As String)
    With m_shtTarget.Cells(m_shtTarget.Rows.Count, 1).End(xlUp)
        .Offset(1, 0) = Now
        .Offset(2, 0) = text
    End With
End Sub


Private Function GetSheet(sheetname As String, Optional CleanSheet As Boolean = True) As Worksheet

    Dim retSheet As Worksheet

    On Error GoTo noSheet
    Set retSheet = ActiveWorkbook.Sheets(sheetname)
    If (CleanSheet = True) Then
        Application.DisplayAlerts = False
        retSheet.Delete
        Application.DisplayAlerts = True
        Set retSheet = ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        retSheet.Name = sheetname
    End If
    GoTo endThis

noSheet:
    Set retSheet = ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    retSheet.Name = sheetname
endThis:
    Set GetSheet = retSheet
End Function
