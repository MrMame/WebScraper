VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_UTL_TableWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements EBY_DBG_LOG_ILoggable


Dim m_logger As EBY_DBG_LOG_ILogger


Private Const STRING_HEADERNAME_DATE As String = "Datum"
Private Const STRING_HEADERNAME_LOCATION As String = "Ort"
Private Const STRING_HEADERNAME_PRICE As String = "Preis"
Private Const STRING_HEADERNAME_NAME As String = "Name"
Private Const STRING_TABLENAME_ADS As String = "Tab_Ads"
Private Const STRING_HEADERNAME_NEGOTIABLE As String = "Verhandelbar"
Private Const STRING_TABLE_VALUE_NEGOTIABLE As String = "X"

Public Sub WriteAds(Ads As Collection, shtTarget As Worksheet)
    
    If (Ads Is Nothing) Then
        Call Err.Raise(91, Description:="Ads Collection is nothing!")
    ElseIf (shtTarget Is Nothing) Then
        Call Err.Raise(91, Description:="shtTargetis nothing!")
    End If
      
    Application.ScreenUpdating = False
    
    ' Write Header and the Columns Format
    With shtTarget.Range("a1")
        .Offset(0, 0) = STRING_HEADERNAME_DATE
        .Offset(0, 0).EntireColumn.NumberFormat = "m/d/yyyy"
        .Offset(0, 1) = STRING_HEADERNAME_LOCATION
        .Offset(0, 2) = STRING_HEADERNAME_PRICE
        .Offset(0, 2).EntireColumn.NumberFormat = "_-* #,##0.00 [$€-407]_-;-* #,##0.00 [$€-407]_-;_-* ""-""?? [$€-407]_-;_-@_-"
        .Offset(0, 3) = STRING_HEADERNAME_NEGOTIABLE
        .Offset(0, 4) = STRING_HEADERNAME_NAME
    End With
 
    
    ' Write the Data
    Dim Ad As EBY_DAT_Ad
    For Each Ad In Ads
        With shtTarget.Cells(shtTarget.Rows.Count, 1).End(xlUp).Offset(1, 0)
            .Offset(0, 0) = Ad.AdDate
            .Offset(0, 1) = Ad.Location
            .Offset(0, 2) = Ad.Price
            If (Ad.Negotiable) Then
                .Offset(0, 3) = STRING_TABLE_VALUE_NEGOTIABLE
            End If
            
            shtTarget.Hyperlinks.Add Anchor:=.Offset(0, 4), _
            Address:=Ad.LinkAddress, _
            TextToDisplay:=Ad.AdName
    
        End With
    
    Next

    ' Create the Table
    shtTarget.ListObjects.Add(xlSrcRange, shtTarget.UsedRange, , xlYes).Name = _
        STRING_TABLENAME_ADS
     Cells.EntireColumn.AutoFit

    Application.ScreenUpdating = True
    
    
    EBY_DBG_LOG_ILoggable_WriteToLogger ("Wrote (" & Ads.Count & ") Ad's into Targetsheet (" & shtTarget.Name & ")")
    
End Sub

'Private Function GetSheet(sheetname As String, Optional CleanSheet As Boolean = True) As Worksheet
'
'    Dim retSheet As Worksheet
'
'    On Error GoTo noSheet
'    Set retSheet = ActiveWorkbook.Sheets(sheetname)
'    If (CleanSheet = True) Then
'        Application.DisplayAlerts = False
'        retSheet.Delete
'        Application.DisplayAlerts = True
'        Set retSheet = ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
'        retSheet.Name = sheetname
'    End If
'    GoTo endThis
'
'noSheet:
'    Set retSheet = ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
'    retSheet.Name = sheetname
'endThis:
'    Set GetSheet = retSheet
'End Function




'   Interface: EBY_DBG_LOG_ILoggable
'
Public Sub EBY_DBG_LOG_ILoggable_SetLogger(Logger As EBY_DBG_LOG_ILogger)
    Set m_logger = Logger
End Sub
Private Sub EBY_DBG_LOG_ILoggable_WriteToLogger(Text As String)
    If (Not m_logger Is Nothing) Then m_logger.Log (Text)
End Sub

