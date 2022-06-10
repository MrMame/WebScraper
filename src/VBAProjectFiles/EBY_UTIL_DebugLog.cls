VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_UTIL_DebugLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub LogToFile(Filename As String, Text As String)

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = FSO.CreateTextFile(Filename)
    oFile.WriteLine Text
    oFile.Close
    Set FSO = Nothing
    Set oFile = Nothing
    
End Sub





