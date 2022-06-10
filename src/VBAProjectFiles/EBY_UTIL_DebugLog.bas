Attribute VB_Name = "EBY_UTIL_DebugLog"
Option Explicit

Public Sub LogToFile(Filename As String, text As String)

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = FSO.CreateTextFile(Filename)
    oFile.WriteLine text
    oFile.Close
    Set FSO = Nothing
    Set oFile = Nothing
    
End Sub




