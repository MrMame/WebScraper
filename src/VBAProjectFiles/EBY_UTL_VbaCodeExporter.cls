VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_UTL_VbaCodeExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'   Example was taken from this website:
'       https://www.rondebruin.nl/win/s9/win002.htm
'
'
'   Add References
'       "Microsoft Visual Basic For Applications Extensibility 5.3"
'       "Microsoft Scripting Runtime"
'   You have to allow access to the VBE Editor in the excel settings.
'


Private Const STRING_TARGETFOLDER_EXPORTFILES As String = "VBAProjectFiles\"

Dim m_wkbSource As Excel.Workbook
Dim m_strExportPath As String
Dim m_boolShowExportMessage As Boolean


Public Property Get Workbook() As Excel.Workbook
    Workbook = m_wkbSource
End Property
Public Property Let Workbook(value As Excel.Workbook)
    m_wkbSource = value
End Property
Public Property Get ExportPath() As String
    ExportPath = m_strExportPath
End Property
Public Property Let ExportPath(value As String)
    m_strExportPath = value
End Property

Public Property Get ShowExportMessage() As Boolean
    ShowExportMessage = m_boolShowExportMessage
End Property
Public Property Let ShowExportMessage(value As Boolean)
    m_boolShowExportMessage = value
End Property


Private Sub Class_Initialize()
    Set m_wkbSource = ActiveWorkbook
    m_strExportPath = m_wkbSource.Path & "\" & STRING_TARGETFOLDER_EXPORTFILES
    m_boolShowExportMessage = True
End Sub


Public Sub ExportModules()
    
    If m_wkbSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
        Exit Sub
    End If
    
    Call CreateExportFolder(CleanUp:=True, strFolderName:=m_strExportPath)
    
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    For Each cmpComponent In m_wkbSource.VBProject.VBComponents
        
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                szFileName = szFileName & ".cls"
        End Select
        ''' Export the component to a text file.
        cmpComponent.Export m_strExportPath & szFileName
    
    Next cmpComponent

    If (m_boolShowExportMessage) Then MsgBox "VBA Modules are export to (" & m_strExportPath & ")"

End Sub


Private Sub CreateExportFolder(CleanUp As Boolean, strFolderName As String)

    Dim WshShell As Object
    Dim FSO As Object
    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(strFolderName) = False Then
        MkDir strFolderName
    End If
    
    If (CleanUp) Then
        On Error Resume Next
        Kill strFolderName & "\*.*"
        If (Err.Number = 53 Or Err.Number = 0) Then
            ' Do nothing
        Else
            Call Err.Raise(Err.Number, Err.Description)
        End If
    End If

End Sub


