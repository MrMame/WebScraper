VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_UTL_VbaCodeExporter__TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------






Public Sub ExportModules_InitExportPath_IsWorkbookVBAProjectFilesSubFolder()
    ' Arrange
    Dim testCodeExporter As New EBY_UTL_VbaCodeExporter
    ' Act
    ' Assert
    Assert.That testCodeExporter.ExportPath, Iz.EqualTo(ActiveWorkbook.Path & "\VBAProjectFiles\")
    ' CleanUp
End Sub

' AccUnit:ClickingMsgBox(vbYes, vbOk, vbNo)
Public Sub ExportModules_ClassComponentExport_ToStandardExportPath()
    ' Arrange
        ' Create Class-Component inside this workbook VBProject for Export
        Dim cmpComponents  As VBComponents
        Set cmpComponents = ActiveWorkbook.VBProject.VBComponents
        Dim comp As VBComponent
        Set comp = cmpComponents.Add(vbext_ct_ClassModule)
        ' If there is already a export of this class existing, delete it first
        Dim testExporter As New EBY_UTL_VbaCodeExporter
        Dim strExportFilename As String
        strExportFilename = comp.Name & ".cls"
        Dim strExportFileFullname As String
        strExportFileFullname = testExporter.ExportPath & strExportFilename
        If (Dir(strExportFileFullname) = strExportFilename) Then Kill strExportFileFullname
    ' Act
        testExporter.ExportModules
    ' Assert
        Call Assert.AreEqualStrings(strExportFilename, Dir(strExportFileFullname))
    ' cleanUp
        Call cmpComponents.Remove(comp) ' Remove Dummy VBProject Component
        Kill strExportFileFullname      ' Remove Exported Test File
End Sub


' AccUnit:ClickingMsgBox(vbOk)
Public Sub ExportModules_ClassComponentExport_ToWriteableCustomExportPath()
    ' Arrange
        ' Create Class-Component inside this workbook VBProject for Export
        Dim cmpComponents  As VBComponents
        Set cmpComponents = ActiveWorkbook.VBProject.VBComponents
        Dim comp As VBComponent
        Set comp = cmpComponents.Add(vbext_ct_ClassModule)
        ' If there is already a export of this class existing, delete it first
        Dim testExporter As New EBY_UTL_VbaCodeExporter
        
        Dim WshShell As Object
        Set WshShell = CreateObject("WScript.Shell")
        Dim strSpecialPath As String
        strSpecialPath = WshShell.SpecialFolders("MyDocuments")
        Dim strTestExportPath As String
        strTestExportPath = strSpecialPath & "\"
        
        Dim strExportFilename As String
        strExportFilename = comp.Name & ".cls"
        Dim strExportFileFullname As String
        strExportFileFullname = testExporter.ExportPath & strExportFilename
        If (Dir(strExportFileFullname) = strExportFilename) Then Kill strExportFileFullname
    ' Act
        testExporter.ExportModules
    ' Assert
        Call Assert.AreEqualStrings(strExportFilename, Dir(strExportFileFullname))
    ' cleanUp
        Call cmpComponents.Remove(comp) ' Remove Dummy VBProject Component
        Kill strExportFileFullname      ' Remove Exported Test File
End Sub





Private Sub RemoveReadOnly(filePath As String)
    Dim FSO As FileSystemObject
    Dim f As Scripting.Folder
    Set FSO = New FileSystemObject
    Set f = FSO.GetFolder(filePath)

    If f.Attributes And ReadOnly Then 'It's read-only. Remove that attribute.
        f.Attributes = f.Attributes - ReadOnly
    End If
End Sub
Private Sub SetReadOnly(filePath As String)
    Dim FSO As FileSystemObject
    Dim f As Scripting.Folder
    Set FSO = New FileSystemObject
    Set f = FSO.GetFolder(filePath)

    If Not f.Attributes And ReadOnly Then 'It's read-only. Remove that attribute.
        f.Attributes = f.Attributes Or ReadOnly
    End If
End Sub

