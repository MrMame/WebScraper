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


Public Sub ExportModules_StandardExportPath_IsSubfolderOfWorkbookCalledVBAProjectFiles()
    ' Arrange
       Dim testExporter As New EBY_UTL_VbaCodeExporter
       Dim assertFolder As String
       assertFolder = ActiveWorkbook.Path & "\VBAProjectFiles\"
    ' Assert
        Call Assert.AreEqualStrings(testExporter.ExportPath, assertFolder)
    ' cleanUp
End Sub


' AccUnit:ClickingMsgBox(vbOk)
Public Sub ExportModules_ClassExportToWriteableCustomExportPath_IsWorking()
    ' Arrange
        ' Create Class-Component inside this workbook VBProject for Export
        Dim comp As VBComponent
        Set comp = CreateTestComponent(vbext_ct_ClassModule)
        ' Define Test Exportpath. If already exsiting, delete folder with all
        ' files.
        Dim WshShell As Object: Set WshShell = CreateObject("WScript.Shell")
        Dim strTestExportPath As String
        strTestExportPath = WshShell.SpecialFolders("MyDocuments") & "\TestPath\"
        Call RemoveFolder(strTestExportPath)
        ' Create Filename for component file
        Dim strTestExportFilename As String
        strTestExportFilename = comp.Name & ".cls"
    ' Act
        Dim testExporter As New EBY_UTL_VbaCodeExporter
        testExporter.ExportPath = strTestExportPath
        testExporter.ExportModules
    ' Assert
        Call Assert.AreEqualStrings(strTestExportFilename, _
                                    Dir(strTestExportPath & strTestExportFilename))
    ' cleanUp
        Call RemoveTestComponent(comp)
        Call RemoveFolder(strTestExportPath)
End Sub





'--------------------------------------------------------------------
' Privates
'--------------------------------------------------------------------

Private Sub RemoveFolder(strPath As String)
    On Error Resume Next
    Kill strPath & "*.*"
    RmDir strPath
    On Error GoTo 0
    If (Dir(strPath) <> "") Then Call Err.Raise(666, "EBY_UTL_VbaCodeExporter___TEST", "Was not able to remove Folder!")
End Sub


Private Function CreateTestComponent(ComponentType As vbext_ComponentType) As VBComponent
    Dim cmpComponents  As VBComponents
    Set cmpComponents = ActiveWorkbook.VBProject.VBComponents
    Dim comp As VBComponent
    Set comp = cmpComponents.Add(ComponentType)
    Set CreateTestComponent = comp
End Function
Private Sub RemoveTestComponent(RemoveComponent As VBComponent)
    Dim cmpComponents  As VBComponents
    Set cmpComponents = ActiveWorkbook.VBProject.VBComponents
    Call cmpComponents.Remove(RemoveComponent) ' Remove Dummy VBProject Component
End Sub




