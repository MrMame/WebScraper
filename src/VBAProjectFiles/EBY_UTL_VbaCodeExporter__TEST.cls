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



Private Const STRING_TESTCOMPONENT_PREFIX As String = "TEST_Component"


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


Public Sub PropertyExportPath_InitValue_IsSubfolderOfWorkbookCalledVBAProjectFiles()
    ' Arrange
       Dim testExporter As New EBY_UTL_VbaCodeExporter
       Dim assertFolder As String
       assertFolder = ActiveWorkbook.Path & "\VBAProjectFiles\"
    ' Assert
        Call Assert.AreEqualStrings(testExporter.ExportPath, assertFolder)
    ' cleanUp
End Sub
Public Sub PropertyWorkbook_InitValue_IsActiveWorkbook()
    ' Arrange
       Dim testExporter As New EBY_UTL_VbaCodeExporter
       Dim assertWorkbook As Workbook
       Set assertWorkbook = ActiveWorkbook
    ' Assert
        Call Assert.AreEqual(testExporter.Workbook, assertWorkbook)
    ' cleanUp
End Sub
Public Sub PropertyShowExportMessage_InitValue_IsTrue()
    ' Arrange
       Dim testExporter As New EBY_UTL_VbaCodeExporter
       Dim assertion As Boolean
       assertion = True
    ' Assert
        Call Assert.AreEqual(testExporter.ShowExportMessage, assertion)
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
        Call RemoveAllTestComponents
        Call RemoveFolder(strTestExportPath)
End Sub


'AccUnit:Row(":C:C")
'AccUnit:Row("")
'AccUnit:Row("\\\")
'AccUnit:Row("\\NotExisting\NoFolder")
Public Sub ExportModules_ClassExportToInvalidExportPath_ThrowsException76(strTestPath As String)
    ' Arrange
        ' Create Class-Component inside this workbook VBProject for Export
        Dim comp As VBComponent
        Set comp = CreateTestComponent(vbext_ct_ClassModule)
        ' Define Test Exportpath. If already exsiting, delete folder with all
        ' files.
        Dim strTestExportPath As String
        strTestExportPath = strTestPath
  
    ' Act
        Dim testExporter As New EBY_UTL_VbaCodeExporter
        testExporter.ExportPath = strTestExportPath
        
    ' Assert
        Call Assert.Throws(76)
        testExporter.ExportModules
    ' cleanUp
        ' NOT WORKING!
        ' If assertion is to catch an exception, all following statements are not processed as soon as
        ' the exception occures. The Clean Up has to be done inside the last test inside this class.
        ' NOT WORKING ----> Call RemoveAllTestComponents
End Sub

Public Sub CLEANUP_REMOVE_ALL_TESTCOMPONENTS()
    Call RemoveAllTestComponents
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
    
    ' Find a valid name fr teh test component
    Dim cnt As Integer: cnt = 1
    Dim strCompName As String: strCompName = STRING_TESTCOMPONENT_PREFIX & cnt
    Dim rerun As Boolean: rerun = False
    Do
        rerun = False
        For Each comp In cmpComponents
            If (comp.Name = strCompName) Then
                cnt = cnt + 1
                strCompName = STRING_TESTCOMPONENT_PREFIX & cnt
                rerun = True
                Exit For
            End If
        Next
    Loop While rerun
    ' Create the Test Component
    Set comp = cmpComponents.Add(ComponentType)
    comp.Name = strCompName
    
    Set CreateTestComponent = comp
End Function

Private Sub RemoveAllTestComponents()
    Dim cmpComponents  As VBComponents
    Set cmpComponents = ActiveWorkbook.VBProject.VBComponents
    Dim comp As VBComponent
    For Each comp In cmpComponents
        If (Left(comp.Name, Len(STRING_TESTCOMPONENT_PREFIX)) = STRING_TESTCOMPONENT_PREFIX) Then
            cmpComponents.Remove comp
        End If
    Next
End Sub




