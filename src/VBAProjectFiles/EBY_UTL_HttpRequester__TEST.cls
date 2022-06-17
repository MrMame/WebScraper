VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_UTL_HttpRequester__TEST"
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

Public Sub GetHTMLDocument_EmptyAddress_ReturnsEmptyHTMLDocument()
    ' Arrange
    Const Expected As Variant = ""
    Dim theHtmlDoc As New HTMLDocument
    Dim strEmptyAddress As String: strEmptyAddress = ""
    Dim testedRequester As New EBY_UTL_HttpRequester
    ' Act
    Call testedRequester.ReadHTMLDocument(strEmptyAddress, theHtmlDoc)
    ' Assert
    Assert.That theHtmlDoc.body.innerHTML, Iz.EqualTo(Expected)
End Sub
