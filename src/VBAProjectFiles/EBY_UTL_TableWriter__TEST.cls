VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_UTL_TableWriter__TEST"
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

Public Sub WriteAds_IfTargetSheetIsNull_ThrowsException91()
    ' Arrange
    Dim tabWriter As New EBY_UTL_TableWriter
    Dim colAds As New Collection
    Dim shtTarget As Worksheet: Set shtTarget = Nothing
    ' Act
    Call Assert.Throws(91)
    Call tabWriter.WriteAds(colAds, shtTarget)
End Sub

Public Sub WriteAds_IfAdsCollectionIsNull_ThrowsException91()
    ' Arrange
    Dim tabWriter As New EBY_UTL_TableWriter
    Dim colAds As Collection: Set colAds = Nothing
    Dim shtTarget As Worksheet: Set shtTarget = ActiveSheet
    ' Act
    Call Assert.Throws(91)
    Call tabWriter.WriteAds(colAds, shtTarget)
End Sub
Public Sub WriteAds_IfTargetSheetIsWriteProtected_ThrowsErrorNumber()
    ' Arrange
    Err.Raise vbObjectError, "WriteAds_TargetSheetIsWriteProtected_ThrowsErrorNumber", "Test not implemented"
    Const Expected As Variant = "expected value"
    Dim Actual As Variant
    ' Act
    Actual = "actual value"
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
Public Sub WriteAds_AdsWrittenHeadernames_AreAsExpected()
    ' Arrange
    Err.Raise vbObjectError, "WriteAds_AdsWrittenHeadernames_AreAsExpected", "Test not implemented"
    Const Expected As Variant = "expected value"
    Dim Actual As Variant
    ' Act
    Actual = "actual value"
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
Public Sub WriteAds_IfTargetsheetHasOldData_OldValuesAreRemoved()
    ' Arrange
    Err.Raise vbObjectError, "WriteAds_IfTargetsheetHasOldData_OldValuesAreRemoved", "Test not implemented"
    Const Expected As Variant = "expected value"
    Dim Actual As Variant
    ' Act
    Actual = "actual value"
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
Public Sub WriteAds_IfAdsCollectionIsEmpty_WrittenTableHasOnlyHeadernames()
    ' Arrange
    Err.Raise vbObjectError, "WriteAds_IfAdsCollectionIsEmpty_WrittenTableHasOnlyHeadernames", "Test not implemented"
    Const Expected As Variant = "expected value"
    Dim Actual As Variant
    ' Act
    Actual = "actual value"
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
