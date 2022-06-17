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

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub GetHTMLDocument_EmptyAddress_ReturnsEmptyHTMLDocument()
    ' Arrange
    Const Expected As Variant = ""
    Dim theHtmlDoc As Variant
    Dim strEmptyAddress As String: strEmptyAddress = ""
    Dim testedRequester As New EBY_UTL_HttpRequester
    ' Act
    Set theHtmlDoc = testedRequester.GetHTMLDocument(strEmptyAddress)
    ' Assert
    Assert.That theHtmlDoc.body.innerHTML, Iz.EqualTo(Expected)
End Sub

