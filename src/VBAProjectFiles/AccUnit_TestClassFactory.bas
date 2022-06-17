Attribute VB_Name = "AccUnit_TestClassFactory"
Option Compare Text
Option Explicit
Option Private Module

Private m_AccUnitTestMsgBox As AccUnit_Integration.TestMessageBox

Public Sub SetAccUnitTestMsgBox(ByRef NewRef As AccUnit_Integration.TestMessageBox)
   Set m_AccUnitTestMsgBox = NewRef
End Sub

Public Function MsgBox(ByVal Prompt As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                       Optional ByVal Title As Variant, Optional ByVal HelpFile As Variant, _
                       Optional ByVal Context As Variant) As VbMsgBoxResult

    If m_AccUnitTestMsgBox Is Nothing Then
        MsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
    Else
        MsgBox = m_AccUnitTestMsgBox.Show(Prompt, Buttons, Title, HelpFile, Context)
    End If

End Function

Public Function AccUnitTestClassFactory_EBY_UTL_HttpRequester__TEST() As Object
   Set AccUnitTestClassFactory_EBY_UTL_HttpRequester__TEST = New EBY_UTL_HttpRequester__TEST
End Function


Public Function AccUnitTestClassFactory_EBY_UTL_VbaCodeExporter__TEST() As Object
   Set AccUnitTestClassFactory_EBY_UTL_VbaCodeExporter__TEST = New EBY_UTL_VbaCodeExporter__TEST
End Function

