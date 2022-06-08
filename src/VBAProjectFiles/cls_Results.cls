VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_colAd As New Collection

Public Sub Add(Ad As cls_Ad)
    Call m_colAd.Add(Ad)
End Sub

Public Function Items() As cls_Ad()
    


End Function


