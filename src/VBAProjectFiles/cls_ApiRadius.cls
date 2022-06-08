VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ApiRadius"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STRING_API_RADIUS_5KM_KEY As String = "KM_5"
Private Const STRING_API_RADIUS_10KM_KEY As String = "KM_10"
Private Const STRING_API_RADIUS_20KM_KEY As String = "KM_20"
Private Const STRING_API_RADIUS_30KM_KEY As String = "KM_30"
Private Const STRING_API_RADIUS_50KM_KEY As String = "KM_50"
Private Const STRING_API_RADIUS_100KM_KEY As String = "KM_100"
Private Const STRING_API_RADIUS_150KM_KEY As String = "KM_150"
Private Const STRING_API_RADIUS_200KM_KEY As String = "KM_200"

Private Const INTEGER_API_RADIUS_5KM_VALUE As Integer = 5
Private Const INTEGER_API_RADIUS_10KM_VALUE As Integer = 10
Private Const INTEGER_API_RADIUS_20KM_VALUE As Integer = 20
Private Const INTEGER_API_RADIUS_30KM_VALUE As Integer = 30
Private Const INTEGER_API_RADIUS_50KM_VALUE As Integer = 50
Private Const INTEGER_API_RADIUS_100KM_VALUE As Integer = 100
Private Const INTEGER_API_RADIUS_150KM_VALUE As Integer = 150
Private Const INTEGER_API_RADIUS_200KM_VALUE As Integer = 200

Dim m_dicRadius As New Dictionary

Private Sub Class_Initialize()
    Call m_dicRadius.Add(STRING_API_RADIUS_5KM_KEY, INTEGER_API_RADIUS_5KM_VALUE)
    Call m_dicRadius.Add(STRING_API_RADIUS_10KM_KEY, INTEGER_API_RADIUS_10KM_VALUE)
    Call m_dicRadius.Add(STRING_API_RADIUS_20KM_KEY, INTEGER_API_RADIUS_20KM_VALUE)
    Call m_dicRadius.Add(STRING_API_RADIUS_30KM_KEY, INTEGER_API_RADIUS_30KM_VALUE)
    Call m_dicRadius.Add(STRING_API_RADIUS_50KM_KEY, INTEGER_API_RADIUS_50KM_VALUE)
    Call m_dicRadius.Add(STRING_API_RADIUS_100KM_KEY, INTEGER_API_RADIUS_100KM_VALUE)
    Call m_dicRadius.Add(STRING_API_RADIUS_150KM_KEY, INTEGER_API_RADIUS_150KM_VALUE)
    Call m_dicRadius.Add(STRING_API_RADIUS_200KM_KEY, INTEGER_API_RADIUS_200KM_VALUE)
End Sub

Public Property Get APIValue_5KM() As Integer
    APIValue_5KM = GetRadiusAPIValue(STRING_API_RADIUS_5KM_KEY)
End Property
Public Property Get APIValue_10KM() As Integer
    APIValue_10KM = GetRadiusAPIValue(STRING_API_RADIUS_10KM_KEY)
End Property
Public Property Get APIValue_20KM() As Integer
    APIValue_20KM = GetRadiusAPIValue(STRING_API_RADIUS_20KM_KEY)
End Property
Public Property Get APIValue_30KM() As Integer
    APIValue_30KM = GetRadiusAPIValue(STRING_API_RADIUS_30KM_KEY)
End Property
Public Property Get APIValue_50KM() As Integer
    APIValue_50KM = GetRadiusAPIValue(STRING_API_RADIUS_50KM_KEY)
End Property
Public Property Get APIValue_100KM() As Integer
    APIValue_100KM = GetRadiusAPIValue(STRING_API_RADIUS_100KM_KEY)
End Property
Public Property Get APIValue_150KM() As Integer
    APIValue_150KM = GetRadiusAPIValue(STRING_API_RADIUS_150KM_KEY)
End Property
Public Property Get APIValue_200KM() As Integer
    APIValue_200KM = GetRadiusAPIValue(STRING_API_RADIUS_200KM_KEY)
End Property


Public Function GetAllRadiusNames() As Variant()
    GetAllRadiusNames = m_dicRadius.Keys
End Function

Public Function GetRadiusAPIValue(RadiusName As String) As Integer
    If (Not m_dicRadius.Exists(RadiusName)) Then
        Call Err.Raise(666, "Unknown Radius Name", "The Name of the radius is unknown!")
    Else
        GetRadiusAPIValue = m_dicRadius.Item(RadiusName)
    End If
End Function


