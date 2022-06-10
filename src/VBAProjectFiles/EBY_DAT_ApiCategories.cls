VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_DAT_ApiCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const STRING_CATEGORY_KEY_ALL As String = "all"
Private Const INTEGER_CATEGORY_VALUE_ALL As Integer = 0

Private Const STRING_CATEGORY_KEY_TOGIVEAWAY As String = "to_give_away"
Private Const INTEGER_CATEGORY_VALUE_TOGIVEAWAY As Integer = 192

Private m_dicCategories As New Dictionary


Public Property Get APIValue_All() As Integer
    APIValue_All = GetCategoryAPIValue(STRING_CATEGORY_KEY_ALL)
End Property
Public Property Get APIValue_ToGiveAway() As Integer
    APIValue_ToGiveAway = GetCategoryAPIValue(STRING_CATEGORY_KEY_TOGIVEAWAY)
End Property


Private Sub Class_Initialize()
    Call m_dicCategories.Add(STRING_CATEGORY_KEY_ALL, INTEGER_CATEGORY_VALUE_ALL)
    Call m_dicCategories.Add(STRING_CATEGORY_KEY_TOGIVEAWAY, INTEGER_CATEGORY_VALUE_TOGIVEAWAY)
End Sub


Public Function GetAllCategories() As Variant()
    GetAllCategories = m_dicCategories.Keys
End Function

Public Function GetCategoryAPIValue(CategoryName As String) As Integer
    If (Not m_dicCategories.Exists(CategoryName)) Then
        Call Err.Raise(666, "Unknown Category Name", "The Name of the category is unknown!")
    Else
        GetCategoryAPIValue = m_dicCategories.Item(CategoryName)
    End If
End Function


