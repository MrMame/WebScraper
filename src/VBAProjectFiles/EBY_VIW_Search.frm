VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EBY_VIW_Search 
   Caption         =   "Search Input"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6105
   OleObjectBlob   =   "EBY_VIW_Search.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "EBY_VIW_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAbort_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()

    Dim strSearchTerm As String
    Dim strLocation As String
    Dim CategoryValue As Integer
    Dim RadiusValue As Integer
       
    strSearchTerm = txtSearchTerm.Text
    strLocation = txtLocation.Text
    CategoryValue = ReadCategoryValue
    RadiusValue = ReadRadiusValue
    
    Call EBY_CTR_Main.SearchEbay(strSearchTerm, strLocation, CategoryValue, RadiusValue)

End Sub



Private Function ReadRadiusValue() As Integer
    If (cmbRadius.Text = "") Then
        ReadRadiusValue = 0
    Else
        Dim theRadius As New EBY_DAT_ApiRadius
        ReadRadiusValue = theRadius.GetRadiusAPIValue(cmbRadius.Text)
    End If
End Function

Private Function ReadCategoryValue() As Integer
    If (cmbCategory.Text = "") Then
        ReadCategoryValue = 0
    Else
        Dim theCategories As New EBY_DAT_ApiCategories
        ReadCategoryValue = theCategories.GetCategoryAPIValue(cmbCategory.Text)
    End If
End Function




Private Sub initControls_Radius_Combobox()
    cmbRadius.Clear
    Dim theRadius As New EBY_DAT_ApiRadius
    Dim radName As Variant
    For Each radName In theRadius.GetAllRadiusNames
        cmbRadius.AddItem (radName)
    Next

End Sub



Private Sub initControls_Categories_Combobox()
    cmbCategory.Clear
    Dim theCategories As New EBY_DAT_ApiCategories
    Dim catName As Variant
    For Each catName In theCategories.GetAllCategories
        cmbCategory.AddItem (catName)
    Next
End Sub

Private Sub UserForm_Initialize()
    initControls_Radius_Combobox
    initControls_Categories_Combobox
End Sub
