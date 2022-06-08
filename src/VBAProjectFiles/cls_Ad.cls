VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Ad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Location As String
Dim m_AdDate As Date
Dim m_AdName As String
Dim m_LinkAddress As String
Dim m_Price As Single
Dim m_Negotiable As Boolean

Public Property Let Location(value As String)
    m_Location = value
End Property
Public Property Get Location() As String
    Location = m_Location
End Property

Public Property Let AdDate(value As Date)
    m_AdDate = value
End Property
Public Property Get AdDate() As Date
    AdDate = m_AdDate
End Property

Public Property Let AdName(value As String)
    m_AdName = value
End Property
Public Property Get AdName() As String
    AdName = m_AdName
End Property

Public Property Let LinkAddress(value As String)
    m_LinkAddress = value
End Property
Public Property Get LinkAddress() As String
    LinkAddress = m_LinkAddress
End Property

Public Property Let Price(value As Single)
    m_Price = value
End Property
Public Property Get Price() As Single
    Price = m_Price
End Property

Public Property Let Negotiable(value As Boolean)
    m_Negotiable = value
End Property
Public Property Get Negotiable() As Boolean
    Negotiable = m_Negotiable
End Property

