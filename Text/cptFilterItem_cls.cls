VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cptFilterItem_cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<cpt_version>v0.2.0</cpt_version>
Option Explicit

Private m_Value As String
Private m_Method As String
Private m_Count As Integer

Public Property Get Value() As String
     Value = m_Value
End Property

Public Property Let Value(ByVal newValue As String)
     m_Value = newValue
End Property

Public Property Get Method() As String
     Method = m_Method
End Property

Public Property Let Method(ByVal newMethod As String)
    m_Method = newMethod
End Property

Public Property Get Count() As Integer
    Count = m_Count
End Property

Public Property Let Count(ByVal newCount As Integer)
    m_Count = newCount
End Property

Private Sub Class_Initialize()
    m_Method = "Equals"
    m_Count = 0
End Sub
