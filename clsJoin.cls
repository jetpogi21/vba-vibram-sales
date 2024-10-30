VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private vSource As String
Private vAlias As String
Private vJoinType As String
Private vRightFields As New clsArray
Private vLeftFields As New clsArray

Private Sub Class_Initialize()
    vJoinType = "INNER"
End Sub

Property Let Source(value As String)
    vSource = value
End Property

Property Let Alias(value As String)
    vAlias = value
End Property

Property Let JoinType(value As String)
    vJoinType = value
End Property

Property Let RightFields(value As Variant)
    vRightFields.arr = value
End Property

Property Let LeftFields(value As Variant)
    vRightFields.arr = value
    vLeftFields.arr = value
End Property

Property Get Source() As String
    Source = vSource
End Property

Property Get Alias() As String
    Alias = vAlias
End Property

Property Get JoinType() As String
    JoinType = vJoinType
End Property

Property Get RightFields() As clsArray
    Set RightFields = vRightFields
End Property

Property Get LeftFields() As clsArray
    Set LeftFields = vLeftFields
End Property

