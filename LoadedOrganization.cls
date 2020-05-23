VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadedOrganization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private prOrgName As String
Private prOrgTin As String

Public Property Get OrgName() As String
    OrgName = prOrgName
End Property
Public Property Let OrgName(ByVal vNewValue As String)
    prOrgName = vNewValue
End Property

Public Property Get OrgTin() As String
    OrgTin = prOrgTin
End Property
Public Property Let OrgTin(ByVal vNewValue As String)
    prOrgTin = vNewValue
End Property

