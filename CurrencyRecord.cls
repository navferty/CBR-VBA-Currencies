VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurrencyRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private prCurrencyDate As Date
Private prCurrencyValue As Variant 'decimal
Private prCurrencyCode As String

Public Property Get CurrencyDate() As Date
    CurrencyDate = prCurrencyDate
End Property
Public Property Let CurrencyDate(ByVal vNewValue As Date)
    prCurrencyDate = vNewValue
End Property

Public Property Get CurrencyValue() As Variant
    CurrencyValue = prCurrencyValue
End Property
Public Property Let CurrencyValue(ByVal vNewValue As Variant)
    prCurrencyValue = CDec(vNewValue)
End Property

Public Property Get CurrencyCode() As String
    CurrencyCode = prCurrencyCode
End Property
Public Property Let CurrencyCode(ByVal vNewValue As String)
    prCurrencyCode = vNewValue
End Property
