Attribute VB_Name = "Module1"
Option Explicit

Public Function GetTodayCurrency(currCode As String, volatileArg As Variant) As Variant 'decimal
    Dim currItem As CurrencyRecord
    Dim col As Collection
    Set col = GetCurrency(currCode, DateAdd("d", -1, Date), Now)
    
    If col.Count = 0 Then
        Exit Function
    End If
    
    Set currItem = col.Item(col.Count - 1)
    GetTodayCurrency = currItem.CurrencyValue
    
    Debug.Print "Currency loaded, value is " + CStr(currItem.CurrencyValue)
End Function
