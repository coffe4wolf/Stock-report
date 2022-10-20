Attribute VB_Name = "m_test"
Option Explicit

Sub test()
    
    ' Call createPurchasingStrategy
    
    
    Application.EnableEvents = False
  
End Sub

Sub test2()

    Dim str As String
    Dim a As Variant
    
    
    str = "aa,b,d,e"
    a = Split(str, ",")
    
    MsgBox (a(0))
'
'    For Each a In Split(str, ",")
'
'        Debug.Print a
'
'    Next a

End Sub


