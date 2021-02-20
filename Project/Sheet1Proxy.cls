VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1Proxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISheet1Proxy

Private Property Get ISheet1Proxy_Field1() As String
    ISheet1Proxy_Field1 = Sheet1.Range("Field1Value").value
End Property

Private Property Let ISheet1Proxy_Field1(ByVal value As String)
    Sheet1.Range("Field1Value").value = value
End Property

Private Property Get ISheet1Proxy_Field2() As String
    ISheet1Proxy_Field2 = Sheet1.Range("Field2Value").value
End Property

Private Property Let ISheet1Proxy_Field2(ByVal value As String)
    Sheet1.Range("Field2Value").value = value
End Property

Private Sub ISheet1Proxy_Clear()
    Sheet1.Range("Field1Value").ClearContents
    Sheet1.Range("Field2Value").ClearContents
End Sub
