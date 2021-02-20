VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TProxy
    Sheet1 As ISheet1Proxy
    Sheet2 As IDataSheet
End Type

Private this As TProxy

Public Property Get Sheet1() As ISheet1Proxy
    Set Sheet1 = this.Sheet1
End Property

Public Property Get Sheet2() As IDataSheet
    Set Sheet2 = this.Sheet2
End Property

Private Sub Class_Initialize()
    Set this.Sheet1 = New Sheet1Proxy
    Set this.Sheet2 = New Sheet2Proxy
End Sub

Private Sub Class_Terminate()
    Set this.Sheet1 = Nothing
    Set this.Sheet2 = Nothing
End Sub
