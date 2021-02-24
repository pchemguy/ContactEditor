VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StorageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Forms.Workbook"
Option Explicit

Private Type TProxy
    SheetA As ISheetAProxy
    SheetB As IDataSheet
End Type
Private this As TProxy


Public Property Get SheetA() As ISheetAProxy
    Set SheetA = this.SheetA
End Property

Public Property Get SheetB() As IDataSheet
    Set SheetB = this.SheetB
End Property

Private Sub Class_Initialize()
    Set this.SheetA = New SheetAProxy
    Set this.SheetB = New SheetBProxy
End Sub

Private Sub Class_Terminate()
    Set this.SheetA = Nothing
    Set this.SheetB = Nothing
End Sub
