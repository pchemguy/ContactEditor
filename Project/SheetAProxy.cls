VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetAProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Forms.Workbook"
Option Explicit

Implements ISheetAProxy


Private Property Get ISheetAProxy_FieldA() As String
    ISheetAProxy_FieldA = SheetA.Range("FieldAValue").value
End Property

Private Property Let ISheetAProxy_FieldA(ByVal value As String)
    SheetA.Range("FieldAValue").value = value
End Property

Private Property Get ISheetAProxy_FieldB() As String
    ISheetAProxy_FieldB = SheetA.Range("FieldBValue").value
End Property

Private Property Let ISheetAProxy_FieldB(ByVal value As String)
    SheetA.Range("FieldBValue").value = value
End Property

Private Sub ISheetAProxy_Clear()
    SheetA.Range("FieldAValue").ClearContents
    SheetA.Range("FieldBValue").ClearContents
End Sub
