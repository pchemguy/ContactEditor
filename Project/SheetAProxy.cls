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


Private Property Get ISheetAProxy_Code() As String
    ISheetAProxy_Code = RecordEditor.Range("Code").Value
End Property

Private Property Let ISheetAProxy_Code(ByVal Value As String)
    RecordEditor.Range("Code").Value = Value
End Property

Private Property Get ISheetAProxy_Value() As String
    ISheetAProxy_Value = RecordEditor.Range("Quantity").Value
End Property

Private Property Let ISheetAProxy_Value(ByVal Value As String)
    RecordEditor.Range("Quantity").Value = Value
End Property

