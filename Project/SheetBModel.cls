VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetBModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Forms.Workbook"
Option Explicit

Private Type TModel
    Code As String
    Name As String
End Type

Private this As TModel

Public Property Get Code() As String
    Code = this.Code
End Property

Public Property Let Code(ByVal Value As String)
    this.Code = Value
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal Value As String)
    this.Name = Value
End Property
