VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Forms.Record Editor"
Option Explicit

Private Type TModel
    Code As String
    Quantity As String
    PossibleValues As Collection
End Type

Private this As TModel

Public Property Get Code() As String
    Code = this.Code
End Property

Public Property Let Code(ByVal Value As String)
    this.Code = Value
End Property

Public Property Get Quantity() As String
    Quantity = this.Quantity
End Property

Public Property Let Quantity(ByVal Value As String)
    this.Quantity = Value
End Property

Public Property Get PossibleValues() As Collection
    Set PossibleValues = this.PossibleValues
End Property

Public Property Set PossibleValues(ByVal Value As Collection)
    Set this.PossibleValues = Value
End Property
