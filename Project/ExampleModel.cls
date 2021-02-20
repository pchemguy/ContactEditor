VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TModel
    Field1 As String
    Field2 As String
    PossibleValues As Collection
End Type

Private this As TModel

Public Property Get Field1() As String
    Field1 = this.Field1
End Property

Public Property Let Field1(ByVal value As String)
    this.Field1 = value
End Property

Public Property Get Field2() As String
    Field2 = this.Field2
End Property

Public Property Let Field2(ByVal value As String)
    this.Field2 = value
End Property

Public Property Get PossibleValues() As Collection
    Set PossibleValues = this.PossibleValues
End Property

Public Property Set PossibleValues(ByVal value As Collection)
    Set this.PossibleValues = value
End Property
