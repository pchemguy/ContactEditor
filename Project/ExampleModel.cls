VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Forms.Example"
Option Explicit

Private Type TModel
    FieldA As String
    FieldB As String
    PossibleValues As Collection
End Type

Private this As TModel

Public Property Get FieldA() As String
    FieldA = this.FieldA
End Property

Public Property Let FieldA(ByVal value As String)
    this.FieldA = value
End Property

Public Property Get FieldB() As String
    FieldB = this.FieldB
End Property

Public Property Let FieldB(ByVal value As String)
    this.FieldB = value
End Property

Public Property Get PossibleValues() As Collection
    Set PossibleValues = this.PossibleValues
End Property

Public Property Set PossibleValues(ByVal value As Collection)
    Set this.PossibleValues = value
End Property
