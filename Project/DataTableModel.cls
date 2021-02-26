VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a data table."
'@Folder "Storage.Table.Model"
'@ModuleDescription "Represents a data table."
'@IgnoreModule ProcedureNotUsed
'@Exposed
Option Explicit


Private Type TDataTableModel
    FieldIndices As Scripting.Dictionary
    IdIndices As Scripting.Dictionary
    FieldNames As Variant
    Values As Variant
End Type
Private this As TDataTableModel

Private Sub Class_Initialize()
    Set this.FieldIndices = New Scripting.Dictionary
    this.FieldIndices.CompareMode = TextCompare
    Set this.IdIndices = New Scripting.Dictionary
    this.IdIndices.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.FieldIndices = Nothing
    Set this.IdIndices = Nothing
End Sub

Public Property Get FieldIndices() As Scripting.Dictionary
    Set FieldIndices = this.FieldIndices
End Property

Public Property Get IdIndices() As Scripting.Dictionary
    Set IdIndices = this.IdIndices
End Property

Public Property Get FieldNames() As Variant
    FieldNames = this.FieldNames
End Property

Public Property Let FieldNames(ByVal FieldNamesArg As Variant)
    this.FieldNames = FieldNamesArg
End Property

Public Property Get Values() As Variant
    Values = this.Values
End Property

Public Property Let Values(ByVal ValuesArg As Variant)
    this.Values = ValuesArg
End Property

Public Sub UpdateValues(ByVal RecordIndex As Long, ByVal FieldIndex As Long, ByVal Value As Variant)
    this.Values(RecordIndex, FieldIndex) = Value
End Sub
