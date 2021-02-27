VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRecordModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a set of property-value pairs."
'@Folder "Storage.Record.Model"
'@ModuleDescription "Represents a set of property-value pairs."
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess
'@Exposed
Option Explicit

Private Type TDataRecordModel
    Record As Scripting.Dictionary
    RecordIndex As Long
    IsDirty As Boolean
End Type
Private this As TDataRecordModel


Private Sub Class_Initialize()
    Set this.Record = New Scripting.Dictionary
    this.Record.CompareMode = TextCompare
    this.RecordIndex = -1
    this.IsDirty = False
End Sub

Private Sub Class_Terminate()
    Set this.Record = Nothing
End Sub

Public Property Get Record() As Scripting.Dictionary
    Set Record = this.Record
End Property

Public Property Get RecordIndex() As Long
    RecordIndex = this.RecordIndex
End Property

Public Property Let RecordIndex(ByVal RecordIndexArg As Long)
    this.RecordIndex = RecordIndexArg
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = this.IsDirty
End Property

Public Function GetField(ByVal FieldName As String) As Variant
    GetField = this.Record(FieldName)
End Function

Public Sub SetField(ByVal FieldName As String, ByVal Value As Variant)
    this.Record(FieldName) = Value
    this.IsDirty = True
End Sub

Public Sub IsNotDirty()
    this.IsDirty = False
End Sub
