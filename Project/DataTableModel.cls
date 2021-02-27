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
'@IgnoreModule IndexedDefaultMemberAccess
'@Exposed
Option Explicit


Private Type TDataTableModel
    FieldIndices As Scripting.Dictionary
    IdIndices As Scripting.Dictionary
    DirtyRecords As Scripting.Dictionary
    FieldNames As Variant
    Values As Variant
End Type
Private this As TDataTableModel

Private Sub Class_Initialize()
    Set this.FieldIndices = New Scripting.Dictionary
    this.FieldIndices.CompareMode = TextCompare
    Set this.IdIndices = New Scripting.Dictionary
    this.IdIndices.CompareMode = TextCompare
    Set this.DirtyRecords = New Scripting.Dictionary
    this.DirtyRecords.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.FieldIndices = Nothing
    Set this.IdIndices = Nothing
    Set this.DirtyRecords = Nothing
End Sub

Public Property Get FieldIndices() As Scripting.Dictionary
    Set FieldIndices = this.FieldIndices
End Property

Public Property Get IdIndices() As Scripting.Dictionary
    Set IdIndices = this.IdIndices
End Property

Public Property Get DirtyRecords() As Scripting.Dictionary
    Set DirtyRecords = this.DirtyRecords
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

Public Property Get IsDirty() As Boolean
    IsDirty = this.DirtyRecords.Count > 0
End Property

Public Function RecordIndexFromId(ByVal RecordId As String) As Long
    RecordIndexFromId = this.IdIndices(RecordId)
End Function

Public Function RecordValuesFromId(ByVal RecordId As String) As Variant
    Dim RecordIndex As Long: RecordIndex = this.IdIndices(RecordId)
    RecordValuesFromId = Application.WorksheetFunction.Index(Values, RecordIndex)
End Function

Public Function FieldIndexFromName(ByVal FieldName As String) As Long
    FieldIndexFromName = this.FieldIndices(FieldName)
End Function

Public Sub IsNotDirty()
    this.DirtyRecords.RemoveAll
End Sub

Public Sub UpdateValuesFromDictionary(ByVal Record As Scripting.Dictionary)
    Const ID_NAME_INDEX As Long = 1
    Dim FieldIdName As String: FieldIdName = this.FieldNames(ID_NAME_INDEX)
    Dim RecordId As String: RecordId = CStr(Record(FieldIdName))
    Dim RecordIndex As Long: RecordIndex = RecordIndexFromId(RecordId)
    this.DirtyRecords(RecordId) = RecordIndex
    
    Dim FieldName As Variant
    Dim FieldIndex As Long
    For Each FieldName In this.FieldNames
        FieldIndex = FieldIndexFromName(FieldName)
        this.Values(RecordIndex, FieldIndex) = Record(CStr(FieldName))
    Next FieldName
End Sub

Public Sub LoadValuesToDictionary(ByVal Record As Scripting.Dictionary, ByVal RecordId As String)
    Dim RecordIndex As Long: RecordIndex = RecordIndexFromId(RecordId)
    Dim FieldName As Variant
    Dim FieldIndex As Long
    For Each FieldName In this.FieldNames
        FieldIndex = FieldIndexFromName(FieldName)
        Record(CStr(FieldName)) = this.Values(RecordIndex, FieldIndex)
    Next FieldName
End Sub
