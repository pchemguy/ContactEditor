VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataCompositeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Composite class incorporating one Table and one Record model with backends. Record submodel is used to represent a row from the Table."
'@Folder "Storage.Composite.Manager"
'@ModuleDescription "Composite class incorporating one Table and one Record model with backends. Record submodel is used to represent a row from the Table."
'@PredeclaredId
'@Exposed
Option Explicit

Private Type TDataCompositeManager
    RecordModel As DataRecordModel
    RecordStorage As IDataRecordStorage
    TableModel As DataTableModel
    TableStorage As IDataTableStorage
End Type
Private this As TDataCompositeManager


Private Sub Class_Initialize()
    Set this.RecordModel = New DataRecordModel
    Set this.TableModel = New DataTableModel
End Sub

Private Sub Class_Terminate()
    Set this.RecordModel = Nothing
    Set this.TableModel = Nothing
End Sub

Public Property Get Record() As Scripting.Dictionary
    Set Record = this.RecordModel.Record
End Property

Public Property Get FieldNames() As Variant
    FieldNames = this.TableModel.FieldNames
End Property

Public Property Get Values() As Variant
    Values = this.TableModel.Values
End Property

Public Sub InitRecord(ByVal ClassName As String, ByVal ConnectionString As String, ByVal TableName As String)
    Set this.RecordStorage = DataRecordFactory.CreateInstance(ClassName, this.RecordModel, ConnectionString, TableName)
End Sub

Public Sub InitTable(ByVal ClassName As String, ByVal ConnectionString As String, ByVal TableName As String)
    Set this.TableStorage = DataTableFactory.CreateInstance(ClassName, this.TableModel, ConnectionString, TableName)
End Sub

Public Sub LoadDataIntoModel()
    this.TableStorage.LoadDataIntoModel
    this.RecordStorage.LoadDataIntoModel
End Sub

Public Sub SaveDataFromModel()
    this.RecordStorage.SaveDataFromModel
    'this.TableStorage.SaveDataFromModel 'Not yet implemented
End Sub

Public Sub LoadRecordFromTable(ByVal RecordId As String)
    Dim RecordIndex As Long: RecordIndex = this.TableModel.IdIndices(RecordId)
    Dim Record As Variant: Record = Application.WorksheetFunction.Index(Values, RecordIndex)
    Dim FieldName As Variant
    For Each FieldName In this.TableModel.FieldNames
        this.RecordModel.Record.Item(CStr(FieldName)) = Record(this.TableModel.FieldIndices(CStr(FieldName)))
    Next FieldName
End Sub

Public Sub UpdateRecordToTable()
    Dim IdName As String: IdName = this.TableModel.FieldNames(1)
    Dim RecordId As String: RecordId = CStr(this.RecordModel.Record.Item(CStr(IdName)))
    Dim RecordIndex As Long: RecordIndex = this.TableModel.IdIndices(RecordId)
    Dim FieldNameIndex As Long
    Dim FieldName As Variant
    For Each FieldName In this.TableModel.FieldNames
        this.TableModel.UpdateValues RecordIndex, this.TableModel.FieldIndices(CStr(FieldName)), this.RecordModel.Record.Item(CStr(FieldName))
    Next FieldName
End Sub

