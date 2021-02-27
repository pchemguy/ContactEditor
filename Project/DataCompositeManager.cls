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
'@IgnoreModule ProcedureNotUsed
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

Public Property Get RecordModel() As DataRecordModel
    Set RecordModel = this.RecordModel
End Property

Public Property Get FieldNames() As Variant
    FieldNames = this.TableModel.FieldNames
End Property

Public Property Get Values() As Variant
    Values = this.TableModel.Values
End Property

Public Property Get Ids() As Variant
    Ids = this.TableStorage.GetIds
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
    this.TableStorage.SaveDataFromModel
End Sub

Public Sub SaveRecordDataToRecordStorage()
    this.RecordStorage.SaveDataFromModel
End Sub

Public Sub LoadRecordFromTable(ByVal RecordId As String)
    this.TableModel.LoadValuesToDictionary this.RecordModel.Record, RecordId
    this.RecordModel.RecordIndex = this.TableModel.RecordIndexFromId(RecordId)
    this.RecordModel.IsNotDirty
End Sub

Public Sub UpdateRecordToTable()
    this.TableModel.UpdateValuesFromDictionary this.RecordModel.Record
End Sub
