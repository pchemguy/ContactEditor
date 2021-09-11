VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataCompositeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Composite class incorporating one Table and one Record model with backends. Record submodel is used to represent a row from the Table."
'@Folder "Storage Library.Manager"
'@ModuleDescription "Composite class incorporating one Table and one Record model with backends. Record submodel is used to represent a row from the Table."
'@Exposed
'@IgnoreModule ProcedureNotUsed
''''
'''' To avoid a factory with overly complicated signature, this class does not
'''' employ the factory pattern. It should be newed as a regular class, and then
'''' the user should call InitRecord and InitTable methods directly to complete
'''' initialization. These methods replace the factory/constructor pattern.
''''
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

'@Description "Initializes Record subpackage components"
Public Sub InitRecord(ByVal ClassName As String, ByVal ConnectionString As String, ByVal TableName As String)
Attribute InitRecord.VB_Description = "Initializes Record subpackage components"
    Set this.RecordStorage = DataRecordFactory.CreateInstance(ClassName, this.RecordModel, ConnectionString, TableName)
End Sub

'@Description "Initializes Table subpackage components"
Public Sub InitTable(ByVal ClassName As String, ByVal ConnectionString As String, ByVal TableName As String)
Attribute InitTable.VB_Description = "Initializes Table subpackage components"
    Set this.TableStorage = DataTableFactory.CreateInstance(ClassName, this.TableModel, ConnectionString, TableName)
End Sub

Public Property Get Record() As Scripting.Dictionary
    Set Record = this.RecordModel.Record
End Property

Public Property Get RecordModel() As DataRecordModel
    Set RecordModel = this.RecordModel
End Property

Public Property Get TableModel() As DataTableModel
    Set TableModel = this.TableModel
End Property

Public Property Get FieldNames() As Variant
    FieldNames = this.TableModel.FieldNames
End Property

Public Property Get Values() As Variant
    Values = this.TableModel.Values
End Property

Public Property Get IDs() As Variant
    IDs = this.TableStorage.GetIds
End Property

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
    this.TableModel.CopyRecordToDictionary this.RecordModel.Record, RecordId
    this.RecordModel.RecordIndex = this.TableModel.RecordIndexFromId(RecordId)
    this.RecordModel.IsNotDirty
End Sub

Public Sub UpdateRecordToTable()
    this.TableModel.UpdateRecordFromDictionary this.RecordModel.Record
End Sub
