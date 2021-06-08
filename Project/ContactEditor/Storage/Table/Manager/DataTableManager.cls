VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ContactEditor.Storage.Table.Manager"
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDataTableManager

Private Type TDataTableManager
    Model As DataTableModel
    Storage As IDataTableStorage
End Type
Private this As TDataTableManager


Public Function Create(ByVal ClassName As String, ByVal ConnectionString As String, ByVal TableName As String) As IDataTableManager
    Dim Result As DataTableManager
    Set Result = New DataTableManager
    Result.Init ClassName, ConnectionString, TableName
    Set Create = Result
End Function

Public Sub Init(ByVal ClassName As String, ByVal ConnectionString As String, ByVal TableName As String)
    Set this.Model = New DataTableModel
    Set this.Storage = DataTableFactory.CreateInstance(ClassName, this.Model, ConnectionString, TableName)
End Sub

Private Sub Class_Terminate()
    Set this.Model = Nothing
    Set this.Storage = Nothing
End Sub

Private Property Get IDataTableManager_Model() As DataTableModel
    Set IDataTableManager_Model = this.Model
End Property

Private Sub IDataTableManager_LoadDataIntoModel()
    this.Storage.LoadDataIntoModel
End Sub

Private Sub IDataTableManager_SaveDataFromModel()
    this.Storage.SaveDataFromModel
End Sub
