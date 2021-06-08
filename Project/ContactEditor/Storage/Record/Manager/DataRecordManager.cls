VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRecordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ContactEditor.Storage.Record.Manager"
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDataRecordManager

Private Type TDataRecordManager
    Model As DataRecordModel
    Storage As IDataRecordStorage
End Type
Private this As TDataRecordManager


Public Function Create(ByVal ClassName As String, ByVal ConnectionString As String, ByVal TableName As String) As IDataRecordManager
    Dim Result As DataRecordManager
    Set Result = New DataRecordManager
    Result.Init ClassName, ConnectionString, TableName
    Set Create = Result
End Function

Public Sub Init(ByVal ClassName As String, ByVal ConnectionString As String, ByVal TableName As String)
    Set this.Model = New DataRecordModel
    Set this.Storage = DataRecordFactory.CreateInstance(ClassName, this.Model, ConnectionString, TableName)
End Sub

Private Sub Class_Terminate()
    Set this.Model = Nothing
    Set this.Storage = Nothing
End Sub

Private Property Get IDataRecordManager_Model() As DataRecordModel
    Set IDataRecordManager_Model = this.Model
End Property

Private Sub IDataRecordManager_LoadDataIntoModel()
    this.Storage.LoadDataIntoModel
End Sub

Private Sub IDataRecordManager_SaveDataFromModel()
    this.Storage.SaveDataFromModel
End Sub
