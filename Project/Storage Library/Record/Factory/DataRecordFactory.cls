VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRecordFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "Storage Library.Record.Factory"
'@PredeclaredId
Option Explicit
Option Compare Text

Implements IDataRecordFactory

Private Type TDataRecordFactory
    ClassName As String
End Type
Private this As TDataRecordFactory


Public Function Create(ByVal ClassName As String) As IDataRecordFactory
    Dim Instance As DataRecordFactory
    Set Instance = New DataRecordFactory
    Instance.Init ClassName
    Set Create = Instance
End Function


Friend Sub Init(ByVal ClassName As String)
    this.ClassName = ClassName
End Sub


'''' Creates an instance of IDataRecordStorage
''''
'''' Args:
''''   ClassName (string):
''''     Name of a DataRecordStorage class. Supported backends:
''''       - "Worksheet"
''''
''''   Model (DataRecordModel):
''''     DataRecordModel instance
''''
''''   ConnectionString (string):
''''     Information passed to the backend necessary to connect to peristent storage.
''''     Currently, only file-based storages have been implemented, and this argument
''''     is a file name (possibly with path) that can be resolved to an absolute path
''''     pointing to the database file. See backend constructors for resolution code.
''''
''''   TableName (string):
''''     Typically the name of the table. It may also include additional information.
''''
'@Description "Creates an instance of IDataRecordStorage"
Public Function CreateInstance(ByVal ClassName As String, _
                               ByVal Model As DataRecordModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataRecordStorage
Attribute CreateInstance.VB_Description = "Creates an instance of IDataRecordStorage"
    Select Case ClassName
        Case "Worksheet"
            Set CreateInstance = DataRecordWSheet.Create(Model, ConnectionString, TableName)
        Case Else
            Dim errorDetails As TError
            With errorDetails
                .Number = ErrNo.NotImplementedErr
                .Name = "NotImplementedErr"
                .Source = "IDataRecordFactory"
                .Description = "Unsupported backend: " & ClassName
                .Message = .Description
            End With
            RaiseError errorDetails
    End Select
End Function


Private Function IDataRecordFactory_CreateInstance(ByVal Model As DataRecordModel, _
                                                   ByVal ConnectionString As String, _
                                                   ByVal TableName As String) As IDataRecordStorage
    Set IDataRecordFactory_CreateInstance = CreateInstance(this.ClassName, Model, ConnectionString, TableName)
End Function
