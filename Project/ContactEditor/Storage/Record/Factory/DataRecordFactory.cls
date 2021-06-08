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
'@Folder "ContactEditor.Storage.Record.Factory"
'@PredeclaredId
Option Explicit
Option Compare Text

Implements IDataRecordFactory

Private Type TDataRecordFactory
    ClassName As String
End Type
Private this As TDataRecordFactory


Public Function Create(ByVal ClassName As String) As IDataRecordFactory
    Dim Result As DataRecordFactory
    Set Result = New DataRecordFactory
    Result.Init ClassName
    Set Create = Result
End Function


Public Sub Init(ByVal ClassName As String)
    this.ClassName = ClassName
End Sub


Public Function CreateInstance(ByVal ClassName As String, _
                               ByVal Model As DataRecordModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataRecordStorage
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

