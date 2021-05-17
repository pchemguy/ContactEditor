VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "ContactEditor.Storage.Table.Factory"
'@PredeclaredId
Option Explicit
Option Compare Text

Implements IDataTableFactory

Private Type TDataTableFactory
    ClassName As String
End Type
Private this As TDataTableFactory


'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As DataTableFactory
Attribute Class.VB_Description = "Returns class reference"
    Set Class = DataTableFactory
End Property


Public Function Create(ByVal ClassName As String) As IDataTableFactory
    Guard.NonDefaultInstance Me
    
    Dim result As DataTableFactory
    Set result = New DataTableFactory
    result.Init ClassName
    Set Create = result
End Function


Public Sub Init(ByVal ClassName As String)
    this.ClassName = ClassName
End Sub


Public Function CreateInstance(ByVal ClassName As String, _
                               ByVal Model As DataTableModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataTableStorage
    Select Case ClassName
        Case "ADODB"
            Set CreateInstance = DataTableADODB.Create(Model, ConnectionString, TableName)
        Case "Worksheet"
            Set CreateInstance = DataTableWSheet.Create(Model, ConnectionString, TableName)
        Case "CSV"
            Set CreateInstance = DataTableCSV.Create(Model, ConnectionString, TableName)
        Case Else
            Dim errorDetails As TError
            With errorDetails
                .Number = ErrNo.NotImplementedErr
                .Name = "NotImplementedErr"
                .Source = "IDataTableFactory"
                .Description = "Unsupported backend: " & ClassName
                .Message = .Description
            End With
            RaiseError errorDetails
    End Select
End Function


Private Function IDataTableFactory_CreateInstance(ByVal Model As DataTableModel, _
                                                  ByVal ConnectionString As String, _
                                                  ByVal TableName As String) As IDataTableStorage
    Set IDataTableFactory_CreateInstance = CreateInstance(this.ClassName, Model, ConnectionString, TableName)
End Function
