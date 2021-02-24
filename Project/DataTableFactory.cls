VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Storage.Table.Factory"
'@PredeclaredId
Option Explicit
Option Compare Text

Implements IDataTableFactory

Private Type TDataTableFactory
    ClassName As String
End Type
Private this As TDataTableFactory


Public Function Create(ByVal ClassName As String) As IDataTableFactory
    Dim result As DataTableFactory
    Set result = New DataTableFactory
    result.Init ClassName
    Set Create = result
End Function


Public Sub Init(ByVal ClassName As String)
    this.ClassName = ClassName
End Sub


Public Function CreateInstance(ByVal ClassName As String, ByVal model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String) As IDataTable
    Select Case ClassName
        Case "SheetTable"
            Set CreateInstance = SheetTable.Create(model, ConnectionString, TableName)
        Case Else
            VBA.Err.Raise vbObjectError, "IDataTableFactory", "Unsupported backend: " & ClassName
    End Select
End Function


Private Function IDataTableFactory_CreateInstance(ByVal model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String) As IDataTable
    IDataTableFactory_CreateInstance = CreateInstance(this.ClassName, model, ConnectionString, TableName)
End Function
