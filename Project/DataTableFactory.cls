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


Private Function IDataTableFactory_CreateInstance(ByVal Model As DataTableModel) As IDataTable
    Select Case this.ClassName
        Case "SheetTable"
            Set IDataTableFactory_CreateInstance = SheetTable.Create(Model)
        Case Else
            VBA.Err.Raise vbObjectError, "IDataTableFactory", "Unsupported backend: " & this.ClassName
    End Select
End Function
