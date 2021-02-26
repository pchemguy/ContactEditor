VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataTableFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Abstract factory for storage objects."
'@Folder "Storage.Table.Factory"
'@ModuleDescription "Abstract factory for storage objects."
'@Interface
'@Exposed
Option Explicit


Public Function CreateInstance(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String) As IDataTableStorage
End Function
