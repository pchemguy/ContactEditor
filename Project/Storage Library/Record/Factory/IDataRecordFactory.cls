VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataRecordFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Abstract factory for record storage objects."
'@Folder "Storage Library.Record.Factory"
'@ModuleDescription "Abstract factory for record storage objects."
'@Interface
'@Exposed
Option Explicit


Public Function CreateInstance(ByVal Model As DataRecordModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataRecordStorage
End Function
