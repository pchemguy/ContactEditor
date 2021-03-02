VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataTableStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Abstracts storage interfaces for DataTableModel. Implemented by storage backends."
'@Folder "ContactEditor.Storage.Table.Backend"
'@ModuleDescription "Abstracts storage interfaces for DataTableModel. Implemented by storage backends."
'@Interface
'@Exposed
Option Explicit


Public Sub LoadDataIntoModel()
End Sub

Public Sub SaveDataFromModel()
End Sub

Public Function GetIds() As Variant
End Function

Public Function GetColumnValues(ByVal FieldName As String) As Variant
End Function
