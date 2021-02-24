Attribute VB_Name = "TestRunner"
'@Folder("Storage")
'@IgnoreModule ProcedureNotUsed
Option Explicit


'Private Sub TestSheetTable()
'    Dim DataModel As DataTableModel
'    Set DataModel = New DataTableModel
'
'    Dim ClassName As String
'    ClassName = "SheetTable"
'
'    Dim DataStructureFactory As IDataTableFactory
'    Set DataStructureFactory = DataTableFactory.Create(ClassName)
'
'    Dim DataStructure As IDataTable
'    Set DataStructure = DataStructureFactory.CreateInstance(DataModel)
'
'    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & CodeNameData.Name
'    Dim TableName As String: TableName = "CodeNameTable"
'    DataStructure.GetData ConnectionString, TableName
'End Sub


'Private Sub TestMapTable()
'    Dim DataModel As DataMapModel
'    Set DataModel = New DataMapModel
'
'    Dim ClassName As String
'    ClassName = "SheetMap"
'
'    Dim DataStructureFactory As IDataMapFactory
'    Set DataStructureFactory = DataMapFactory.Create(ClassName)
'
'    Dim DataStructure As IDataMap
'    Set DataStructure = DataStructureFactory.CreateInstance(DataModel)
'
'    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & RecordEditor.Name
'    DataStructure.GetData ConnectionString
'End Sub


Private Sub TestSheetTable()
    Dim StorageTableModel As DataTableModel: Set StorageTableModel = New DataTableModel
    Dim ClassName As String: ClassName = "SheetTable"
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & CodeNameData.Name
    Dim TableName As String: TableName = "CodeNameTable"
    
    Dim StorageManager As IDataTable
    Set StorageManager = DataTableFactory.CreateInstance(ClassName, StorageTableModel, ConnectionString, TableName)
    
    StorageManager.GetData
End Sub


Private Sub TestSheetMap()
    Dim StorageMapModel As DataMapModel: Set StorageMapModel = New DataMapModel
    Dim ClassName As String: ClassName = "SheetMap"
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & RecordEditor.Name
    
    Dim StorageManager As IDataMap
    Set StorageManager = DataMapFactory.CreateInstance(ClassName, StorageMapModel, ConnectionString)
    
    StorageManager.GetData
End Sub

