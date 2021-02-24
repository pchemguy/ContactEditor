Attribute VB_Name = "TestRunner"
'@Folder("Storage")
'@IgnoreModule ProcedureNotUsed
Option Explicit


Private Sub TestSheetTable()
    Dim ClassName As String
    ClassName = "SheetTable"
    
    Dim DataStructureFactory As IDataTableFactory
    Set DataStructureFactory = DataTableFactory.Create(ClassName)
    
    Dim DataModel As DataTableModel
    Set DataModel = New DataTableModel
    
    Dim DataStructure As IDataTable
    Set DataStructure = DataStructureFactory.CreateInstance(DataModel)
    
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & CodeNameData.Name
    Dim TableName As String: TableName = "TableA"
    DataStructure.GetData ConnectionString, TableName
End Sub


Private Sub TestMapTable()
    Dim ClassName As String
    ClassName = "SheetMap"
    
    Dim DataStructureFactory As IDataMapFactory
    Set DataStructureFactory = DataMapFactory.Create(ClassName)
    
    Dim DataModel As DataMapModel
    Set DataModel = New DataMapModel
    
    Dim DataStructure As IDataMap
    Set DataStructure = DataStructureFactory.CreateInstance(DataModel)
    
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & RecordEditor.Name
    Dim FieldNames As Variant: FieldNames = Array("FieldA", "FieldB")
    DataStructure.GetData ConnectionString, FieldNames
End Sub

