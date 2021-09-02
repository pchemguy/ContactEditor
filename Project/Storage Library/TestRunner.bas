Attribute VB_Name = "TestRunner"
'@IgnoreModule
'@Folder "Storage Library"
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
    Dim StorageTableModel As DataTableModel
    Set StorageTableModel = New DataTableModel
    Dim ClassName As String
    ClassName = "Worksheet"
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    Dim TableName As String
    TableName = "Contacts1"
    
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableFactory.CreateInstance(ClassName, StorageTableModel, ConnectionString, TableName)
    
    StorageManager.LoadDataIntoModel
End Sub


Private Sub TestSheetRecord()
    Dim StorageRecordModel As DataRecordModel
    Set StorageRecordModel = New DataRecordModel
    Dim ClassName As String
    ClassName = "Worksheet"
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    Dim TableName As String
    TableName = ContactBrowser.Name
    
    Dim StorageManager As IDataRecordStorage
    Set StorageManager = DataRecordFactory.CreateInstance(ClassName, StorageRecordModel, ConnectionString, TableName)
    
    StorageManager.LoadDataIntoModel
End Sub


Private Sub TestDataRecordManager()
    Dim ClassName As String
    ClassName = "Worksheet"
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    Dim TableName As String
    TableName = ContactBrowser.Name
    
    Dim Storman As IDataRecordManager
    Set Storman = DataRecordManager.Create(ClassName, ConnectionString, TableName)
    
    Storman.LoadDataIntoModel
End Sub


Private Sub TestDataTableManager()
    Dim ClassName As String
    ClassName = "Worksheet"
    Dim TableName As String
    TableName = Contacts.Name
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    
    Dim Storman As IDataTableManager
    Set Storman = DataTableManager.Create(ClassName, ConnectionString, TableName)
    
    Storman.LoadDataIntoModel
End Sub


Private Sub TestArray()
    Dim Target As Variant
    Target = Application.WorksheetFunction.Index(Range("Contacts").Value, 1)
    Target = Application.WorksheetFunction.Transpose(Range("ContactsIds").Value)
End Sub


Private Sub TestDataCompositeManager()
    Dim Storman As DataCompositeManager
    Set Storman = New DataCompositeManager

    Dim ClassName As String
    ClassName = "Worksheet"
    Dim TableName As String
    TableName = Contacts.Name
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    Storman.InitTable ClassName, ConnectionString, TableName
    
    TableName = ContactBrowser.Name
    ConnectionString = ThisWorkbook.Name
    Storman.InitRecord ClassName, ConnectionString, TableName
    
    Storman.LoadDataIntoModel
    'Storman.LoadRecordFromTable "10"
    
    Storman.UpdateRecordToTable
    
    'Storman.LoadDataIntoModel
End Sub


Private Sub TestUpdate()
    Playground.Range("A1:E1").Value = Array(1, 2, 3, 4, 5)
    Dim Val As Variant
    Dim ColumnIndex As Long
    ColumnIndex = 3 - 1
    Val = Application.WorksheetFunction.Transpose(Range("Contacts").Offset(1, ColumnIndex).Resize(Range("Contacts").Rows.Count - 1, 1))
End Sub


Private Sub TestEnum()
    Dim TestVar As ADODB.DataTypeEnum
    Debug.Print VarType(TestVar)
End Sub
