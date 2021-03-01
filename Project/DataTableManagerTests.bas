Attribute VB_Name = "DataTableManagerTests"
'@Folder("ContactEditor.Storage.Table.Manager")
'@TestModule
'@IgnoreModule AssignmentNotUsed, VariableNotUsed, LineLabelNotUsed, UnhandledOnErrorResumeNext, IndexedDefaultMemberAccess
Option Explicit
Option Private Module


#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Guard = Nothing
End Sub


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetDataTableManager() As DataTableManager
    Dim ClassName As String: ClassName = "Worksheet"
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & TestSheet.Name
    Dim TableName As String: TableName = "TestContacts"
    
    Dim Storman As IDataTableManager
    Set Storman = DataTableManager.Create(ClassName, ConnectionString, TableName)
    
    Storman.LoadDataIntoModel
    Set zfxGetDataTableManager = Storman
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("DataTableManager")
Private Sub ztcModel_ValidatesLoadedData()
    On Error GoTo TestFail
    
Arrange:
    Dim Storman As IDataTableManager
    Set Storman = zfxGetDataTableManager
Act:
    Dim StorageModel As DataTableModel: Set StorageModel = Storman.Model
Assert:
    With StorageModel
        Assert.AreEqual 6, .FieldIndexFromName("TestEmail"), "FieldIndexFromName - field wrong index"
        Assert.AreEqual 90, .RecordIndexFromId("90"), "RecordIndexFromId - wrong record index"
    End With
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("DataTableModel")
Private Sub ztcRecordValuesFromId_ValidatesLoadedData()
    On Error GoTo TestFail
    
Arrange:
    Dim Storman As IDataTableManager
    Set Storman = zfxGetDataTableManager
    Dim StorageModel As DataTableModel: Set StorageModel = Storman.Model
Act:
    Dim Record As Variant: Record = StorageModel.RecordValuesFromId("4")
Assert:
    With StorageModel
        Assert.IsTrue IsArray(Record), "Did not return record"
        Assert.AreEqual 1, LBound(Record), "Wrong index base"
        Assert.AreEqual 8, UBound(Record), "Wrong field count"
        Assert.AreEqual "Edna.Jennings@neuf.fr", Record(6), "Wrong field value"
    End With
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("DataTableModel")
Private Sub ztcCopyRecordToDictionary_ValidatesCopyToDict()
    On Error GoTo TestFail
    
Arrange:
    Dim Storman As IDataTableManager
    Set Storman = zfxGetDataTableManager
    Dim StorageModel As DataTableModel: Set StorageModel = Storman.Model
    Dim Record As Scripting.Dictionary
    Set Record = New Scripting.Dictionary
Act:
    StorageModel.CopyRecordToDictionary Record, "4"
Assert:
    With StorageModel
        Assert.AreEqual 8, Record.Count, "Wrong field count"
        Assert.IsTrue Record.Exists("TestEmail"), "Missing field"
        Assert.AreEqual "Edna.Jennings@neuf.fr", Record("TestEmail"), "Wrong field value"
    End With
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("DataTableModel")
Private Sub ztcUpdateRecordFromDictionary_ValidatesCopyFromDict()
    On Error GoTo TestFail
    
Arrange:
    Dim Storman As IDataTableManager
    Set Storman = zfxGetDataTableManager
    Dim StorageModel As DataTableModel: Set StorageModel = Storman.Model
    Dim Record As Scripting.Dictionary
    Set Record = New Scripting.Dictionary
Act:
    StorageModel.CopyRecordToDictionary Record, "4"
    Record("TestEmail") = "Edna.Jennings@@neuf.fr"
    StorageModel.UpdateRecordFromDictionary Record
Assert:
    With StorageModel
        Assert.IsTrue .IsDirty, "DataTable must be dirty"
        Assert.AreEqual "Edna.Jennings@@neuf.fr", .Values(4, 6), "Wrong field value"
        Assert.AreEqual 1, .DirtyRecords.Count, "Wrong number of dirty records"
        Assert.AreEqual 4, .DirtyRecords("4"), "Wrong entry in dirty records"
    End With
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("DataTableManager")
Private Sub ztcSaveDataFromModel_ValidatesBackendDataSaving()
    On Error GoTo TestFail
    
Arrange:
    ActiveSheet.Range("TestContactsBody").Range("B11").Value = "Hannah11"
    ActiveSheet.Range("TestContactsBody").Range("G22").Value = "Ukraine22"
    
    Dim Storman As IDataTableManager
    Set Storman = zfxGetDataTableManager
    Dim StorageModel As DataTableModel: Set StorageModel = Storman.Model
    Dim Record As Scripting.Dictionary
    Set Record = New Scripting.Dictionary
Act:
Assert:
    With StorageModel
        .CopyRecordToDictionary Record, "11"
        Assert.AreEqual "Hannah11", Record("TestFirstName"), "Wrong field value"
        Record("TestFirstName") = "Hannah"
        .UpdateRecordFromDictionary Record
        
        .CopyRecordToDictionary Record, "22"
        Assert.AreEqual "Ukraine22", Record("TestCountry"), "Wrong field value"
        Record("TestCountry") = "Ukraine"
        .UpdateRecordFromDictionary Record
    
        Assert.IsTrue .IsDirty, "Table should be dirty"
        Assert.AreEqual "Hannah", .Values(11, 2), "Wrong field value"
        Assert.AreEqual "Ukraine", .Values(22, 7), "Wrong field value"
        Assert.AreEqual 2, .DirtyRecords.Count, "Dirty records count mismatch"
        Assert.AreEqual 11, .DirtyRecords("11"), "Dirty records mismatch/missing"
        Assert.AreEqual 22, .DirtyRecords("22"), "Dirty records mismatch/missing"
    End With
    
    Storman.SaveDataFromModel
    Assert.IsFalse StorageModel.IsDirty, "Table should not be dirty"
    Assert.AreEqual 0, StorageModel.DirtyRecords.Count, "Dirty records should be empty"
    Assert.AreEqual "Hannah", ActiveSheet.Range("TestContactsBody").Range("B11").Value, "Wrong saved field"
    Assert.AreEqual "Ukraine", ActiveSheet.Range("TestContactsBody").Range("G22").Value, "Wrong saved field"
        
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub
