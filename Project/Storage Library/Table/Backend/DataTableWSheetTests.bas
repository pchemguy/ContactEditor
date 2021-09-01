Attribute VB_Name = "DataTableWSheetTests"
'@Folder "Storage Library.Table.Backend"
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
End Sub


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetDataTableModel() As DataTableModel
    Dim StorageModel As DataTableModel: Set StorageModel = New DataTableModel
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & TestSheet.Name
    Dim TableName As String: TableName = "TestContacts"
        
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableWSheet.Create(StorageModel, ConnectionString, TableName)
    StorageManager.LoadDataIntoModel
    Set zfxGetDataTableModel = StorageModel
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("DataTableWSheet")
Private Sub ztcCreate_ValidatesCreationOfDataStorage()
    On Error GoTo TestFail

Arrange:
    Dim StorageModel As DataTableModel: Set StorageModel = New DataTableModel
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & ActiveSheet.Name
    Dim TableName As String: TableName = "TestContacts"
Act:
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableWSheet.Create(StorageModel, ConnectionString, TableName)
Assert:
    Assert.IsNotNothing StorageManager, "DataTableWSheet creation error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTableWSheet")
Private Sub ztcCreate_ThrowsOnInavlidConnectionString()
    On Error Resume Next

Arrange:
    Dim StorageModel As DataTableModel: Set StorageModel = New DataTableModel
    Dim ConnectionString As String: ConnectionString = "InvalidConnectionString"
    Dim TableName As String: TableName = "TestContacts"
Act:
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableWSheet.Create(StorageModel, ConnectionString, TableName)
Assert:
    AssertExpectedError Assert, ErrNo.CustomErr

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTableWSheet")
Private Sub ztcCreate_ThrowsOnInavlidExcelObjectNames()
    On Error Resume Next

Arrange:
    Dim StorageModel As DataTableModel: Set StorageModel = New DataTableModel
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "?!?" & ActiveSheet.Name
    Dim TableName As String: TableName = "TestContacts"
Act:
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableWSheet.Create(StorageModel, ConnectionString, TableName)
Assert:
    AssertExpectedError Assert, ErrNo.CustomErr

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTableWSheet")
Private Sub ztcCreate_ThrowsOnInavlidRangeNames()
    On Error Resume Next

Arrange:
    Dim StorageModel As DataTableModel: Set StorageModel = New DataTableModel
    Dim ConnectionString As String: ConnectionString = ThisWorkbook.Name & "!" & ActiveSheet.Name
    Dim TableName As String: TableName = "InvalidTableName"
Act:
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableWSheet.Create(StorageModel, ConnectionString, TableName)
Assert:
    AssertExpectedError Assert, ErrNo.CustomErr

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTableModel")
Private Sub ztcModel_ValidatesLoadedData()
    On Error GoTo TestFail

Arrange:
Act:
    Dim StorageModel As DataTableModel: Set StorageModel = zfxGetDataTableModel
Assert:
    With StorageModel
        Assert.IsNotNothing .DirtyRecords, "Dirty records dictionary is not set"
        Assert.AreEqual 0, .DirtyRecords.Count, "Dirty records count should be 0"
        Assert.IsFalse .IsDirty, "Model should not be dirty"
        
        Assert.IsNotNothing .FieldIndices, "FieldIndices dictionary is not set"
        Assert.AreEqual 8, .FieldIndices.Count, "FieldIndices - wrong field count"
        Assert.IsTrue .FieldIndices.Exists("TestEmail"), "FieldIndices - missing field"
        Assert.AreEqual 6, .FieldIndices("TestEmail"), "FieldIndices - index mismatch"
        
        Assert.IsTrue IsArray(.FieldNames), "FieldNames is not set"
        Assert.AreEqual 1, LBound(.FieldNames, 1), "FieldNames - wrong index base"
        Assert.AreEqual 8, UBound(.FieldNames, 1), "FieldNames - wrong field count"
        Assert.AreEqual "TestEmail", .FieldNames(6), "FieldNames - item mismatch"
        
        Assert.IsNotNothing .IdIndices, "IdIndices dictionary is not set"
        Assert.AreEqual 100, .IdIndices.Count, "IdIndices - wrong record count"
        Assert.AreEqual 90, .IdIndices("90"), "IdIndices - wrong record index"
        
        Assert.IsTrue IsArray(.Values), "Values is not set"
        Assert.AreEqual 1, LBound(.Values, 1), "Values - wrong record index base"
        Assert.AreEqual 100, UBound(.Values, 1), "Values - wrong record count"
        Assert.AreEqual 1, LBound(.Values, 2), "Values - wrong field index base"
        Assert.AreEqual 8, UBound(.Values, 2), "Values - wrong field count"
        Assert.AreEqual "Edna.Jennings@neuf.fr", .Values(4, 6), "Values - field mismatch"
    End With
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
