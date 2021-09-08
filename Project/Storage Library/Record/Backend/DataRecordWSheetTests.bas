Attribute VB_Name = "DataRecordWSheetTests"
'@Folder "Storage Library.Record.Backend"
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


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetDataRecordModel() As DataRecordModel
    Dim StorageModel As DataRecordModel
    Set StorageModel = New DataRecordModel
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    Dim TableName As String
    TableName = TestContacts.Name
    
    Dim StorageManager As IDataRecordStorage
    Set StorageManager = DataRecordWSheet.Create(StorageModel, ConnectionString, TableName)
    StorageManager.LoadDataIntoModel
    Set zfxGetDataRecordModel = StorageModel
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("DataRecordWSheet")
Private Sub ztcCreate_ValidatesCreationOfDataStorage()
    On Error GoTo TestFail

Arrange:
    Dim StorageModel As DataRecordModel
    Set StorageModel = New DataRecordModel
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    Dim TableName As String
    TableName = ActiveSheet.Name
Act:
    Dim StorageManager As IDataRecordStorage
    Set StorageManager = DataRecordWSheet.Create(StorageModel, ConnectionString, TableName)
Assert:
    Assert.IsNotNothing StorageManager, "DataRecordWSheet creation error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataRecordWSheet")
Private Sub ztcCreate_ThrowsOnInavlidConnectionString()
    On Error Resume Next

Arrange:
    Dim StorageModel As DataRecordModel
    Set StorageModel = New DataRecordModel
    Dim ConnectionString As String
    ConnectionString = "InvalidConnectionString"
    Dim TableName As String
    TableName = vbNullString
Act:
    Dim StorageManager As IDataRecordStorage
    Set StorageManager = DataRecordWSheet.Create(StorageModel, ConnectionString, TableName)
Assert:
    Guard.AssertExpectedError Assert, ErrNo.CustomErr

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataRecordWSheet")
Private Sub ztcCreate_ThrowsOnInavlidExcelObjectNames()
    On Error Resume Next

Arrange:
    Dim StorageModel As DataRecordModel
    Set StorageModel = New DataRecordModel
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name & "?!?"
    Dim TableName As String
    TableName = ActiveSheet.Name
Act:
    Dim StorageManager As IDataRecordStorage
    Set StorageManager = DataRecordWSheet.Create(StorageModel, ConnectionString, TableName)
Assert:
    Guard.AssertExpectedError Assert, ErrNo.CustomErr

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataRecordModel")
Private Sub ztcModel_ValidatesLoadedData()
    On Error GoTo TestFail

Arrange:
Act:
    Dim StorageModel As DataRecordModel
    Set StorageModel = zfxGetDataRecordModel
Assert:
    Assert.AreEqual 8, StorageModel.Record.Count, "Field count mismatch"
    Assert.IsTrue StorageModel.Record.Exists("Testid"), "Expected field not present"
    Assert.AreEqual "Edna.Jennings@neuf.fr", StorageModel.Record("TestEmail"), "Field value mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataRecordModel")
Private Sub ztcModel_ValidatesDirtyStatus()
    On Error GoTo TestFail

    Dim StorageModel As DataRecordModel: Set StorageModel = zfxGetDataRecordModel
    Assert.IsFalse StorageModel.IsDirty, "Model should not be dirty"
    StorageModel.SetField "TestEmail", "Edna.Jennings@@neuf.fr"
    Assert.IsTrue StorageModel.IsDirty, "Model should be dirty"
    Assert.AreEqual "Edna.Jennings@@neuf.fr", StorageModel.GetField("TestEmail"), "Set/Get field error"
        
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
