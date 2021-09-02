Attribute VB_Name = "DataRecordFactoryTests"
Attribute VB_Description = "Tests for the DbRecordset class."
'@ModuleDescription "Tests for the DbRecordset class."
'@Folder "Storage Library.Record.Factory"
'@TestModule
'@IgnoreModule AssignmentNotUsed, VariableNotUsed, LineLabelNotUsed, UnhandledOnErrorResumeNext
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
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("DataRecordStorage")
Private Sub ztcCreateInstance_ValidatesCreationOfDataStorage()
    On Error GoTo TestFail

Arrange:
    Dim StorageModel As DataRecordModel
    Set StorageModel = New DataRecordModel
    Dim ClassName As String
    ClassName = "Worksheet"
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    Dim TableName As String
    TableName = ActiveSheet.Name
Act:
    Dim StorageManager As IDataRecordStorage
    Set StorageManager = DataRecordFactory.CreateInstance(ClassName, StorageModel, ConnectionString, TableName)
Assert:
    Assert.IsNotNothing StorageManager, "DataRecordWSheet creation error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataRecordStorage")
Private Sub ztcCreateInstance_ThrowsOnUnsupportedDataStorage()
    On Error Resume Next

Arrange:
    Dim StorageModel As DataRecordModel
    Set StorageModel = New DataRecordModel
    Dim ClassName As String
    ClassName = "BadBackend"
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Name
    Dim TableName As String
    TableName = ActiveSheet.Name
Act:
    Dim StorageManager As IDataRecordStorage
    Set StorageManager = DataRecordFactory.CreateInstance(ClassName, StorageModel, ConnectionString, TableName)
Assert:
    AssertExpectedError Assert, ErrNo.NotImplementedErr

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
