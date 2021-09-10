Attribute VB_Name = "DataTableCSVTests"
'@Folder "Storage Library.Table.Backend.Tests"
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

Private Const LIB_NAME As String = "StorageLibrary"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


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


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetDataTableModel() As DataTableModel
    Dim StorageModel As DataTableModel
    Set StorageModel = New DataTableModel
    Dim ConnectionString As String
    ConnectionString = ThisWorkbook.Path & PATH_SEP & REL_PREFIX
    Dim TableName As String
    TableName = LIB_NAME & ".xsv!sep=,"
    
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableCSV.Create(StorageModel, ConnectionString, TableName)
    StorageManager.LoadDataIntoModel
    Set zfxGetDataTableModel = StorageModel
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("DataTableCSV")
Private Sub ztcCreate_ValidatesCreationOfDataStorage()
    On Error GoTo TestFail

Arrange:
    Dim StorageModel As DataTableModel
    Set StorageModel = New DataTableModel
    Dim ConnectionString As String
    ConnectionString = vbNullString
    Dim TableName As String
    TableName = vbNullString
Act:
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableCSV.Create(StorageModel, ConnectionString, TableName)
Assert:
    Assert.IsNotNothing StorageManager, "DataTableCSV creation error"

CleanExit:
    Exit Sub
TestFail:
    If Err.Number = ErrNo.FileNotFoundErr Then
        Assert.Inconclusive "Target file not found. This test require particular settings and this error may be ignored"
    Else
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("DataTableCSV")
Private Sub ztcCreate_ThrowsOnInavlidConnectionString()
    On Error Resume Next

Arrange:
    Dim StorageModel As DataTableModel
    Set StorageModel = New DataTableModel
    Dim ConnectionString As String
    ConnectionString = "InvalidConnectionString"
    Dim TableName As String
    TableName = "TestContacts"
Act:
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableCSV.Create(StorageModel, ConnectionString, TableName)
Assert:
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTableCSV")
Private Sub ztcCreate_ThrowsOnInavlidTableName()
    On Error Resume Next

Arrange:
    Dim StorageModel As DataTableModel
    Set StorageModel = New DataTableModel
    Dim ConnectionString As String
    ConnectionString = vbNullString
    Dim TableName As String
    TableName = "TestContacts"
Act:
    Dim StorageManager As IDataTableStorage
    Set StorageManager = DataTableCSV.Create(StorageModel, ConnectionString, TableName)
Assert:
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr

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
    Dim StorageModel As DataTableModel
    Set StorageModel = zfxGetDataTableModel
Assert:
    With StorageModel
        Assert.IsNotNothing .DirtyRecords, "Dirty records dictionary is not set"
        Assert.AreEqual 0, .DirtyRecords.Count, "Dirty records count should be 0"
        Assert.IsFalse .IsDirty, "Model should not be dirty"
        
        Assert.IsNotNothing .FieldIndices, "FieldIndices dictionary is not set"
        Assert.AreEqual 8, .FieldIndices.Count, "FieldIndices - wrong field count"
        Assert.IsTrue .FieldIndices.Exists("Email"), "FieldIndices - missing field"
        Assert.AreEqual 6, .FieldIndices("Email"), "FieldIndices - index mismatch"
        
        Assert.IsTrue IsArray(.FieldNames), "FieldNames is not set"
        Assert.AreEqual 1, LBound(.FieldNames, 1), "FieldNames - wrong index base"
        Assert.AreEqual 8, UBound(.FieldNames, 1), "FieldNames - wrong field count"
        Assert.AreEqual "Email", .FieldNames(6), "FieldNames - item mismatch"
        
        Assert.IsNotNothing .IdIndices, "IdIndices dictionary is not set"
        Assert.AreEqual 1000, .IdIndices.Count, "IdIndices - wrong record count"
        Assert.AreEqual 90, .IdIndices("90"), "IdIndices - wrong record index"
        
        Assert.IsTrue IsArray(.Values), "Values is not set"
        Assert.AreEqual 1, LBound(.Values, 1), "Values - wrong record index base"
        Assert.AreEqual 1000, UBound(.Values, 1), "Values - wrong record count"
        Assert.AreEqual 1, LBound(.Values, 2), "Values - wrong field index base"
        Assert.AreEqual 8, UBound(.Values, 2), "Values - wrong field count"
        Assert.AreEqual "Edna.Jennings@neuf.fr", .Values(4, 6), "Values - field mismatch"
    End With
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
