Attribute VB_Name = "DataCompositeManagerTests"
'@Folder "Storage Library.Manager"
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


Private Function zfxDataCompositeManager() As DataCompositeManager
    Dim Storman As DataCompositeManager
    Set Storman = New DataCompositeManager
    
    Dim ClassName As String
    Dim TableName As String
    Dim ConnectionString As String
    
    '''' Binds TableModel to its backend
    ClassName = "Worksheet"
    ConnectionString = ThisWorkbook.Name
    TableName = "TestContacts"
    Storman.InitTable ClassName, ConnectionString, TableName
    
    '''' Binds RecordModel to its backend
    ClassName = "Worksheet"
    ConnectionString = ThisWorkbook.Name
    TableName = TestContacts.Name
    Storman.InitRecord ClassName, ConnectionString, TableName
    
    Storman.LoadDataIntoModel
    Set zfxDataCompositeManager = Storman
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("DataCompositeManager")
Private Sub ztcModel_ValidatesLoadedData()
    On Error GoTo TestFail
    
Arrange:
    Dim Storman As DataCompositeManager
    Set Storman = zfxDataCompositeManager
Act:
Assert:
    With Storman
        Assert.AreEqual TestContacts.Range("TestEmail").Value, .Record("TestEmail"), "RecordModel data mismatch"
        Assert.AreEqual "Edna.Jennings@neuf.fr", .Values(4, 6), "TableModel data mismatch"
    End With
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataCompositeManager")
Private Sub ztcFieldNames_ValidatesFieldNames()
    On Error GoTo TestFail
    
Arrange:
    Dim Storman As DataCompositeManager
    Set Storman = zfxDataCompositeManager
Act:
    Dim FieldNames As Variant: FieldNames = Storman.FieldNames
Assert:
    Assert.IsTrue IsArray(FieldNames), "FieldNames is not set"
    Assert.AreEqual 1, LBound(FieldNames, 1), "FieldNames - wrong base index"
    Assert.AreEqual 8, UBound(FieldNames, 1), "FieldNames - wrong count"
    Assert.AreEqual vbString, VarType(FieldNames(1)), "FieldNames - expected array of strings"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataCompositeManager")
Private Sub ztcIds_ValidatesIds()
    On Error GoTo TestFail
    
Arrange:
    Dim Storman As DataCompositeManager
    Set Storman = zfxDataCompositeManager
Act:
    Dim IDs As Variant: IDs = Storman.IDs
Assert:
    Assert.IsTrue IsArray(IDs), "Ids is not set"
    Assert.AreEqual 1, LBound(IDs, 1), "Ids - wrong base index"
    Assert.AreEqual 100, UBound(IDs, 1), "Ids - wrong count"
    Assert.AreEqual vbString, VarType(IDs(1)), "Ids - expected array of strings"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataCompositeManager")
Private Sub ztcLoadUpdateRecordTable_ValidatesRecordTableTransfer()
    On Error GoTo TestFail
    
Arrange:
    Dim Storman As DataCompositeManager
    Set Storman = zfxDataCompositeManager
Act:
Assert:
    Storman.LoadRecordFromTable "4"
    Assert.AreEqual "Edna.Jennings@neuf.fr", Storman.Record("TestEmail"), "LoadRecordFromTable data mismatch"
    
    Storman.Record("TestEmail") = "Edna.Jennings@@neuf.fr"
    Storman.UpdateRecordToTable
    Assert.AreEqual "Edna.Jennings@@neuf.fr", Storman.Values(4, 6), "UpdateRecordToTable data mismatch"
    
    Storman.LoadRecordFromTable "4"
    Assert.AreEqual "Edna.Jennings@@neuf.fr", Storman.Record("TestEmail"), "LoadRecordFromTable data mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
