Attribute VB_Name = "DbRecordsetTests"
Attribute VB_Description = "Tests for the DbRecordset class."
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription "Tests for the DbRecordset class."
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, VariableNotUsed, AssignmentNotUsed
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


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


''Private Function zfxGetSingleParameterSelectSQL() As String
''    zfxGetSingleParameterSelectSQL = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ?;"
''End Function


Private Function zfxGetTwoParameterSelectSQL() As String
    zfxGetTwoParameterSelectSQL = "SELECT * FROM people WHERE age >= ? AND country = ?"
End Function


Private Function zfxGetStubDbCommand() As IDbCommand
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection

    Dim stubCommandBase As StubDbCommandBase
    Set stubCommandBase = New StubDbCommandBase
    
    Set zfxGetStubDbCommand = DbCommand.Create(stubConnection, stubCommandBase)
End Function


'===================================================='
'================= TESTING FIXTURES ================='
'===================================================='


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("AdoRecordset")
Private Sub ztcGetAdoRecordset_ValidatesDefaultAdoRecordset()
    On Error GoTo TestFail
    
Arrange:
    Dim Recordset As IDbRecordset
    Set Recordset = DbRecordset.Create(zfxGetStubDbCommand)
    Dim SQLQuery As String
    SQLQuery = zfxGetTwoParameterSelectSQL
Act:
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = Recordset.GetAdoRecordset(SQLQuery)
Assert:
    Assert.AreNotEqual 1, AdoRecordset.MaxRecords, "Regular recordset should have MaxRecords=0 by default"
    Assert.AreEqual ADODB.CursorLocationEnum.adUseClient, AdoRecordset.CursorLocation, "CursorLocation should be set to adUseClient for a disconnected recordset."
    Assert.AreEqual 10, AdoRecordset.CacheSize, "Expected CacheSize=10"
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenStatic, AdoRecordset.CursorType, "Expectec CursorType=adOpenStatic for a disconnected recordset."
    Assert.AreEqual SQLQuery, AdoRecordset.Source, "SQL query mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("UpdateRecord")
Private Sub ztcUpdateRecord_ThrowsIfWrongLockType()
    On Error Resume Next
    Dim Recordset As IDbRecordset
    Set Recordset = DbRecordset.Create(zfxGetStubDbCommand)
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = Recordset.GetAdoRecordset(vbNullString)
    Dim ValuesDict As Scripting.Dictionary
    Set ValuesDict = New Scripting.Dictionary
    Recordset.UpdateRecord 0, ValuesDict
    Guard.AssertExpectedError Assert, ErrNo.AdoFeatureNotAvailableErr
End Sub


'@TestMethod("UpdateRecord")
Private Sub ztcUpdateRecord_ThrowsIfRecordSetIsClosed()
    On Error Resume Next
    Dim Recordset As IDbRecordset
    Set Recordset = DbRecordset.Create(zfxGetStubDbCommand)
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = Recordset.GetAdoRecordset(vbNullString)
    AdoRecordset.LockType = adLockBatchOptimistic
    Dim ValuesDict As Scripting.Dictionary
    Set ValuesDict = New Scripting.Dictionary
    Recordset.UpdateRecord 0, ValuesDict
    Guard.AssertExpectedError Assert, ErrNo.IncompatibleStatusErr
End Sub


'@TestMethod("UpdateRecord")
Private Sub ztcUpdateRecord_ThrowsIfValuesDictNotSet()
    On Error Resume Next
    Dim Recordset As IDbRecordset
    Set Recordset = DbRecordset.Create(zfxGetStubDbCommand)
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = Recordset.GetAdoRecordset(vbNullString)
    AdoRecordset.LockType = adLockBatchOptimistic
    Recordset.UpdateRecord 0, Nothing
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub
