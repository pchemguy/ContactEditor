Attribute VB_Name = "DbRecordsetTests"
Attribute VB_Description = "Tests for the DbRecordset class."
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription "Tests for the DbRecordset class."
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext
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


Private Function zfxGetSingleParameterSelectSql() As String
    zfxGetSingleParameterSelectSql = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ?;"
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
Act:
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = Recordset.GetAdoRecordset(vbNullString)
Assert:
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
    Assert.AreNotEqual 1, AdoRecordset.MaxRecords, "Regular recordset should have MaxRecords=0 by default"
    Assert.AreEqual ADODB.CursorLocationEnum.adUseClient, AdoRecordset.CursorLocation, "CursorLocation should be set to adUseClient for a disconnected recordset."
    Assert.AreEqual 10, AdoRecordset.CacheSize, "Expected CacheSize=10"
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenStatic, AdoRecordset.CursorType, "Expectec CursorType=adOpenStatic for a disconnected recordset."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
