Attribute VB_Name = "DbRecordsetTests"
Attribute VB_Description = "Tests for the DbRecordset class."
'@Folder "SecureADODB.DbRecordset"
'@TestModule
'@ModuleDescription("Tests for the DbRecordset class.")
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


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetSingleParameterSelectSql() As String
    zfxGetSingleParameterSelectSql = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ?;"
End Function


Private Function zfxGetStubParameter() As ADODB.Parameter
    Dim stubAdoCommand As ADODB.Command: Set stubAdoCommand = New ADODB.Command
    Set zfxGetStubParameter = stubAdoCommand.CreateParameter("StubInteger", adInteger, adParamInput, , 42)
End Function


Private Function zfxGetStubDbCommand() As IDbCommand
    Dim stubExConnection As StubDbConnection
    Set stubExConnection = New StubDbConnection

    Dim stubBase As StubDbCommandBase
    Set stubBase = New StubDbCommandBase
    
    Set zfxGetStubDbCommand = DbCommand.Create(stubExConnection, stubBase)
End Function


'===================================================='
'================= TESTING FIXTURES ================='
'===================================================='


'@TestMethod("Test Fixture")
Private Sub zfxGetStubParameter_ValidatesStubParameter()
    On Error GoTo TestFail
    
Arrange:
    
Act:
    Dim stubParameter As ADODB.Parameter: Set stubParameter = zfxGetStubParameter
Assert:
    Assert.AreEqual "StubInteger", stubParameter.Name, "Stub ADODB.Parameter name mismatch: " & "StubInteger" & " vs. " & stubParameter.Name
    Assert.AreEqual adInteger, stubParameter.Type, "Stub ADODB.Parameter type mismatch: " & adInteger & " vs. " & stubParameter.Type
    Assert.AreEqual adParamInput, stubParameter.Direction, "Stub ADODB.Parameter direction mismatch: " & adParamInput & " vs. " & stubParameter.Direction
    Assert.AreEqual 42, stubParameter.Value, "Stub ADODB.Parameter value mismatch: " & 42 & " vs. " & stubParameter.Value
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Test Fixture")
Private Sub ztcGetSingleParameterSelectSql_ValidatesStubQuery()
    On Error GoTo TestFail
    
Arrange:
    Dim Expected As String
    Expected = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ?;"
Act:
    Dim Actual As String
    Actual = zfxGetSingleParameterSelectSql
Assert:
    Assert.AreEqual Expected, Actual, "Stub query mismatch: " & Expected & " vs. " & Actual
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Test Fixture")
Private Sub zfxGetStubDbCommand_ValidatesStubDbCommand()
    On Error GoTo TestFail
    
Arrange:
    Dim Expected As String
    Expected = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ?;"
Act:
    Dim stubCommand As IDbCommand
    Set stubCommand = zfxGetStubDbCommand
    
    Dim stubAdoCommand As ADODB.Command
    Set stubAdoCommand = stubCommand.AdoCommand(zfxGetSingleParameterSelectSql, zfxGetStubParameter)
Assert:
    Assert.IsNotNothing stubCommand, "GetStubDbCommand command did not return IDbCommand"
    Assert.IsNotNothing stubAdoCommand, "GetStubDbCommand: AdoCommand was not set"
    Assert.IsTrue TypeOf stubAdoCommand Is ADODB.Command, "GetStubDbCommand: AdoCommand type mismatch"
    Assert.AreEqual ADODB.CommandTypeEnum.adCmdText, stubAdoCommand.CommandType, "GetStubDbCommand: command type mismatch"
    Assert.AreEqual Expected, stubAdoCommand.CommandText, "GetStubDbCommand: command text mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Create")
Private Sub ztcCreate_ValidatesCreationOfDisconnectedFullRecordser()
    On Error GoTo TestFail
    
Arrange:
    Dim Recordset As IDbRecordset
    Set Recordset = DbRecordset.Create(zfxGetStubDbCommand)
Act:
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = Recordset.AdoRecordset(vbNullString)
Assert:
    AssertExpectedError Assert, ErrNo.PassedNoErr
    Assert.AreNotEqual 1, AdoRecordset.MaxRecords, "Regular recordset should have MaxRecords=0 or >1"
    Assert.AreEqual ADODB.CursorLocationEnum.adUseClient, AdoRecordset.CursorLocation, "CursorLocation should be set to adUseClient for a disconnected recordset."
    Assert.AreEqual 10, AdoRecordset.CacheSize, "Expected CacheSize=10"
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenStatic, AdoRecordset.CursorType, "Expectec CursorType=adOpenStatic for a disconnected recordset."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Create")
Private Sub ztcCreate_ValidatesCreationOfDisconnectedScalarRecordser()
    On Error GoTo TestFail
    
Arrange:
    Dim Recordset As IDbRecordset
    Set Recordset = DbRecordset.Create(zfxGetStubDbCommand, True)
Act:
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = Recordset.AdoRecordset(vbNullString)
Assert:
    AssertExpectedError Assert, ErrNo.PassedNoErr
    Assert.AreEqual 1, AdoRecordset.MaxRecords, "Scalar recordset should have MaxRecords=1"
    Assert.AreEqual ADODB.CursorLocationEnum.adUseClient, AdoRecordset.CursorLocation, "CursorLocation should be set to adUseClient for a disconnected recordset."
    Assert.AreEqual 10, AdoRecordset.CacheSize, "Expected CacheSize=10"
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenStatic, AdoRecordset.CursorType, "Expected CursorType=adOpenStatic for a disconnected recordset."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Create")
Private Sub ztcCreate_ValidatesCreationOfOnlineFullRecordser()
    On Error GoTo TestFail
    
Arrange:
    Dim Recordset As IDbRecordset
    Set Recordset = DbRecordset.Create(zfxGetStubDbCommand, , False)
Act:
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = Recordset.AdoRecordset(vbNullString)
Assert:
    AssertExpectedError Assert, ErrNo.PassedNoErr
    Assert.AreNotEqual 1, AdoRecordset.MaxRecords, "Regular recordset should have MaxRecords=0 or >1"
    Assert.AreEqual ADODB.CursorLocationEnum.adUseServer, AdoRecordset.CursorLocation, "CursorLocation should be set to adUseServer for an online recordset."
    Assert.AreEqual 10, AdoRecordset.CacheSize, "Expected CacheSize=10"
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenForwardOnly, AdoRecordset.CursorType, "Expected CursorType=adOpenForwardOnly for a disconnected recordset."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

