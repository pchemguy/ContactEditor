Attribute VB_Name = "DbParametersTests"
'@Folder("SecureADODB.DbParameters")
'@TestModule
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule LineLabelNotUsed: Using test template
'@IgnoreModule AssignmentNotUsed: Dummy assignments when testing for error
'@IgnoreModule UnhandledOnErrorResumeNext: Testing for error
'@IgnoreModule VariableNotUsed: Testing for error
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


Public Function zfxGetSUT() As IDbParameters
    Set zfxGetSUT = DbParameters.Create(zfxGetDefaultMappings)
End Function


Public Function zfxGetAdoCommandWith2PlaceHolders() As ADODB.Command
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM people WHERE id <= ? AND last_name <> ?"
    AdoCommand.CommandText = SQLQuery
    Set zfxGetAdoCommandWith2PlaceHolders = AdoCommand
End Function


Public Function zfxGetDefaultMappings() As ITypeMap
    Set zfxGetDefaultMappings = AdoTypeMappings.Default
End Function


'===================================================='
'================= TESTING FIXTURES ================='
'===================================================='

'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Process Values")
Private Sub ztcFromValue_VerifiesPreparedPropertiesForInteger()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParamProps As Scripting.Dictionary
    Set AdoParamProps = sut.DebugFromValue(19, "Age")
Assert:
    Assert.AreEqual "Age", AdoParamProps("Name"), "Name mismatch"
    Assert.AreEqual adInteger, AdoParamProps("Type"), "Type mismatch"
    Assert.AreEqual adParamInput, AdoParamProps("Direction"), "Direction mismatch"
    Assert.AreEqual 0, AdoParamProps("Size"), "Size mismatch"
    Assert.AreEqual 19, AdoParamProps("Value"), "Value property mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Process Values")
Private Sub ztcFromValue_VerifiesPreparedPropertiesForReal()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParamProps As Scripting.Dictionary
    Set AdoParamProps = sut.DebugFromValue(19, "Age", "Double")
Assert:
    Assert.AreEqual adDouble, AdoParamProps("Type"), "Type mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Process Values")
Private Sub ztcFromValue_VerifiesPreparedPropertiesForString()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParamProps As Scripting.Dictionary
    Set AdoParamProps = sut.DebugFromValue("Age")
Assert:
    Assert.AreEqual adVarWChar, AdoParamProps("Type"), "Type mismatch"
    Assert.AreEqual 4, AdoParamProps("Size"), "Size mismatch"
    Assert.AreEqual "Age", AdoParamProps("Value"), "Value property mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Process Values")
Private Sub ztcFromValue_VerifiesPreparedPropertiesForBlank()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParamProps As Scripting.Dictionary
    Set AdoParamProps = sut.DebugFromValue(vbNullString)
Assert:
    Assert.AreEqual adVarWChar, AdoParamProps("Type"), "Type mismatch"
    Assert.AreEqual 1, AdoParamProps("Size"), "Size mismatch"
    Assert.AreEqual vbNullString, AdoParamProps("Value"), "Value property mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Process Values")
Private Sub ztcFromValue_VerifiesPreparedPropertiesForEmpty()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParamProps As Scripting.Dictionary
    Set AdoParamProps = sut.DebugFromValue(Null)
Assert:
    Assert.AreEqual vbNullString, AdoParamProps("Name"), "Name mismatch"
    Assert.AreEqual adVarChar, AdoParamProps("Type"), "Type mismatch"
    Assert.AreEqual 1, AdoParamProps("Size"), "Size mismatch"
    Assert.IsTrue IsNull(AdoParamProps("Value")), "Value property mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Create ADODB Parameter")
Private Sub ztcCreateParameter_VerifiesPropertiesForInteger()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParam As ADODB.Parameter
    Set AdoParam = sut.CreateParameter(19, "Age")
Assert:
    Assert.AreEqual "Age", AdoParam.Name, "Name mismatch"
    Assert.AreEqual adInteger, AdoParam.Type, "Type mismatch"
    Assert.AreEqual adParamInput, AdoParam.Direction, "Direction mismatch"
    Assert.AreEqual 0, AdoParam.Size, "Size mismatch"
    Assert.AreEqual 19, AdoParam.Value, "Value property mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Create ADODB Parameter")
Private Sub ztcCreateParameter_VerifiesPropertiesForReal()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParam As ADODB.Parameter
    Set AdoParam = sut.CreateParameter(19, "Age", "Double")
Assert:
    Assert.AreEqual adDouble, AdoParam.Type, "Type mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Create ADODB Parameter")
Private Sub ztcCreateParameter_VerifiesPropertiesForString()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParam As ADODB.Parameter
    Set AdoParam = sut.CreateParameter("Age")
Assert:
    Assert.AreEqual adVarWChar, AdoParam.Type, "Type mismatch"
    Assert.AreEqual 4, AdoParam.Size, "Size mismatch"
    Assert.AreEqual "Age", AdoParam.Value, "Value property mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Create ADODB Parameter")
Private Sub ztcCreateParameter_VerifiesPropertiesForBlank()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParam As ADODB.Parameter
    Set AdoParam = sut.CreateParameter(vbNullString)
Assert:
    Assert.AreEqual adVarWChar, AdoParam.Type, "Type mismatch"
    Assert.AreEqual 1, AdoParam.Size, "Size mismatch"
    Assert.AreEqual vbNullString, AdoParam.Value, "Value property mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Create ADODB Parameter")
Private Sub ztcCreateParameter_VerifiesPropertiesForEmpty()
    On Error GoTo TestFail

Arrange:
    Dim sut As DbParameters
    Set sut = zfxGetSUT
Act:
    Dim AdoParam As ADODB.Parameter
    Set AdoParam = sut.CreateParameter(Null)
Assert:
    Assert.AreEqual adVarChar, AdoParam.Type, "Type mismatch"
    Assert.AreEqual 1, AdoParam.Size, "Size mismatch"
    Assert.IsTrue IsNull(AdoParam.Value), "Value property mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Validate parameter values")
Private Sub ztcValidateParameterValues_ThrowsGivenNoArgs()
    On Error Resume Next
    Dim sut As DbParameters
    Set sut = zfxGetSUT
    Dim cmd As ADODB.Command
    Set cmd = zfxGetAdoCommandWith2PlaceHolders()
    Dim ValidateStatus As Boolean
    ValidateStatus = sut.ValidateParameterValues(cmd)
    Guard.AssertExpectedError Assert, ErrNo.CustomErr
End Sub


'@TestMethod("Validate parameter values")
Private Sub ztcValidateParameterValues_ThrowsGivenValueSQLMismatch()
    On Error Resume Next
    Dim sut As DbParameters
    Set sut = zfxGetSUT
    Dim cmd As ADODB.Command
    Set cmd = zfxGetAdoCommandWith2PlaceHolders()
    Dim ValidateStatus As Boolean
    ValidateStatus = sut.ValidateParameterValues(cmd, 1)
    Guard.AssertExpectedError Assert, ErrNo.CustomErr
End Sub


'@TestMethod("Create ADODB Parameters")
Private Sub ztcFromValues_VerifiesCreateTwoParams()
    On Error GoTo TestFail

Arrange:
    Dim sut As IDbParameters
    Set sut = zfxGetSUT
    Dim cmd As ADODB.Command
    Set cmd = zfxGetAdoCommandWith2PlaceHolders()
Act:
    sut.FromValues cmd, 19, "Age"
    Dim AdoParams As ADODB.Parameters
    Set AdoParams = cmd.Parameters
Assert:
    Assert.AreEqual 2, AdoParams.Count, "Parameter count mismatch"
    Assert.AreEqual 19, AdoParams(0).Value, "Parameter #1 value mismatch"
    Assert.AreEqual adInteger, AdoParams(0).Type, "Parameter #1 type mismatch"
    Assert.AreEqual "Age", AdoParams(1).Value, "Parameter #2 value mismatch"
    Assert.AreEqual adVarWChar, AdoParams(1).Type, "Parameter #2 type mismatch"
    Assert.AreEqual 4, AdoParams(1).Size, "Parameter #2 size mismatch"
    Assert.AreEqual adParamInput, AdoParams(1).Direction, "Parameter #2 direction mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Create ADODB Parameters")
Private Sub ztcFromValues_VerifiesUpdateTwoParams()
    On Error GoTo TestFail

Arrange:
    Dim sut As IDbParameters
    Set sut = zfxGetSUT
    Dim cmd As ADODB.Command
    Set cmd = zfxGetAdoCommandWith2PlaceHolders()
Act:
    sut.FromValues cmd, 19, Null
    sut.FromValues cmd, Null, vbNullString
    Dim AdoParams As ADODB.Parameters
    Set AdoParams = cmd.Parameters
Assert:
    Assert.AreEqual AdoParams(0).Type, AdoTypeMappings.DefaultNullMapping, "Parameter #1 type mismatch"
    Assert.IsTrue IsNull(AdoParams(0).Value), "Parameter #1 value mismatch"
    Assert.AreEqual vbNullString, AdoParams(1).Value, "Parameter #2 value mismatch"
    Assert.AreEqual adVarWChar, AdoParams(1).Type, "Parameter #2 type mismatch"
    Assert.AreEqual 1, AdoParams(1).Size, "Parameter #2 size mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcGetSQL_VerifiesQuery()
    On Error GoTo TestFail

Arrange:
    Dim sut As IDbParameters
    Set sut = zfxGetSUT
    Dim cmd As ADODB.Command
    Set cmd = zfxGetAdoCommandWith2PlaceHolders()
    sut.FromValues cmd, 19, "John"
    Dim Expected As String
    Expected = "SELECT * FROM people WHERE id <= 19 AND last_name <> 'John'"
Act:
    Dim Actual As String
    Actual = sut.GetSQL(cmd)
Assert:
    Assert.AreEqual Expected, Actual, "SQLQuery text mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcGetSQL_VerifiesQueryWithNull()
    On Error GoTo TestFail

Arrange:
    Dim sut As IDbParameters
    Set sut = zfxGetSUT
    Dim cmd As ADODB.Command
    Set cmd = zfxGetAdoCommandWith2PlaceHolders()
    sut.FromValues cmd, True, Null
    Dim Expected As String
    Expected = "SELECT * FROM people WHERE id <= True AND last_name <> Null"
Act:
    Dim Actual As String
    Actual = sut.GetSQL(cmd)
Assert:
    Assert.AreEqual Expected, Actual, "SQLQuery text mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
