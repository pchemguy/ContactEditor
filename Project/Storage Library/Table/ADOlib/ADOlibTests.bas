Attribute VB_Name = "ADOlibTests"
'@Folder "Storage Library.Table.ADOlib"
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess
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
'=================== TEST FIXTURES =================='
'===================================================='


Private Function zfxFieldNames() As Variant
    Dim FieldNames(1 To 5) As String
    FieldNames(1) = "id"
    FieldNames(2) = "FirstName"
    FieldNames(3) = "LastName"
    FieldNames(4) = "Age"
    FieldNames(5) = "Gender"
    zfxFieldNames = FieldNames
End Function


Private Function zfxFieldTypes() As Variant
    Dim FieldTypes(1 To 5) As ADODB.DataTypeEnum
    FieldTypes(1) = adInteger
    FieldTypes(2) = adVarWChar
    FieldTypes(3) = adVarWChar
    FieldTypes(4) = adInteger
    FieldTypes(5) = adVarWChar
    zfxFieldTypes = FieldTypes
End Function


Private Function zfxRecordValues() As Scripting.Dictionary
    Dim RecordValues As Scripting.Dictionary
    Set RecordValues = New Scripting.Dictionary
    With RecordValues
        .CompareMode = TextCompare
        .Item("id") = 4
        .Item("FirstName") = "Edna"
        .Item("LastName") = "Jennings"
        .Item("Age") = 26
        .Item("Gender") = "male"
    End With
    Set zfxRecordValues = RecordValues
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("ConnectionString")
Private Sub ztcGetSQLiteConnectionString_ValidatesDefaultString()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & _
                VerifyOrGetDefaultPath(vbNullString, Array("db", "sqlite")) & _
               ";SyncPragma=NORMAL;FKSupport=True;"
Act:
    Dim Actual As String
    Actual = ADOlib.GetSQLiteConnectionString(vbNullString)("ADO")
Assert:
    Assert.AreEqual Expected, Actual, "Default SQLite ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    If Err.Number = ErrNo.FileNotFoundErr Then
        Assert.Inconclusive "Target file not found. This test require particular settings and this error may be ignored"
    Else
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("ADO Parameters")
Private Sub ztcSetAdoParamsForRecordUpdate_ValidatesParams()
    On Error GoTo TestFail

Arrange:
Act:
    Dim AdoCommand As ADODB.Command: Set AdoCommand = New ADODB.Command
    Dim AdoParams As ADODB.Parameters: Set AdoParams = AdoCommand.Parameters
    ADOlib.MakeAdoParamsForRecordUpdate zfxFieldNames, zfxFieldTypes, AdoCommand
Assert:
    Assert.AreEqual 5, AdoParams.Count, "Parameters - wrong count"
    Assert.AreEqual "FirstName", AdoParams(0).Name, "First param name mismatch"
    Assert.AreEqual adVarWChar, AdoParams(0).Type, "First param type mismatch"
    Assert.AreEqual "id", AdoParams(4).Name, "Last param name mismatch"
    Assert.AreEqual adInteger, AdoParams(4).Type, "Last param type mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADO Parameters")
Private Sub ztcRecordValuesToAdoParams_ValidatesUpdatedParams()
    On Error GoTo TestFail

Arrange:
    Dim AdoCommand As ADODB.Command: Set AdoCommand = New ADODB.Command
    Dim AdoParams As ADODB.Parameters: Set AdoParams = AdoCommand.Parameters
    ADOlib.MakeAdoParamsForRecordUpdate zfxFieldNames, zfxFieldTypes, AdoCommand
Act:
    ADOlib.RecordToAdoParams zfxRecordValues, AdoCommand
Assert:
    Assert.AreEqual "Edna", AdoParams(0).Value, "Updated value mismatch"
    Assert.AreEqual "Jennings", AdoParams(1).Value, "Updated value mismatch"
    Assert.AreEqual 26, AdoParams(2).Value, "Updated value mismatch"
    Assert.AreEqual "male", AdoParams(3).Value, "Updated value mismatch"
    Assert.AreEqual 4, AdoParams(4).Value, "Updated value mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADO Parameters")
Private Sub ztcRecordValuesToAdoParams_ValidatesUpdatedParamsIdAsText()
    On Error GoTo TestFail

Arrange:
    Dim AdoCommand As ADODB.Command: Set AdoCommand = New ADODB.Command
    Dim AdoParams As ADODB.Parameters: Set AdoParams = AdoCommand.Parameters
    ADOlib.MakeAdoParamsForRecordUpdate zfxFieldNames, zfxFieldTypes, AdoCommand, CastIdAsText
Act:
    ADOlib.RecordToAdoParams zfxRecordValues, AdoCommand
Assert:
    Assert.AreEqual adInteger, AdoParams(2).Type, "Param type mismatch"
    Assert.AreEqual adVarWChar, AdoParams(4).Type, "Param type mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADO Parameters")
Private Sub ztcRecordValuesToAdoParams_ValidatesUpdatedParamsAllAsText()
    On Error GoTo TestFail

Arrange:
    Dim AdoCommand As ADODB.Command: Set AdoCommand = New ADODB.Command
    Dim AdoParams As ADODB.Parameters: Set AdoParams = AdoCommand.Parameters
    ADOlib.MakeAdoParamsForRecordUpdate zfxFieldNames, zfxFieldTypes, AdoCommand, CastAllAsText
Act:
    ADOlib.RecordToAdoParams zfxRecordValues, AdoCommand
Assert:
    Assert.AreEqual adVarWChar, AdoParams(2).Type, "Param type mismatch"
    Assert.AreEqual adVarWChar, AdoParams(4).Type, "Param type mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
