Attribute VB_Name = "ADOlibTests"
'@Folder "ContactEditor.Storage.Table.ADOlib"
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


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


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
               ";SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
Act:
    Dim Actual As String
    Actual = ADOlib.GetSQLiteConnectionString(vbNullString)("ADO")
Assert:
    Assert.AreEqual Expected, Actual, "Default SQLite ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("ADO Parameters")
Private Sub ztcSetAdoParamsForRecordUpdate_ValidatesParams()
    On Error GoTo TestFail

Arrange:
    Dim FieldNames(1 To 5) As String
    Dim FieldTypes(1 To 5) As ADODB.DataTypeEnum
    FieldNames(1) = "id":        FieldTypes(1) = adInteger
    FieldNames(2) = "FirstName": FieldTypes(2) = adVarWChar
    FieldNames(3) = "LastName":  FieldTypes(3) = adVarWChar
    FieldNames(4) = "Age":       FieldTypes(4) = adInteger
    FieldNames(5) = "Gender":    FieldTypes(5) = adVarWChar
Act:
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    Dim AdoParamsCol As ADODB.Parameters
    Set AdoParamsCol = AdoCommand.Parameters
    Dim AdoParamsArray As Variant
    AdoParamsArray = ADOlib.SetAdoParamsForRecordUpdate(FieldNames, FieldTypes, AdoParamsCol)
Assert:
    Assert.IsTrue VarType(AdoParamsArray) = vbArray + vbObject, "AdoParamsArray - wrong type"
    Assert.AreEqual 1, LBound(AdoParamsArray), "AdoParamsArray - wrong base"
    Assert.AreEqual 5, UBound(AdoParamsArray), "AdoParamsArray - wrong count"
    Assert.IsTrue TypeOf AdoParamsArray(1) Is ADODB.Parameter, "AdoParamsArray - wrong element type"
    Assert.AreEqual "FirstName", AdoParamsArray(1).Name, "Array - first param name mismatch"
    Assert.AreEqual adVarWChar, AdoParamsArray(1).Type, "Array - first param type mismatch"
    Assert.AreEqual "id", AdoParamsArray(5).Name, "Array - last param name mismatch"
    Assert.AreEqual adInteger, AdoParamsArray(5).Type, "Array - last param type mismatch"
    
    Assert.AreEqual 5, AdoParamsCol.Count, "Parameters collection - wrong count"
    Assert.AreEqual "FirstName", AdoParamsCol(0).Name, "Col - first param name mismatch"
    Assert.AreEqual adVarWChar, AdoParamsCol(0).Type, "Col - first param type mismatch"
    Assert.AreEqual "id", AdoParamsCol(4).Name, "Col - last param name mismatch"
    Assert.AreEqual adInteger, AdoParamsCol(4).Type, "Col - last param type mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("ADO Parameters")
Private Sub ztcRecordValuesToAdoParams_ValidatesUpdatedParams()
    On Error GoTo TestFail

Arrange:
    Dim FieldNames(1 To 5) As String
    Dim FieldTypes(1 To 5) As ADODB.DataTypeEnum
    FieldNames(1) = "id":        FieldTypes(1) = adInteger
    FieldNames(2) = "FirstName": FieldTypes(2) = adVarWChar
    FieldNames(3) = "LastName":  FieldTypes(3) = adVarWChar
    FieldNames(4) = "Age":       FieldTypes(4) = adInteger
    FieldNames(5) = "Gender":    FieldTypes(5) = adVarWChar
    
    Dim RecordValues() As Variant
    RecordValues = Array(4, "Edna", "Jennings", 26, "male")
    
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    Dim AdoParamsCol As ADODB.Parameters
    Set AdoParamsCol = AdoCommand.Parameters
    Dim AdoParamsArray As Variant
    AdoParamsArray = ADOlib.SetAdoParamsForRecordUpdate(FieldNames, FieldTypes, AdoParamsCol)
Act:
    ADOlib.RecordValuesToAdoParams RecordValues, AdoParamsArray
    ADOlib.RecordValuesToAdoParams RecordValues, AdoParamsCol
Assert:
    Assert.AreEqual "Edna", AdoParamsArray(1).Value, "Array - updated value mismatch"
    Assert.AreEqual "Jennings", AdoParamsArray(2).Value, "Array - updated value mismatch"
    Assert.AreEqual 26, AdoParamsArray(3).Value, "Array - updated value mismatch"
    Assert.AreEqual "male", AdoParamsArray(4).Value, "Array - updated value mismatch"
    Assert.AreEqual 4, AdoParamsArray(5).Value, "Array - updated value mismatch"

    Assert.AreEqual "Edna", AdoParamsCol(0).Value, "Col - updated value mismatch"
    Assert.AreEqual "Jennings", AdoParamsCol(1).Value, "Col - updated value mismatch"
    Assert.AreEqual 26, AdoParamsCol(2).Value, "Col - updated value mismatch"
    Assert.AreEqual "male", AdoParamsCol(3).Value, "Col - updated value mismatch"
    Assert.AreEqual 4, AdoParamsCol(4).Value, "Col - updated value mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub

