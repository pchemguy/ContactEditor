Attribute VB_Name = "DbMetaDataITests"
'@Folder "SecureADODB.DbManager.DbMetaData"
'@TestModule
'@IgnoreModule UnhandledOnErrorResumeNext: Tests for expected errors do not reset error handling
'@IgnoreModule VariableNotUsed: Tests for expected errors may use dummy assignments
'@IgnoreModule AssignmentNotUsed: Tests for expected errors may use dummy assignments
'@IgnoreModule LineLabelNotUsed: Using standard test template
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit
Option Private Module

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetLibPrefix(ByVal LibName As String) As String
    Dim PathNameComponents As Variant
    PathNameComponents = Array( _
        ThisWorkbook.Path, _
        "Library", _
        LibName _
    )
    Dim PathName As String
    PathName = Join(PathNameComponents, Application.PathSeparator) & Application.PathSeparator
    zfxGetLibPrefix = PathName
End Function


Public Function zfxGetMeta() As DbMetaData
    Dim DbFileName As String
    DbFileName = zfxGetLibPrefix("SecureADODB") & "SecureADODBTest.db"
    Dim DbOK As Variant
    Dim Status As Long
    On Error Resume Next
    DbOK = VerifyOrGetDefaultPath(DbFileName, Empty)
    Status = Err.Number
    Select Case Status
        Case ErrNo.FileNotFoundErr
            Debug.Print vbNewLine & _
                "============================================================" & vbNewLine & _
                "<DbMetaTests.bas>: Test database <" & DbFileName & ">" & vbNewLine & _
                "is not found! Did you copy the <SecureADODB Fork> library" & vbNewLine & _
                "without the test db? Related tests should be reported as" & vbNewLine & _
                "inconclusive. You can copy original test database," & vbNewLine & _
                "adjust tests, or disable them." & vbNewLine & _
                "============================================================" & vbNewLine
        Case Is > 0
            Debug.Print vbNewLine & _
                "============================================================" & vbNewLine & _
                "<DbMetaTests.bas>: Unexpected error occured:" & vbNewLine & _
                "Err.Number       - " & CStr(Err.Number) & vbNewLine & _
                "Err.Source       - " & CStr(Err.Source) & vbNewLine & _
                "Err.Description  - " & CStr(Err.Description) & vbNewLine & _
                "============================================================" & vbNewLine
    End Select
    On Error GoTo 0
    
    If Status > 0 Then
        Set zfxGetMeta = Nothing
    Else
        Dim DbConnStr As DbConnectionString
        Set DbConnStr = DbConnectionString.CreateFileDb("sqlite", DbFileName)
        Set zfxGetMeta = DbMetaData.Create(DbConnStr)
    End If
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Factory Guard")
Private Sub ztcCreate_ThrowsGivenNullConnectionString()
    On Error Resume Next
    Dim sut As DbMetaData
    Set sut = DbMetaData.Create(Nothing)
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub


'@TestMethod("DbManager.Command")
Private Sub ztiQueryTableADOXMeta_VerifiesTableMeta()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As DbMetaData
    Set dbm = zfxGetMeta
    If dbm Is Nothing Then GoTo TestInconclusive
Act:
    Dim TableName As String
    TableName = "contacts"
    Dim FieldNames() As String
    Dim FieldTypes() As ADODB.DataTypeEnum
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = New Scripting.Dictionary
    FieldMap.CompareMode = TextCompare
    dbm.QueryTableADOXMeta TableName, FieldNames, FieldTypes, FieldMap

Assert:
    Assert.IsTrue IsArray(FieldNames), "FieldNames should be an array"
    Assert.AreEqual 1, LBound(FieldNames), "FieldNames lower index should be 1"
    Assert.AreEqual 8, UBound(FieldNames), "FieldNames upper index should be 8"
    Assert.AreEqual "id", FieldNames(1), "FieldNames(1) should be 'id'"
    Assert.AreEqual "FirstName", FieldNames(2), "FieldNames(2) should be 'FirstName'"
    
    Assert.IsTrue IsArray(FieldTypes), "FieldTypes should be an array"
    Assert.AreEqual 1, LBound(FieldTypes), "FieldTypes lower index should be 1"
    Assert.AreEqual 8, UBound(FieldTypes), "FieldTypes upper index should be 8"
    Assert.AreEqual adInteger, FieldTypes(1), "FieldTypes(1) should be <adInteger>"
    Assert.AreEqual adVarWChar, FieldTypes(2), "FieldTypes(2) should be <adVarWChar>"
    
    Assert.AreEqual 8, FieldMap.Count, "FieldMap size should be 8"
    Assert.AreEqual 1, FieldMap("id"), "FieldMap('id') should be 1"
    Assert.AreEqual 2, FieldMap("FirstName"), "FieldMap('FirstName') should be 2"
    
CleanExit:
    Exit Sub
TestInconclusive:
    Assert.Inconclusive "Test database is not available."
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
