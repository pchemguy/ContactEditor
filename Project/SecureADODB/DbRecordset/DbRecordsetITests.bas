Attribute VB_Name = "DbRecordsetITests"
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

Private Const LIB_NAME As String = "SecureADODB"
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


Private Function zfxGetTwoParameterSelectSQL() As String
    zfxGetTwoParameterSelectSQL = "SELECT * FROM people WHERE age >= ? AND country = ?"
End Function


Private Function zfxGetDbRecordset() As IDbRecordset
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName)
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10, LockType:=adLockBatchOptimistic)

    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.OpenRecordset(zfxGetTwoParameterSelectSQL, 45, "South Korea")
    
    Set zfxGetDbRecordset = rst
End Function


Private Function zfxGetNewRecord2Data() As Scripting.Dictionary
    Dim RecordValues As Scripting.Dictionary
    Set RecordValues = New Scripting.Dictionary

    With RecordValues
        .CompareMode = TextCompare
        .Item("id") = 334
        .Item("first_name") = "Nicolas"
        .Item("last_name") = "Parrott"
        .Item("age") = 50
        .Item("gender") = "male"
        .Item("email") = "Nicolas.Parrott@rediffmail.com"
        .Item("country") = "South Korea"
        .Item("domain") = "comicbookmovie.com"
    End With

    Set zfxGetNewRecord2Data = RecordValues
End Function


''Private Function zfxGetNewRecord4Data() As Scripting.Dictionary
''    Dim RecordValues As Scripting.Dictionary
''    Set RecordValues = New Scripting.Dictionary
''
''    With RecordValues
''        .CompareMode = TextCompare
''        .Item("id") = 370
''        .Item("first_name") = "Malcolm"
''        .Item("last_name") = "Eakins"
''        .Item("age") = 61
''        .Item("gender") = "male"
''        .Item("email") = "Malcolm.Eakins@aim.com"
''        .Item("country") = "South Korea"
''        .Item("domain") = "formstack.com"
''    End With
''
''    Set zfxGetNewRecord4Data = RecordValues
''End Function
''
''
''Private Function zfxGetNewRecordBadData() As Scripting.Dictionary
''    Dim RecordValues As Scripting.Dictionary
''    Set RecordValues = New Scripting.Dictionary
''
''    With RecordValues
''        .Item("bad_field_name") = "dummy value"
''    End With
''
''    Set zfxGetNewRecordBadData = RecordValues
''End Function


'===================================================='
'================= TESTING FIXTURES ================='
'===================================================='


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("AdoRecordset")
Private Sub ztcOpenRecordset_ValidatesRecordset()
    On Error GoTo TestFail
    
Arrange:
Act:
    Dim rst As IDbRecordset
    Set rst = zfxGetDbRecordset()
Assert:
    Assert.AreEqual 11, rst.AdoRecordset.RecordCount, "Record count mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

    
'@TestMethod("UpdateRecord")
Private Sub ztcUpdateRecord_ThrowsIfValuesDictNotSet()
    On Error Resume Next
    Dim Recordset As IDbRecordset
    Set Recordset = zfxGetDbRecordset()
    Dim ValuesDict As Scripting.Dictionary
    Set ValuesDict = New Scripting.Dictionary
    '@Ignore ValueRequired: false positive
    Recordset.UpdateRecord Recordset.AdoRecordset.RecordCount + 1, ValuesDict
    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub


'@TestMethod("UpdateRecord")
Private Sub ztcUpdateRecord_ValidatesRecordUpdate()
    On Error GoTo TestFail
    
Arrange:
    Dim Recordset As IDbRecordset
    Set Recordset = zfxGetDbRecordset()
    Dim ValuesDict As Scripting.Dictionary
    Set ValuesDict = zfxGetNewRecord2Data
Act:
    Dim rst As IDbRecordset
    Set rst = zfxGetDbRecordset()
Assert:
    Assert.AreEqual 11, rst.AdoRecordset.RecordCount, "Record count mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
