Attribute VB_Name = "SQLlibTests"
'@Folder "ContactEditor.Storage.Table.SQL"
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
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetSQL() As SQLlib
    Dim TableName As String: TableName = "people"
    Set zfxGetSQL = SQLlib.Create(TableName)
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("SQL")
Private Sub ztcSelectAll_ValidatesWildcardQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT * FROM """ & SQL.TableName & """"
Act:
    Dim Actual As String
    Actual = SQL.SelectAll
Assert:
    Assert.AreEqual Expected, Actual, "Wildcard query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectAll_ValidatesFieldsQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT id, FirstName, LastName FROM """ & SQL.TableName & """"
Act:
    Dim Actual As String
    Actual = SQL.SelectAll(Array("id", "FirstName", "LastName"))
Assert:
    Assert.AreEqual Expected, Actual, "Wildcard query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectOne_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT * FROM """ & SQL.TableName & """ LIMIT 1"
Act:
    Dim Actual As String
    Actual = SQL.SelectOne
Assert:
    Assert.AreEqual Expected, Actual, "SelectOne query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("SQL")
Private Sub ztcAsText_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "CAST(id AS TEXT) AS id"
Act:
    Dim Actual As String
    Actual = SQL.AsText("id")
Assert:
    Assert.AreEqual Expected, Actual, "SQLAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectIdAsText_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT CAST(id AS TEXT) AS id, FirstName, LastName, Age FROM """ & SQL.TableName & """"
Act:
    Dim Actual As String
    Actual = SQL.SelectIdAsText(Array("id", "FirstName", "LastName", "Age"))
Assert:
    Assert.AreEqual Expected, Actual, "SelectIdAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectAllAsText_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT CAST(id AS TEXT) AS id, FirstName, LastName, CAST(Age AS TEXT) AS Age, Gender FROM """ & SQL.TableName & """"
Act:
    Dim Actual As String
    Actual = SQL.SelectAllAsText(Array("id", "FirstName", "LastName", "Age", "Gender"), _
                                 Array(adInteger, adVarWChar, adVarWChar, adInteger, adVarWChar))
Assert:
    Assert.AreEqual Expected, Actual, "SelectIdAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


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
    Actual = SQLlib.GetSQLiteConnectionString(vbNullString)("ADO")
Assert:
    Assert.AreEqual Expected, Actual, "CheckPath failed with default path"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


