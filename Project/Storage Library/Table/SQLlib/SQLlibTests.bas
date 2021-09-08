Attribute VB_Name = "SQLlibTests"
'@Folder "Storage Library.Table.SQLlib"
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
    Expected = "SELECT * FROM [" & SQL.TableName & "]"
Act:
    Dim Actual As String
    Actual = SQL.SelectAll
Assert:
    Assert.AreEqual Expected, Actual, "Wildcard query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectAll_ValidatesFieldsQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT [id], [FirstName], [LastName] FROM [" & SQL.TableName & "]"
Act:
    Dim Actual As String
    Actual = SQL.SelectAll(Array("id", "FirstName", "LastName"))
Assert:
    Assert.AreEqual Expected, Actual, "Wildcard query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectOne_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT * FROM [" & SQL.TableName & "] LIMIT 1"
Act:
    Dim Actual As String
    SQL.SetLimit 1
    Actual = SQL.SelectAll
    SQL.SetLimit
Assert:
    Assert.AreEqual Expected, Actual, "SelectOne query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcAsText_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "CAST([id] AS TEXT) AS [id]"
Act:
    Dim Actual As String
    Actual = SQL.AsText("id")
Assert:
    Assert.AreEqual Expected, Actual, "SQLAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectIdAsText_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT CAST([id] AS TEXT) AS [id], [FirstName], [LastName], [Age] FROM [" & SQL.TableName & "]"
                
Act:
    Dim Actual As String
    Actual = SQL.SelectIdAsText(Array("id", "FirstName", "LastName", "Age"))
Assert:
    Assert.AreEqual Expected, Actual, "SelectIdAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectAllAsText_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT CAST([id] AS TEXT) AS [id], [FirstName], [LastName], CAST([Age] AS TEXT) AS [Age], [Gender] FROM [" & SQL.TableName & "]"
Act:
    Dim Actual As String
    Actual = SQL.SelectAllAsText(Array("id", "FirstName", "LastName", "Age", "Gender"), _
                                 Array(adInteger, adVarWChar, adVarWChar, adInteger, adVarWChar))
Assert:
    Assert.AreEqual Expected, Actual, "SelectIdAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcUpdateSingleRecord_ValidatesQuery()
    On Error GoTo TestFail

Arrange:
    Dim SQL As SQLlib: Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "UPDATE [" & SQL.TableName & "] SET ([FirstName], [LastName], [Age], [Gender], [Email]) = (?, ?, ?, ?, ?) WHERE [id] = ?"
Act:
    Dim Actual As String
    Actual = SQL.UpdateSingleRecord(Array("id", "FirstName", "LastName", "Age", "Gender", "Email"))
Assert:
    Assert.AreEqual Expected, Actual, "UpdateSingleRecord query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
