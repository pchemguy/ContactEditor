Attribute VB_Name = "DbConnectionStringTests"
'@Folder "SecureADODB.DbManager.DbConnectionString"
'@TestModule
'@IgnoreModule
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
'==================== TEST CASES ===================='
'===================================================='

'@TestMethod("ConnectionString")
Private Sub ztcConnectionString_ValidatesDefaultSQLiteString()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & _
                VerifyOrGetDefaultPath(vbNullString, Array("db", "sqlite")) & _
               ";SyncPragma=NORMAL;FKSupport=True;"
Act:
    Dim ActualADO As String
    ActualADO = DbConnectionString.CreateFileDb("sqlite").ConnectionString
    Dim ActualQT As String
    ActualQT = DbConnectionString.CreateFileDb("sqlite").QTConnectionString
Assert:
    Assert.AreEqual Expected, ActualADO, "Default SQLite ADO ConnectionString mismatch"
    Assert.AreEqual "OLEDB;" & Expected, ActualQT, "Default SQLite QT ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConnectionString")
Private Sub ztcConnectionString_ValidatesDefaultSQLiteStringWithBlankDriver()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & _
                VerifyOrGetDefaultPath(vbNullString, Array("db", "sqlite")) & _
               ";SyncPragma=NORMAL;FKSupport=True;"
Act:
    Dim ActualADO As String
    ActualADO = DbConnectionString.CreateFileDb("sqlite", "").ConnectionString
    Dim ActualQT As String
    ActualQT = DbConnectionString.CreateFileDb("sqlite", "").QTConnectionString
Assert:
    Assert.AreEqual Expected, ActualADO, "Default SQLite ADO ConnectionString with blank driver mismatch"
    Assert.AreEqual "OLEDB;" & Expected, ActualQT, "Default SQLite QT ConnectionString with blank driver mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConnectionString")
Private Sub ztcConnectionString_ValidatesDefaultCSVString()
    On Error GoTo TestFail

Arrange:
    Dim Driver As String
    #If Win64 Then
        Driver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If
    Dim CSVPath As String
    CSVPath = VerifyOrGetDefaultPath(vbNullString, Array("xsv", "csv"))
    CSVPath = Left$(CSVPath, Len(CSVPath) - Len(ThisWorkbook.VBProject.Name) - 5)
    Dim Expected As String
    Expected = "Driver=" + Driver + ";" + "DefaultDir=" + CSVPath + ";"
Act:
    Dim ActualADO As String
    ActualADO = DbConnectionString.CreateFileDb("csv").ConnectionString
    Dim ActualQT As String
    ActualQT = DbConnectionString.CreateFileDb("csv").QTConnectionString
Assert:
    Assert.AreEqual Expected, ActualADO, "Default CSV ADO ConnectionString mismatch"
    Assert.AreEqual "OLEDB;" & Expected, ActualQT, "Default CSV QT ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConnectionString")
Private Sub ztcConnectionString_ValidatesCSVStringOnValidName()
    On Error GoTo TestFail

Arrange:
    Dim Driver As String
    #If Win64 Then
        Driver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If
    Dim CSVPath As String
    CSVPath = VerifyOrGetDefaultPath("README.md", Array("xsv", "csv"))
    CSVPath = Left$(CSVPath, Len(CSVPath) - Len("README.md") - 1)
    Dim Expected As String
    Expected = "Driver=" + Driver + ";" + "DefaultDir=" + CSVPath + ";"
Act:
    Dim ActualADO As String
    ActualADO = DbConnectionString.CreateFileDb("csv", "README.md").ConnectionString
    Dim ActualQT As String
    ActualQT = DbConnectionString.CreateFileDb("csv", "README.md").QTConnectionString
Assert:
    Assert.AreEqual Expected, ActualADO, "CSV ADO ConnectionString (valid name) mismatch"
    Assert.AreEqual "OLEDB;" & Expected, ActualQT, "CSV QT ConnectionString (valid name) mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConnectionString")
Private Sub ztcConnectionString_ThrowsForXLSBackend()
    On Error Resume Next
    Dim ConnectionString As String
    ConnectionString = DbConnectionString.CreateFileDb("xls").ConnectionString
    AssertExpectedError Assert, ErrNo.NotImplementedErr
End Sub


