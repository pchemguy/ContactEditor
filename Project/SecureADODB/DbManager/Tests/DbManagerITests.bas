Attribute VB_Name = "DbManagerITests"
'@Folder "SecureADODB.DbManager.Tests"
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


Private Function zfxGetDbManager( _
            Optional ByVal DbType As String = "sqlite", _
            Optional ByVal BaseName As String = "SecureADODB") As IDbManager
    Dim FileName As String
    FileName = BaseName & "." & IIf(DbType = "csv", "csv", "db")
    Dim LibName As String
    LibName = "SecureADODB"
    Dim PathName As String
    PathName = zfxGetLibPrefix(LibName) & FileName
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb(DbType, PathName, vbNullString, LoggerTypeEnum.logDisabled)
    Set zfxGetDbManager = dbm
End Function


Private Function zfxGetConnectionString( _
            Optional ByVal DbType As String = "sqlite", _
            Optional ByVal BaseName As String = "SecureADODB") As String
    zfxGetConnectionString = zfxGetDbManager(DbType, BaseName).DbConnStr.ConnectionString
End Function


Private Function zfxGetSQLSelect0P(TableName As String) As String
    zfxGetSQLSelect0P = "SELECT * FROM " & TableName & " WHERE age >= 45 AND country = 'South Korea' ORDER BY id DESC"
End Function


Private Function zfxGetSQLSelect1P(TableName As String) As String
    zfxGetSQLSelect1P = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea' ORDER BY id DESC"
End Function


Private Function zfxGetSQLSelect2P(TableName As String) As String
    zfxGetSQLSelect2P = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = ? ORDER BY id DESC"
End Function


Private Function zfxGetSQLInsert0P(TableName As String) As String
    zfxGetSQLInsert0P = _
        "INSERT INTO " & TableName & " (id, first_name, last_name, age, gender, email, country, domain) " & _
        "VALUES " & _
            "(" & CStr(GenerateSerialID) & ", 'first_name1', 'last_name1', 32, 'male', 'first_name1.last_name1@domain.com', 'Country', 'domain.com'), " & _
            "(" & CStr(GenerateSerialID + 1) & ", 'first_name2', 'last_name2', 32, 'male', 'first_name2.last_name2@domain.com', 'Country', 'domain.com')"
End Function


Private Function zfxGetCSVTableName() As String
    zfxGetCSVTableName = "SecureADODB.csv"
End Function


Private Function zfxGetSQLiteTableName() As String
    zfxGetSQLiteTableName = "people"
End Function


Private Function zfxGetSQLiteTableNameInsert() As String
    zfxGetSQLiteTableNameInsert = "people_insert"
End Function


Private Function zfxGetParameterOne() As Variant
    zfxGetParameterOne = 45
End Function


Private Function zfxGetParameterTwo() As Variant
    zfxGetParameterTwo = "South Korea"
End Function


'===================================================='
'================= TESTING FIXTURES ================='
'===================================================='


'===================================================='
'================ TEST MOCK DATABASE ================'
'===================================================='


'@TestMethod("DbManager.Command")
Private Sub ztiDbManagerCommand_VerifiesAdoCommand()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager
    Set dbm = zfxGetDbManager("sqlite", "SecureADODB")
    Dim SQLSelect2P As String
    SQLSelect2P = zfxGetSQLSelect2P(zfxGetSQLiteTableName)
Act:
    Dim cmdAdo As ADODB.Command
    Set cmdAdo = dbm.Command.AdoCommand(SQLSelect2P, zfxGetParameterOne, zfxGetParameterTwo)
Assert:
    Assert.IsNotNothing cmdAdo.ActiveConnection, "ActiveConnection of the Command object is not set."
    Assert.AreEqual ADODB.ObjectStateEnum.adStateOpen, cmdAdo.ActiveConnection.State, "ActiveConnection of the Command object is not open."
    Assert.IsTrue cmdAdo.Prepared, "Prepared property of the Command object not set."
    Assert.AreEqual 2, cmdAdo.Parameters.Count, "Command should have two parameters set."
    Assert.AreEqual ADODB.DataTypeEnum.adInteger, cmdAdo.Parameters.Item(0).Type, "Param #1 type should be adInteger."
    Assert.AreEqual 45, cmdAdo.Parameters.Item(0).Value, "Param #1 value should be 45."
    Assert.AreEqual ADODB.DataTypeEnum.adVarWChar, cmdAdo.Parameters.Item(1).Type, "Param #2 type should be adVarWChar."
    Assert.AreEqual "South Korea", cmdAdo.Parameters.Item(1).Value, "Param #2 value should be South Korea."
    Assert.AreNotEqual vbNullString, cmdAdo.CommandText
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbManager.Recordset")
Private Sub ztiDbManagerRecordset_VerifiesAdoRecordsetDefaultDisconnectedArray()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager
    Set dbm = zfxGetDbManager("sqlite", "SecureADODB")
    Dim SQLSelect2P As String
    SQLSelect2P = zfxGetSQLSelect2P(zfxGetSQLiteTableName)
Act:
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = dbm.Recordset.OpenRecordset(SQLSelect2P, zfxGetParameterOne, zfxGetParameterTwo)
Assert:
    Assert.IsNothing rstAdo.ActiveConnection, "ActiveConnection of the Recordset object should be nothing."
    Assert.IsNothing rstAdo.ActiveCommand, "ActiveCommand of the Recordset object should not be set."
    Assert.IsFalse IsFalsy(rstAdo.Source), "The Source property of the Recordset object is not set."
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenStatic, rstAdo.CursorType, "The CursorType of the Recordset object should be adOpenStatic."
    Assert.AreEqual ADODB.CursorLocationEnum.adUseClient, rstAdo.CursorLocation, "The CursorLocation of the Recordset object should be adUseClient."
    Assert.AreNotEqual 1, rstAdo.MaxRecords, "The MaxRecords of the Recordset object should not be set to 1 for a regular Recordset."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbManager.Recordset")
Private Sub ztiDbManagerRecordset_VerifiesAdoRecordsetScalar()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager
    Set dbm = zfxGetDbManager("sqlite", "SecureADODB")
    Dim SQLSelect2P As String
    SQLSelect2P = zfxGetSQLSelect2P(zfxGetSQLiteTableName)
Act:
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(CacheSize:=15)
    Dim Result As Variant
    Result = rst.OpenScalar(SQLSelect2P, zfxGetParameterOne, zfxGetParameterTwo)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.AdoRecordset
Assert:
    Assert.AreEqual 1, rstAdo.RecordCount, "The RecordCount of the Recordset object should be 1 for a scalar query."
    Assert.AreEqual 1, rstAdo.MaxRecords, "The MaxRecords of the Recordset object should be set to 1 for a scalar query."
    Assert.AreEqual 15, rstAdo.CacheSize, "The CacheSize of the Recordset object should be set to 15."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbManager.Recordset")
Private Sub ztiDbManagerRecordset_VerifiesAdoRecordsetOnlineArray()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager
    Set dbm = zfxGetDbManager("sqlite", "SecureADODB")
    Dim SQLSelect2P As String
    SQLSelect2P = zfxGetSQLSelect2P(zfxGetSQLiteTableName)
Act:
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = dbm.Recordset(Disconnected:=False).OpenRecordset(SQLSelect2P, zfxGetParameterOne, zfxGetParameterTwo)
Assert:
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenForwardOnly, rstAdo.CursorType, "The CursorType of the Recordset object should be adOpenForwardOnly."
    Assert.AreEqual ADODB.CursorLocationEnum.adUseServer, rstAdo.CursorLocation, "The CursorLocation of the Recordset object should be adUseServer."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbManager.Recordset.Query")
Private Sub ztiDbManagerOpenRecordset_VerifiesAdoRecordsetDisconnectedArraySQLite()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager
    Set dbm = zfxGetDbManager("sqlite", "SecureADODB")
    Dim SQLSelect2P As String
    SQLSelect2P = zfxGetSQLSelect2P(zfxGetSQLiteTableName)
Act:
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = dbm.Recordset.OpenRecordset(SQLSelect2P, zfxGetParameterOne, zfxGetParameterTwo)
Assert:
    Assert.AreEqual 11, rstAdo.RecordCount, "Recordset SQLite SELECT query RecordCount mismatch."
    Assert.AreEqual 2, rstAdo.PageCount, "Recordset SQLite SELECT query PageCount mismatch."
    Assert.AreEqual 8, rstAdo.Fields.Count, "Recordset SQLite SELECT query did not return expected number of fields."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbManager.Recordset.Query")
Private Sub ztiDbManagerOpenRecordset_VerifiesAdoRecordsetDisconnectedArrayCSV()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager
    Set dbm = zfxGetDbManager("csv", "SecureADODB")
    Dim SQLSelect1P As String
    SQLSelect1P = zfxGetSQLSelect1P(zfxGetCSVTableName)
Act:
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = dbm.Recordset.OpenRecordset(SQLSelect1P, zfxGetParameterOne)
Assert:
    Assert.AreEqual 11, rstAdo.RecordCount, "Recordset CSV SELECT query RecordCount mismatch."
    Assert.AreEqual 2, rstAdo.PageCount, "Recordset CSV SELECT query PageCount mismatch."
    Assert.AreEqual 8, rstAdo.Fields.Count, "Recordset CSV SELECT query did not return expected number of fields."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbManager.Recordset.Query")
Private Sub ztiDbManagerFactoryGuard_ThrowsIfRequestedTransactionNotSupported()
    On Error Resume Next
    Dim dbm As IDbManager
    Set dbm = zfxGetDbManager("csv", "SecureADODB")
    dbm.Begin
    Guard.AssertExpectedError Assert, ErrNo.AdoInvalidTransactionErr
End Sub


'@TestMethod("DbManager.Recordset.Query")
Private Sub ztiDbManagerOpenRecordset_VerifiesAdoRecordsetOnlineArraySQLite()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager
    Set dbm = zfxGetDbManager("sqlite", "SecureADODB")
    Dim SQLSelect2P As String
    SQLSelect2P = zfxGetSQLSelect2P(zfxGetSQLiteTableName)
Act:
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = dbm.Recordset(Disconnected:=False).OpenRecordset(SQLSelect2P, zfxGetParameterOne, zfxGetParameterTwo)
    Dim Result As Variant
    Result = rstAdo.GetRows
Assert:
    Assert.AreEqual -1, rstAdo.RecordCount, "Recordset SQLite SELECT query RecordCount mismatch."
    Assert.AreEqual -1, rstAdo.PageCount, "Recordset SQLite SELECT query PageCount mismatch."
    Assert.AreEqual ADODB.PositionEnum.adPosEOF, rstAdo.AbsolutePosition, "Recordset SQLite SELECT - AbsolutePosition mismatch."
    Assert.IsTrue IsArray(Result), "GetRows on recordset SQLite SELECT query did not return an array."
    Assert.AreEqual 7, UBound(Result, 1), "Recordset SQLite SELECT query did not return expected number of fields."
    Assert.AreEqual 10, UBound(Result, 2), "Recordset SQLite SELECT query did not return expected number of records."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbManager.Recordset.Query")
Private Sub ztiDbManagerOpenRecordset_VerifiesAdoRecordsetScalarCSV()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", zfxGetLibPrefix("SecureADODB") & "SecureADODB.csv")
    Dim SQLSelect As String
    SQLSelect = zfxGetSQLSelect0P(zfxGetCSVTableName)
Act:
    Dim Result As Variant
    Result = dbm.Recordset.OpenScalar(SQLSelect)
Assert:
    Assert.AreEqual 906, Result, "Scalar CSV SELECT query result mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbManager.Command.Query")
Private Sub ztiDbManagerExecuteNonQuery_VerifiesInsertSQLite()
    On Error GoTo TestFail
    
Arrange:
    '''' True (default value) parameter is next line activates transactions.
    '''' Transaction is activated in the DbManager constructor and, if not committed,
    '''' is rolledback in its destructor. Execution status indicates the result of an
    '''' individual executed command regardless of whether an active transaction is
    '''' present and, if present, regardless of whether it is later committed or rolledback.
    '''' Set to false below to disable transactions and activate the autocommit mode to see
    '''' the result of the test insert in the database.
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", zfxGetLibPrefix("SecureADODB") & "SecureADODB.db")
    Dim conn As IDbConnection
    Set conn = dbm.Connection
    Dim SQLInsert0P As String
    SQLInsert0P = zfxGetSQLInsert0P(zfxGetSQLiteTableNameInsert)
Act:
    dbm.Command.ExecuteNonQuery SQLInsert0P
    Dim RecordsAffected As Long
    RecordsAffected = conn.RecordsAffected
    Dim ExecuteStatus As ADODB.EventStatusEnum: ExecuteStatus = conn.ExecuteStatus
Assert:
    Assert.AreEqual ADODB.EventStatusEnum.adStatusOK, ExecuteStatus, "Execution status mismatch."
    Assert.AreEqual 2, RecordsAffected, "Execution status mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
