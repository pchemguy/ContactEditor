Attribute VB_Name = "DataTableSecureADODBTests"
'@Folder "Storage Library.Table.Backend"
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit
Option Private Module


#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If

Const TEST_TABLE As String = "Contacts"

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


Private Function zfxGetDataTableSecureADODB() As DataTableSecureADODB
    Dim StorageModel As DataTableModel
    Set StorageModel = New DataTableModel
    Dim ConnectionString As String
    ConnectionString = ADOlib.GetSQLiteConnectionString()("ADO")
    Dim TableName As String
    TableName = TEST_TABLE
    Set zfxGetDataTableSecureADODB = DataTableSecureADODB.Create(StorageModel, ConnectionString, TableName)
End Function


Private Function zfxGetDataTableModel() As DataTableModel
    Dim StorageModel As DataTableModel
    Set StorageModel = New DataTableModel
    Dim ConnectionString As String
    ConnectionString = ADOlib.GetSQLiteConnectionString()("ADO")
    Dim TableName As String
    TableName = TEST_TABLE
    
    Dim SMiDefault As DataTableSecureADODB
    Set SMiDefault = DataTableSecureADODB.Create(StorageModel, ConnectionString, TableName)
    Dim StorageManager As IDataTableStorage
    Set StorageManager = SMiDefault.SelfIDataTableStorage
    StorageManager.LoadDataIntoModel
    Set zfxGetDataTableModel = StorageModel
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("ADODB")
Private Sub ztcGetAdoCommand_ValidatesAdoCommandTable()
    On Error GoTo TestFail

Arrange:
    Dim SMiDefault As DataTableSecureADODB
    Set SMiDefault = zfxGetDataTableSecureADODB
    Dim SQLQuery As String
    SQLQuery = SQLlib.Create(TEST_TABLE).SelectAll
Act:
    Dim AdoCommand As ADODB.Command: Set AdoCommand = SMiDefault.AdoCommandInit(SQLQuery)
Assert:
    Assert.AreEqual SQLQuery, AdoCommand.CommandText, "SQL query mismatch"
    Assert.AreEqual ADODB.CommandTypeEnum.adCmdText, AdoCommand.CommandType, "Command type mismatch"
    Assert.IsNotNothing AdoCommand.ActiveConnection, "ActiveConnection is not set."
    Assert.AreEqual ADODB.CursorLocationEnum.adUseClient, AdoCommand.ActiveConnection.CursorLocation, "Cursor location mismatch"
    Assert.AreEqual ADODB.ObjectStateEnum.adStateOpen, AdoCommand.ActiveConnection.State, "Connection is not opened"
    Assert.AreEqual 0, AdoCommand.ActiveConnection.Errors.Count, "Connection errors occured: #" & AdoCommand.ActiveConnection.Errors.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADODB")
Private Sub ztcGetAdoCommand_ValidatesAdoCommandQuery()
    On Error GoTo TestFail

Arrange:
    Dim SMiDefault As DataTableSecureADODB
    Set SMiDefault = zfxGetDataTableSecureADODB
    Dim SQLQuery As String
    SQLQuery = SQLlib.Create(TEST_TABLE).SelectAll("COUNT(*)")
Act:
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = SMiDefault.AdoCommandInit(SQLQuery, ADODB.CursorLocationEnum.adUseServer)
Assert:
    Assert.AreEqual SQLQuery, AdoCommand.CommandText, "SQL query mismatch"
    Assert.AreEqual ADODB.CommandTypeEnum.adCmdText, AdoCommand.CommandType, "Command type mismatch"
    Assert.IsTrue AdoCommand.Prepared, "Expected prepared command"
    Assert.IsNotNothing AdoCommand.ActiveConnection, "ActiveConnection is not set."
    Assert.AreEqual ADODB.CursorLocationEnum.adUseServer, AdoCommand.ActiveConnection.CursorLocation, "Cursor location mismatch"
    Assert.AreEqual ADODB.ObjectStateEnum.adStateOpen, AdoCommand.ActiveConnection.State, "Connection is not opened"
    Assert.AreEqual 0, AdoCommand.ActiveConnection.Errors.Count, "Connection errors occured: #" & AdoCommand.ActiveConnection.Errors.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADODB")
Private Sub ztcCollectTableMetadata_ValidatesTableMetadata()
    On Error GoTo TestFail

Arrange:
    Dim SMiDefault As DataTableSecureADODB
    Set SMiDefault = zfxGetDataTableSecureADODB
Act:
    SMiDefault.AdoCommandInit SQLlib.Create(TEST_TABLE).SelectAll
Assert:
    Assert.AreEqual 8, SMiDefault.FieldMap.Count, "FiledMap count mismatch"
    Assert.AreEqual 1, SMiDefault.FieldMap("id"), "FiledMap 'id' index mismatch"
    Assert.AreEqual 1, LBound(SMiDefault.FieldNames, 1), "FieldNames base is not 1"
    Assert.AreEqual 8, UBound(SMiDefault.FieldNames, 1), "FieldNames count mismatch"
    Assert.AreEqual "id", SMiDefault.FieldNames(1), "FieldNames - 'id' mismatch"
    Assert.AreEqual 1, LBound(SMiDefault.FieldTypes, 1), "FieldTypes base is not 1"
    Assert.AreEqual 8, UBound(SMiDefault.FieldTypes, 1), "FieldTypes count mismatch"
    Assert.AreEqual ADODB.DataTypeEnum.adInteger, SMiDefault.FieldTypes(1), "FieldTypes - 'id' type mismatch"
    Assert.AreEqual ADODB.DataTypeEnum.adVarWChar, SMiDefault.FieldTypes(2), "FieldTypes - 'FirstName' type mismatch"
    Assert.AreEqual ADODB.DataTypeEnum.adInteger, SMiDefault.FieldTypes(4), "FieldTypes - 'FirstName' type mismatch"
            
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADODB")
Private Sub ztcAdoRecordset_ValidatesAdoRecordset()
    On Error GoTo TestFail

Arrange:
    Dim SMiDefault As DataTableSecureADODB
    Set SMiDefault = zfxGetDataTableSecureADODB
Act:
    SMiDefault.AdoCommandInit SQLlib.Create(TEST_TABLE).SelectAll
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = SMiDefault.GetAdoRecordset
Assert:
    Assert.IsNothing AdoRecordset.ActiveCommand, "Expected disconnected recordset"
    Assert.IsNothing AdoRecordset.ActiveConnection, "Expected disconnected recordset"
    Assert.AreEqual 10, AdoRecordset.CacheSize, "CacheSize mismatch"
    Assert.AreEqual 1000, AdoRecordset.RecordCount, "RecordCount mismatch"
    Assert.AreEqual 8, AdoRecordset.Fields.Count, "Fields.Count mismatch"
    Assert.AreEqual "LastName", AdoRecordset.Fields(2).Name, "Field name mismatch"
        
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADODB")
Private Sub ztcRecords_ValidatesRecordsAsIs()
    On Error GoTo TestFail

Arrange:
    Dim SMiDefault As DataTableSecureADODB
    Set SMiDefault = zfxGetDataTableSecureADODB
Act:
    Dim Records As Variant: Records = SMiDefault.Records
Assert:
    Assert.IsFalse IsEmpty(Records), "Records is empty"
    Assert.AreEqual 1, LBound(Records, 1), "Records base is not 1"
    Assert.AreEqual 1000, UBound(Records, 1), "Records mismatch"
    Assert.AreEqual 1, LBound(Records, 2), "Fields base is not 1"
    Assert.AreEqual 8, UBound(Records, 2), "Fields count mismatch"
    Assert.AreEqual vbLong, VarType(Records(1, 1)), "ID field type mismatch"
    Assert.AreEqual vbString, VarType(Records(1, 2)), "FirstName field type mismatch"
    Assert.AreEqual vbLong, VarType(Records(1, 4)), "Age field type mismatch"
    Assert.AreEqual "Edna.Jennings@neuf.fr", Records(4, 6), "Field value mismatch"
        
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADODB")
Private Sub ztcRecords_ValidatesRecordsIdAsText()
    On Error GoTo TestFail

Arrange:
    Dim SMiDefault As DataTableSecureADODB
    Set SMiDefault = zfxGetDataTableSecureADODB
Act:
    Dim Records As Variant
    Records = SMiDefault.Records(SQLlib.Create(TEST_TABLE).SelectIdAsText(SMiDefault.FieldNames))
Assert:
    Assert.IsFalse IsEmpty(Records), "Records is empty"
    Assert.AreEqual 1, LBound(Records, 1), "Records base is not 1"
    Assert.AreEqual 1000, UBound(Records, 1), "Records mismatch"
    Assert.AreEqual 1, LBound(Records, 2), "Fields base is not 1"
    Assert.AreEqual 8, UBound(Records, 2), "Fields count mismatch"
    Assert.AreEqual vbString, VarType(Records(1, 1)), "ID field type mismatch"
    Assert.AreEqual vbString, VarType(Records(1, 2)), "FirstName field type mismatch"
    Assert.AreEqual vbLong, VarType(Records(1, 4)), "Age field type mismatch"
    Assert.AreEqual "Edna.Jennings@neuf.fr", Records(4, 6), "Field value mismatch"
        
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ADODB")
Private Sub ztcRecords_ValidatesRecordsAllAsText()
    On Error GoTo TestFail

Arrange:
    Dim SMiDefault As DataTableSecureADODB
    Set SMiDefault = zfxGetDataTableSecureADODB
Act:
    Dim Records As Variant
    Records = SMiDefault.Records( _
        SQLlib.Create(TEST_TABLE).SelectAllAsText(SMiDefault.FieldNames, SMiDefault.FieldTypes))
Assert:
    Assert.IsFalse IsEmpty(Records), "Records is empty"
    Assert.AreEqual 1, LBound(Records, 1), "Records base is not 1"
    Assert.AreEqual 1000, UBound(Records, 1), "Records mismatch"
    Assert.AreEqual 1, LBound(Records, 2), "Fields base is not 1"
    Assert.AreEqual 8, UBound(Records, 2), "Fields count mismatch"
    Assert.AreEqual VBA.VbVarType.vbString, VarType(Records(1, 1)) And VBA.VbVarType.vbString, "ID field type mismatch"
    Assert.AreEqual VBA.VbVarType.vbString, VarType(Records(1, 2)) And VBA.VbVarType.vbString, "FirstName field type mismatch"
    Assert.AreEqual VBA.VbVarType.vbString, VarType(Records(1, 4)) And VBA.VbVarType.vbString, "Age field type mismatch"
    Assert.AreEqual "Edna.Jennings@neuf.fr", Records(4, 6), "Field value mismatch"
        
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTableModel")
Private Sub ztcModel_ValidatesLoadedData()
    On Error GoTo TestFail

Arrange:
Act:
    Dim StorageModel As DataTableModel
    Set StorageModel = zfxGetDataTableModel
Assert:
    With StorageModel
        Assert.IsNotNothing .DirtyRecords, "Dirty records dictionary is not set"
        Assert.AreEqual 0, .DirtyRecords.Count, "Dirty records count should be 0"
        Assert.IsFalse .IsDirty, "Model should not be dirty"

        Assert.IsNotNothing .FieldIndices, "FieldIndices dictionary is not set"
        Assert.AreEqual 8, .FieldIndices.Count, "FieldIndices - wrong field count"
        Assert.IsTrue .FieldIndices.Exists("Email"), "FieldIndices - missing field"
        Assert.AreEqual 6, .FieldIndices("Email"), "FieldIndices - index mismatch"

        Assert.IsTrue IsArray(.FieldNames), "FieldNames is not set"
        Assert.AreEqual 1, LBound(.FieldNames, 1), "FieldNames - wrong index base"
        Assert.AreEqual 8, UBound(.FieldNames, 1), "FieldNames - wrong field count"
        Assert.AreEqual "Email", .FieldNames(6), "FieldNames - item mismatch"

        Assert.IsNotNothing .IdIndices, "IdIndices dictionary is not set"
        Assert.AreEqual 1000, .IdIndices.Count, "IdIndices - wrong record count"
        Assert.AreEqual 90, .IdIndices("90"), "IdIndices - wrong record index"

        Assert.IsTrue IsArray(.Values), "Values is not set"
        Assert.AreEqual 1, LBound(.Values, 1), "Values - wrong record index base"
        Assert.AreEqual 1000, UBound(.Values, 1), "Values - wrong record count"
        Assert.AreEqual 1, LBound(.Values, 2), "Values - wrong field index base"
        Assert.AreEqual 8, UBound(.Values, 2), "Values - wrong field count"
        Assert.AreEqual "Edna.Jennings@neuf.fr", .Values(4, 6), "Values - field mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
