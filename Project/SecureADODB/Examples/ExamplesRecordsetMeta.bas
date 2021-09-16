Attribute VB_Name = "ExamplesRecordsetMeta"
'@Folder "SecureADODB.Examples"
'@IgnoreModule AssignmentNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
'@IgnoreModule FunctionReturnValueDiscarded
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "SecureADODB"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


Private Sub Array2DToWSheetTest()
    Dim DummyArray As Variant
    DummyArray = Buffer.Range("A1:C5").Value
    DummyArray(1, 1) = "1,1"
    DummyArray(5, 3) = "5,3"
    Buffer.Range("A1:C5").Value = DummyArray
    Buffer.Range("D1").Resize( _
        UBound(DummyArray, 1) - LBound(DummyArray, 1) + 1, _
        UBound(DummyArray, 2) - LBound(DummyArray, 2) + 1 _
    ).Value = DummyArray
End Sub


Private Sub GetRecordsetMeta()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName)
    
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE id > 10 AND id <= ? AND gender = ?"
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10, LockType:=adLockBatchOptimistic)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.OpenRecordset(SQLQuery, 20, "male")
    
    Dim RecordsetMeta As DbRecordsetMeta
    Set RecordsetMeta = DbRecordsetMeta.Create(rstAdo)
    
    Dim RecordsetAttributes As Variant
    RecordsetAttributes = RecordsetMeta.GetRecordsetAttrbutes(Buffer.Range("E1"))
    Dim RecordsetProperties As Variant
    RecordsetProperties = RecordsetMeta.GetRecordsetProperties(Buffer.Range("A1"))
    Dim CursorOptions As Variant
    CursorOptions = RecordsetMeta.GetCursorOptions(Buffer.Range("I1"))
End Sub


Private Sub GetFieldsAttributes()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName)
    
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE id > 10 AND id <= ? AND gender = ?"
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10, LockType:=adLockBatchOptimistic)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.OpenRecordset(SQLQuery, 20, "male")
    
    Dim RecordsetMeta As DbRecordsetMeta
    Set RecordsetMeta = DbRecordsetMeta.Create(rstAdo)

    Dim FieldsAttributes As Variant
    FieldsAttributes = RecordsetMeta.GetFieldsAttributes(Buffer.Range("B1"))
End Sub


Private Sub GetFieldsProperties()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName)
    
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE id > 10 AND id <= ? AND gender = ?"
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10, LockType:=adLockBatchOptimistic)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.OpenRecordset(SQLQuery, 20, "male")
    
    Dim RecordsetMeta As DbRecordsetMeta
    Set RecordsetMeta = DbRecordsetMeta.Create(rstAdo)

    Dim FieldsProperties As Variant
    FieldsProperties = RecordsetMeta.GetFieldsProperties(Buffer.Range("A10"))
End Sub


'''' .OpenSchema support by the SQLiteODBC driver is very poor. PK info N/A.
'''' PK info is available via the KEYCOLUMN property of the field in
'''' a recordset at least for SQLite via SQLiteODBC.
'''' IMPORTANT: to get PK info, the recordset's LockType must be "updatable",
'''' that is at least adLockOptimistic. With default adLockReadOnly, PK info
'''' is not set.
Private Sub GetPK()
    Dim FileName As String
    FileName = REL_PREFIX & "SQLiteDBVBALibrary.db"

    Dim TableName As String
    TableName = "data_audit_log"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName)
    
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE 1 = 0 AND operation = ? AND timestamp > ?"
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10, LockType:=adLockOptimistic)
    
    Dim Result As Variant
    Result = rst.OpenScalar(SQLQuery, "LOG", 0)
    
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.AdoRecordset
    
    Dim FieldCount As Long
    FieldCount = rstAdo.Fields.Count
    Dim PKFlags() As Boolean
    ReDim PKFlags(1 To FieldCount)
    Dim FieldIndex As Long
    For FieldIndex = 1 To FieldCount
        PKFlags(FieldIndex) = rstAdo.Fields(FieldIndex - 1).Properties("KEYCOLUMN")
    Next FieldIndex
End Sub
