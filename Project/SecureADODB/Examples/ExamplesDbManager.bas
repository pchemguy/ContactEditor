Attribute VB_Name = "ExamplesDbManager"
'@Folder "SecureADODB.Examples"
'@IgnoreModule AssignmentNotUsed, EmptyModule, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "SecureADODB"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


Private Sub CSVSingleParameterQueryTableTest()
    Dim FileName As String
    FileName = LIB_NAME & ".csv"

    Dim TableName As String
    TableName = FileName
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", REL_PREFIX & FileName, vbNullString, LoggerTypeEnum.logPrivate)

    Debug.Print dbm.Connection.AdoConnection.Properties("Transaction DDL").Value
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10)
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45)
    
    rst.RecordsetToQT Buffer.Range("A1")
End Sub


'''' Throws "Unsupported backend" Error
Private Sub InvalidTypeTest()
    Dim FileName As String
    FileName = LIB_NAME & ".csv"

    Dim TableName As String
    TableName = FileName
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    '''' Throws "Unsupported backend" Error
    Set dbm = DbManager.CreateFileDb("Driver=", REL_PREFIX & FileName, vbNullString, LoggerTypeEnum.logPrivate)
End Sub


Private Sub CSVSingleParameterQueryScalarTest()
    Dim FileName As String
    FileName = LIB_NAME & ".csv"

    Dim TableName As String
    TableName = FileName
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea' ORDER BY id DESC"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", REL_PREFIX & FileName, vbNullString, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10)
    
    Dim Result As Variant
    Result = rst.OpenScalar(SQLQuery, 45)
    
    Debug.Print "===== " & CStr(Result) & " ====="
    
    rst.RecordsetToQT Buffer.Range("A1")
End Sub


Private Sub SQLiteSingleParameterQueryTableTest()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10)
    
    Debug.Print dbm.Connection.AdoConnection.Properties("Transaction DDL")
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45)
    
'''' Before .Open
''''   Result.LockType = adLockBatchOptimistic
'''' After .Open
''''   Result.MarshalOptions = adMarshalModifiedOnly
    
    rst.RecordsetToQT Buffer.Range("A1")
End Sub


Private Sub SQLiteMetaTest()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, LoggerTypeEnum.logPrivate)
        
    Dim FieldNames() As String
    Dim FieldTypes() As ADODB.DataTypeEnum
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = New Scripting.Dictionary
    FieldMap.CompareMode = TextCompare
    dbm.DbMeta.QueryTableADOXMeta TableName, FieldNames, FieldTypes, FieldMap
    
    Dim FieldCount As Long
    FieldCount = FieldMap.Count
    Dim FieldIndex As Long
    Dim FieldName As String
    Dim FieldType As String
    Dim FieldData() As String
    ReDim FieldData(1 To FieldCount)
    For FieldIndex = 1 To FieldCount
        FieldName = FieldNames(FieldIndex)
        FieldType = AdoTypeMappings.DataTypeEnumAsText(CStr(FieldTypes(FieldIndex)))
        FieldType = FieldType & String(12 - Len(FieldType), " ")
        FieldData(FieldIndex) = CStr(FieldIndex) & ". " & _
                                FieldName & String(12 - Len(FieldName), " ") & vbTab & "|" & vbTab & _
                                FieldType & "|" & vbTab & _
                                CStr(FieldMap(FieldName)) & " <= '" & FieldName & "'"
    Next FieldIndex
    
    Debug.Print Join(FieldData, vbNewLine)
End Sub


Private Sub CSVMetaTest()
    Dim FileName As String
    FileName = LIB_NAME & ".csv"

    Dim TableName As String
    TableName = FileName
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", REL_PREFIX & FileName, vbNullString, LoggerTypeEnum.logPrivate)
        
    Dim FieldNames() As String
    Dim FieldTypes() As ADODB.DataTypeEnum
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = New Scripting.Dictionary
    FieldMap.CompareMode = TextCompare
    dbm.DbMeta.QueryTableADOXMeta TableName, FieldNames, FieldTypes, FieldMap
    
    Dim FieldCount As Long
    FieldCount = FieldMap.Count
    Dim FieldIndex As Long
    Dim FieldName As String
    Dim FieldType As String
    Dim FieldData() As String
    ReDim FieldData(1 To FieldCount)
    For FieldIndex = 1 To FieldCount
        FieldName = FieldNames(FieldIndex)
        FieldType = AdoTypeMappings.DataTypeEnumAsText(CStr(FieldTypes(FieldIndex)))
        FieldType = FieldType & String(12 - Len(FieldType), " ")
        FieldData(FieldIndex) = CStr(FieldIndex) & ". " & _
                                FieldName & String(12 - Len(FieldName), " ") & vbTab & "|" & vbTab & _
                                FieldType & vbTab & "|" & vbTab & _
                                CStr(FieldMap(FieldName)) & " <= '" & FieldName & "'"
    Next FieldIndex
    
    Debug.Print Join(FieldData, vbNewLine)
End Sub


Private Sub SQLiteInsertTest()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people_insert"
    Dim SQLQuery As String
    SQLQuery = "INSERT INTO " & TableName & " (id, first_name, last_name, age, gender, email, country, domain)" & _
               "VALUES (" & GenerateSerialID & ", 'first_name', 'last_name', 32, 'male', 'first_name.last_name@domain.com', 'Country', 'domain.com')"
               
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, LoggerTypeEnum.logPrivate)
    
    Dim cmd As IDbCommand
    Set cmd = dbm.Command
    cmd.ExecuteNonQuery SQLQuery
    
    Dim conn As IDbConnection
    Set conn = dbm.Connection
    Dim RecordsAffected As Long
    RecordsAffected = conn.RecordsAffected
    Dim ExecuteStatus As ADODB.EventStatusEnum
    ExecuteStatus = conn.ExecuteStatus
End Sub


Private Sub SQLiteTwoParameterQueryTableTest()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = ?"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, LoggerTypeEnum.logPrivate)

    Dim Log As ILogger
    Set Log = dbm.LogController

    Dim conn As IDbConnection
    Set conn = dbm.Connection
    Dim connAdo As ADODB.Connection
    Set connAdo = conn.AdoConnection
    
    Dim cmd As IDbCommand
    Set cmd = dbm.Command
    Dim cmdAdo As ADODB.Command
    Set cmdAdo = cmd.AdoCommand(SQLQuery, 45, "South Korea")
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.OpenRecordset(SQLQuery, 45, "South Korea")
    
    rst.RecordsetToQT Buffer.Range("A1")
End Sub


Private Sub CSVTwoParameterQueryTableTest()
    Dim FileName As String
    FileName = LIB_NAME & ".csv"

    Dim TableName As String
    TableName = FileName
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = ?"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", REL_PREFIX & FileName, vbNullString, LoggerTypeEnum.logPrivate)

    Dim Log As ILogger
    Set Log = dbm.LogController

    Dim conn As IDbConnection
    Set conn = dbm.Connection
    Dim connAdo As ADODB.Connection
    Set connAdo = conn.AdoConnection
    
    Dim cmd As IDbCommand
    Set cmd = dbm.Command
    Dim cmdAdo As ADODB.Command
    Set cmdAdo = cmd.AdoCommand(SQLQuery, 45, "South Korea")
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10)
    Dim rstAdo As ADODB.Recordset
    
    Set rstAdo = rst.OpenRecordset(SQLQuery, 45, "South Korea")
End Sub


Private Sub SQLiteTwoParameterQueryTableUpdateRstTest()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = ?"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, LoggerTypeEnum.logPrivate)
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10, LockType:=adLockBatchOptimistic)

    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.OpenRecordset(SQLQuery, 45, "South Korea")
    
    Dim TargetRecordIndex As Long
    Dim NumRecords As Long
    
    With rstAdo
        TargetRecordIndex = 2
        '@Ignore ValueRequired: False positive
        NumRecords = TargetRecordIndex - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        .Fields(1) = .Fields(1) & "XXX"
        .Fields("last_name") = .Fields("last_name") & "YYY"
    
        TargetRecordIndex = 4
        '@Ignore ValueRequired: False positive
        NumRecords = TargetRecordIndex - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        .Fields(1) = .Fields(1) & "XXX"
        .Fields("last_name") = .Fields("last_name") & "YYY"
    End With
    
    Dim WSQueryTable As Excel.QueryTable
    Set WSQueryTable = rst.RecordsetToQT(Buffer.Range("A1"))
        
    rstAdo.MarshalOptions = adMarshalModifiedOnly
    Set rstAdo.ActiveConnection = dbm.Connection.AdoConnection
    rstAdo.UpdateBatch

    With rstAdo
        TargetRecordIndex = 2
        '@Ignore ValueRequired: False positive
        NumRecords = TargetRecordIndex - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        .Fields(1) = "Nicolas"
        .Fields("last_name") = "Parrott"
    
        TargetRecordIndex = 4
        '@Ignore ValueRequired: False positive
        NumRecords = TargetRecordIndex - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        .Fields(1) = "Malcolm"
        .Fields("last_name") = "Eakins"
    End With
    rstAdo.UpdateBatch
End Sub


Private Sub SQLiteTwoParameterQueryTableUpdateRstTransactionChangesTest()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = ?"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, LoggerTypeEnum.logPrivate)
    
    Dim SQLiteSQLTotalChanges As String
    SQLiteSQLTotalChanges = "SELECT total_changes()"
    
    Dim RstTotalChanges As IDbRecordset
    Set RstTotalChanges = dbm.Recordset()
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10, LockType:=adLockBatchOptimistic)

    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.OpenRecordset(SQLQuery, 45, "South Korea")
    
    Dim TargetRecordIndex As Long
    Dim NumRecords As Long
    
    With rstAdo
        TargetRecordIndex = 2
        '@Ignore ValueRequired: False positive
        NumRecords = TargetRecordIndex - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        .Fields(1) = .Fields(1) & "XXX"
        .Fields("last_name") = .Fields("last_name") & "YYY"
    
        TargetRecordIndex = 4
        '@Ignore ValueRequired: False positive
        NumRecords = TargetRecordIndex - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        .Fields(1) = .Fields(1) & "XXX"
        .Fields("last_name") = .Fields("last_name") & "YYY"
    End With
    
    Dim WSQueryTable As Excel.QueryTable
    Set WSQueryTable = rst.RecordsetToQT(Buffer.Range("A1"))
        
    rstAdo.MarshalOptions = adMarshalModifiedOnly
    Set rstAdo.ActiveConnection = dbm.Connection.AdoConnection
    
    Dim TotalChanges As Variant
    
    TotalChanges = RstTotalChanges.OpenScalar(SQLiteSQLTotalChanges)
    Debug.Print TotalChanges
    dbm.Begin
    rstAdo.UpdateBatch
    dbm.Commit
    TotalChanges = RstTotalChanges.OpenScalar(SQLiteSQLTotalChanges)
    Debug.Print TotalChanges

    With rstAdo
        TargetRecordIndex = 2
        '@Ignore ValueRequired: False positive
        NumRecords = TargetRecordIndex - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        .Fields(1) = "Nicolas"
        .Fields("last_name") = "Parrott"
    
        TargetRecordIndex = 4
        '@Ignore ValueRequired: False positive
        NumRecords = TargetRecordIndex - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        .Fields(1) = "Malcolm"
        .Fields("last_name") = "Eakins"
    End With
    
    dbm.Connection.ExpectedRecordsAffected = 2
    dbm.Begin
    rstAdo.UpdateBatch
    dbm.Commit
End Sub
