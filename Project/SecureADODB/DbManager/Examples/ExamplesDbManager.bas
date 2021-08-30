Attribute VB_Name = "ExamplesDbManager"
'@Folder "SecureADODB.DbManager.Examples"
'@IgnoreModule AssignmentNotUsed, EmptyModule, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
Option Explicit


Private Sub CSVSingleParameterQueryTableTest()
    Dim FileName As String
    FileName = ThisWorkbook.VBProject.Name & ".csv"

    Dim TableName As String
    TableName = FileName
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", FileName, vbNullString, False, LoggerTypeEnum.logPrivate)

    Debug.Print dbm.Connection.AdoConnection.Properties("Transaction DDL").Value
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45)
    
    rst.RecordsetToQT Buffer.Range("A1")
End Sub


'''' Throws "Unsupported backend" Error
Private Sub InvalidTypeTest()
    Dim FileName As String
    FileName = ThisWorkbook.VBProject.Name & ".csv"

    Dim TableName As String
    TableName = FileName
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    '''' Throws "Unsupported backend" Error
    Set dbm = DbManager.CreateFileDb("Driver=", FileName, vbNullString, True, LoggerTypeEnum.logPrivate)
End Sub


Private Sub CSVSingleParameterQueryScalarTest()
    Dim FileName As String
    FileName = ThisWorkbook.VBProject.Name & ".csv"

    Dim TableName As String
    TableName = FileName
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea' ORDER BY id DESC"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", FileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=True, Disconnected:=True, CacheSize:=10)
    
    Dim Result As Variant
    Result = rst.OpenScalar(SQLQuery, 45)
    
    Debug.Print "===== " & CStr(Result) & " ====="
    
    rst.RecordsetToQT Buffer.Range("A1")
End Sub


Private Sub SQLiteSingleParameterQueryTableTest()
    Dim FileName As String
    FileName = ThisWorkbook.VBProject.Name & ".db"

    Dim TableName As String
    TableName = "people"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    
    Debug.Print dbm.Connection.AdoConnection.Properties("Transaction DDL")
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45)
    
    rst.RecordsetToQT Buffer.Range("A1")
End Sub


Private Sub SQLiteMetaTest()
    Dim FileName As String
    FileName = ThisWorkbook.VBProject.Name & ".db"

    Dim TableName As String
    TableName = "people"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, True, LoggerTypeEnum.logPrivate)
        
    Dim FieldNames() As String
    Dim FieldTypes() As ADODB.DataTypeEnum
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = New Scripting.Dictionary
    FieldMap.CompareMode = TextCompare
    dbm.DbMeta.QueryTableADOXMeta TableName, FieldNames, FieldTypes, FieldMap
    
    Dim ADODBTypeMapping As Scripting.Dictionary
    Set ADODBTypeMapping = New Scripting.Dictionary
    ADODBTypeMapping.CompareMode = TextCompare
    With ADODBTypeMapping
        .Add CStr(adBoolean), "Boolean   /  adBoolean"
        .Add CStr(adCurrency), "Currency  /  adCurrency"
        .Add CStr(adDate), "Date      /  adDate"
        .Add CStr(adDouble), "Double    /  adDouble"
        .Add CStr(adInteger), "Long      /  adInteger"
        .Add CStr(adSingle), "Single    /  adSingle"
        .Add CStr(adVarWChar), "String    /  adVarWChar"
        .Add CStr(adVarChar), "String    /  adVarChar"
    End With
    
    Dim FieldCount As Long
    FieldCount = FieldMap.Count
    Dim FieldIndex As Long
    Dim FieldName As String
    Dim FieldType As String
    Dim FieldData() As String
    ReDim FieldData(1 To FieldCount)
    For FieldIndex = 1 To FieldCount
        FieldName = FieldNames(FieldIndex)
        FieldType = ADODBTypeMapping(CStr(FieldTypes(FieldIndex)))
        FieldType = FieldType & String(25 - Len(FieldType), " ")
        FieldData(FieldIndex) = CStr(FieldIndex) & ". " & _
                                FieldName & String(12 - Len(FieldName), " ") & vbTab & "|" & vbTab & _
                                FieldType & vbTab & "|" & vbTab & _
                                CStr(FieldMap(FieldName)) & " <= '" & FieldName & "'"
    Next FieldIndex
    
    Debug.Print Join(FieldData, vbNewLine)
End Sub


Private Sub CSVMetaTest()
    Dim FileName As String
    FileName = ThisWorkbook.VBProject.Name & ".csv"

    Dim TableName As String
    TableName = FileName
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", FileName, vbNullString, True, LoggerTypeEnum.logPrivate)
        
    Dim FieldNames() As String
    Dim FieldTypes() As ADODB.DataTypeEnum
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = New Scripting.Dictionary
    FieldMap.CompareMode = TextCompare
    dbm.DbMeta.QueryTableADOXMeta TableName, FieldNames, FieldTypes, FieldMap
    
    Dim ADODBTypeMapping As Scripting.Dictionary
    Set ADODBTypeMapping = New Scripting.Dictionary
    ADODBTypeMapping.CompareMode = TextCompare
    With ADODBTypeMapping
        .Add CStr(adBoolean), "Boolean   /  adBoolean"
        .Add CStr(adCurrency), "Currency  /  adCurrency"
        .Add CStr(adDate), "Date      /  adDate"
        .Add CStr(adDouble), "Double    /  adDouble"
        .Add CStr(adInteger), "Long      /  adInteger"
        .Add CStr(adSingle), "Single    /  adSingle"
        .Add CStr(adVarWChar), "String    /  adVarWChar"
        .Add CStr(adVarChar), "String    /  adVarChar"
    End With
    
    Dim FieldCount As Long
    FieldCount = FieldMap.Count
    Dim FieldIndex As Long
    Dim FieldName As String
    Dim FieldType As String
    Dim FieldData() As String
    ReDim FieldData(1 To FieldCount)
    For FieldIndex = 1 To FieldCount
        FieldName = FieldNames(FieldIndex)
        FieldType = ADODBTypeMapping(CStr(FieldTypes(FieldIndex)))
        FieldType = FieldType & String(25 - Len(FieldType), " ")
        FieldData(FieldIndex) = CStr(FieldIndex) & ". " & _
                                FieldName & String(12 - Len(FieldName), " ") & vbTab & "|" & vbTab & _
                                FieldType & vbTab & "|" & vbTab & _
                                CStr(FieldMap(FieldName)) & " <= '" & FieldName & "'"
    Next FieldIndex
    
    Debug.Print Join(FieldData, vbNewLine)
End Sub


Private Sub SQLiteInsertTest()
    Dim FileName As String
    FileName = ThisWorkbook.VBProject.Name & ".db"

    Dim TableName As String
    TableName = "people_insert"
    Dim SQLQuery As String
    SQLQuery = "INSERT INTO " & TableName & " (id, first_name, last_name, age, gender, email, country, domain)" & _
               "VALUES (" & GenerateSerialID & ", 'first_name', 'last_name', 32, 'male', 'first_name.last_name@domain.com', 'Country', 'domain.com')"
               
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, True, LoggerTypeEnum.logPrivate)
    
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
    FileName = ThisWorkbook.VBProject.Name & ".db"

    Dim TableName As String
    TableName = "people"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = ?"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, True, LoggerTypeEnum.logPrivate)

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
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.AdoRecordset(SQLQuery, 45, "South Korea")
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45, "South Korea")

    rst.RecordsetToQT Buffer.Range("A1")
End Sub


'''' This routine should raise 'Type is invalid' error.
'''' For parametrized queries, SecureADODB uses VBA(String)=>ADODB(adVarWChar)
'''' correspondence. However, the CSV backend and driver do not support
'''' ADODB(adVarWChar). Instead, VBA(String)=>ADODB(adVarChar) should be used
'''' (remove 'W' from 'adVarWChar').
Private Sub CSVTwoParameterQueryTableTest()
    Dim FileName As String
    FileName = ThisWorkbook.VBProject.Name & ".csv"

    Dim TableName As String
    TableName = FileName
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = ?"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("csv", FileName, vbNullString, True, LoggerTypeEnum.logPrivate)

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
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.AdoRecordset(SQLQuery, 45, "South Korea")
    
    Dim Result As ADODB.Recordset
    '''' Should fail with 'Type is invalid' error
    Set Result = rst.OpenRecordset(SQLQuery, 45, "South Korea")
End Sub
