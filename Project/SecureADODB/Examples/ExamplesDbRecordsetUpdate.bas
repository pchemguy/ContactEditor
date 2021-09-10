Attribute VB_Name = "ExamplesDbRecordsetUpdate"
'@Folder "SecureADODB.Examples"
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "SecureADODB"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


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

