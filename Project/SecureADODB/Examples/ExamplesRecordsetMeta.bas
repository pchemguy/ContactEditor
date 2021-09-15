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


Private Sub GetRecordsetCoreAttributes()
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
    
    Dim RecordsetProperties As Variant
    RecordsetProperties = RecordsetMeta.GetRecordsetProperties(Buffer.Range("A16"))
    Dim RecordsetCoreAttributes As Variant
    RecordsetCoreAttributes = RecordsetMeta.GetRecordsetCoreAttrbutes(Buffer.Range("A1"))
End Sub
