Attribute VB_Name = "ExamplesOpenSchema"
'@Folder "SecureADODB.Examples"
'@IgnoreModule
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "SecureADODB"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


Private Sub SQLiteConnectionOpenSchemaTableListTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    Dim FileName As String
    FileName = "SQLiteDBVBALibrary.db"
    
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & FileName
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoConnection As ADODB.Connection
    Set AdoConnection = New ADODB.Connection
    With AdoConnection
        .ConnectionString = adoConnStr
        .CursorLocation = adUseClient
        .Open
        Set AdoRecordset = .OpenSchema(adSchemaTables)
    End With
    
    DbRecordset.RecordsetToQT Buffer.Range("A1"), AdoRecordset
End Sub


Private Sub SQLiteConnectionOpenSchemaColumnsTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    Dim FileName As String
    FileName = "SQLiteDBVBALibrary.db"
    
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & FileName
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    Dim RstFilter As String
    RstFilter = "[TABLE_CATALOG] = 'main' AND [TABLE_NAME] = 'companies'"
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoConnection As ADODB.Connection
    Set AdoConnection = New ADODB.Connection
    With AdoConnection
        .ConnectionString = adoConnStr
        .CursorLocation = adUseClient
        .Open
        Set AdoRecordset = .OpenSchema(adSchemaColumns)
    End With
    AdoRecordset.Filter = RstFilter
    
    Dim AdoStream As ADODB.Stream
    Set AdoStream = New ADODB.Stream
    AdoRecordset.Save AdoStream, adPersistXML

    Dim FilteredRecordset As ADODB.Recordset
    Set FilteredRecordset = New ADODB.Recordset
    FilteredRecordset.Open AdoStream
    
    Debug.Print "RecordCount: " & CStr(AdoRecordset.RecordCount)
    
    DbRecordset.RecordsetToQT Buffer.Range("A1"), FilteredRecordset
End Sub
