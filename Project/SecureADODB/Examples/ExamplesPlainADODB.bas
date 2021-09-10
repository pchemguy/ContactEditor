Attribute VB_Name = "ExamplesPlainADODB"
'@Folder "SecureADODB.Examples"
'@IgnoreModule
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "SecureADODB"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


Private Sub SQLiteRecordSetOpenBasicTest()
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim adoConnStr As String
    Dim sSQL As String
    
    sDriver = "SQLite3 ODBC Driver"
    sDatabaseExt = ".db"
    sTable = "people"
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & sDatabaseExt
    adoConnStr = "Driver=" & sDriver & ";" & _
                 "Database=" & sDatabase & ";"
    
    sSQL = "SELECT * FROM """ & sTable & """"
        
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseClient
    AdoRecordset.Open _
            Source:=sSQL, _
            ActiveConnection:=adoConnStr, _
            CursorType:=adOpenKeyset, _
            LockType:=adLockReadOnly, _
            Options:=(adCmdText Or adAsyncFetch)
    Set AdoRecordset.ActiveConnection = Nothing
End Sub


Private Sub CSVRecordSetOpenBasicTest()
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim adoConnStr As String
    Dim sSQL As String
    
    #If Win64 Then
        sDriver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        sDriver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If
    sDatabaseExt = ".csv"
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX
    sTable = LIB_NAME & sDatabaseExt
    adoConnStr = "Driver=" & sDriver & ";" & _
                 "DefaultDir=" & sDatabase & ";"
    
    sSQL = "SELECT * FROM """ & sTable & """"
    sSQL = sTable
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseClient
    AdoRecordset.Open _
            Source:=sSQL, _
            ActiveConnection:=adoConnStr, _
            CursorType:=adOpenKeyset, _
            LockType:=adLockReadOnly, _
            Options:=(adCmdTable Or adAsyncFetch)
    Set AdoRecordset.ActiveConnection = Nothing
End Sub


Private Sub CSVRecordSetOpenBasicTest2()
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim adoConnStr As String
    Dim sSQL As String
    
    #If Win64 Then
        sDriver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        sDriver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If
    sDatabaseExt = ".csv"
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX
    sTable = LIB_NAME & sDatabaseExt
    adoConnStr = "Driver=" & sDriver & ";" & _
                 "DefaultDir=" & sDatabase & ";"
    
    sSQL = "SELECT * FROM """ & sTable & """"
    sSQL = sTable
    
    Dim AdoConnection As ADODB.Connection
    Set AdoConnection = New ADODB.Connection
    AdoConnection.ConnectionString = adoConnStr
    
    On Error Resume Next
    AdoConnection.Open
    Debug.Print AdoConnection.Errors.Count
    Debug.Print AdoConnection.Properties("Transaction DDL")
    AdoConnection.BeginTrans
    On Error GoTo 0
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseClient
    AdoRecordset.Open _
            Source:=sSQL, _
            ActiveConnection:=adoConnStr, _
            CursorType:=adOpenKeyset, _
            LockType:=adLockReadOnly, _
            Options:=(adCmdTable Or adAsyncFetch)
    Set AdoRecordset.ActiveConnection = Nothing
End Sub


Private Sub SQLiteRecordSetOpenBasicTest2()
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim adoConnStr As String
    Dim sSQL As String
    
    sDriver = "SQLite3 ODBC Driver"
    sDatabaseExt = ".db"
    sTable = "people"
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & sDatabaseExt
    adoConnStr = "Driver=" & sDriver & ";" & _
                 "Database=" & sDatabase & ";"
    
    sSQL = "SELECT * FROM """ & sTable & """"
        
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseServer
    AdoRecordset.Open _
            Source:=sSQL, _
            ActiveConnection:=adoConnStr, _
            CursorType:=adOpenKeyset, _
            LockType:=adLockReadOnly, _
            Options:=(adCmdText Or adAsyncFetch)
    On Error Resume Next
    Set AdoRecordset.ActiveConnection = Nothing
    On Error GoTo 0
End Sub


Private Sub SQLiteRecordSetOpenTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & ".db"
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    sSQL = "SELECT * FROM people WHERE id <= 45 AND last_name <> 'machinery'"
    
    Dim AdoConnection As ADODB.Connection
    Set AdoConnection = New ADODB.Connection
    On Error Resume Next
    AdoConnection.Open adoConnStr
    On Error GoTo 0
    If AdoConnection.State = ADODB.ObjectStateEnum.adStateOpen Then AdoConnection.Close
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseClient
    AdoRecordset.Open Source:=sSQL, _
                      ActiveConnection:=adoConnStr, _
                      CursorType:=adOpenKeyset, _
                      LockType:=adLockReadOnly, _
                      Options:=(adCmdText Or adAsyncFetch)
    Set AdoRecordset.ActiveConnection = Nothing
End Sub


Private Sub SQLiteRecordSetOpenCommandSourceTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & ".db"
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    sSQL = "SELECT * FROM people WHERE id <= 45 AND last_name <> 'machinery'"
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    
    With AdoCommand
        .CommandType = adCmdText
        .CommandText = sSQL
        .ActiveConnection = adoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With
    
    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    AdoCommand.ActiveConnection.Close
End Sub


' Could not make it to work with named parameters
Private Sub SQLiteRecordSetOpenCommandSourceTwoParameterTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & ".db"
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    sSQL = "SELECT * FROM people WHERE id <= ? AND last_name <> ?"
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    
    Dim Mappings As ITypeMap
    Set Mappings = AdoTypeMappings.Default
    Dim provider As IDbParameters
    Set provider = DbParameters.Create(Mappings)
    provider.FromValues AdoCommand, 45, "Simon"
    
    With AdoCommand
        .CommandType = adCmdText
        .CommandText = sSQL
        .Prepared = True
        '.NamedParameters = True
        .ActiveConnection = adoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With
        
    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    AdoCommand.ActiveConnection.Close
    Debug.Print "RecordCount: " & CStr(AdoRecordset.RecordCount)
End Sub


Private Sub SQLiteRecordSetOpenCmdSrc2ParamsTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & ".db"
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    sSQL = "SELECT * FROM people WHERE id <= ? AND last_name <> ?"
    
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    
    Dim Mappings As ITypeMap
    Set Mappings = AdoTypeMappings.Default
    Dim provider As IDbParameters
    Set provider = DbParameters.Create(Mappings)
    provider.FromValues AdoCommand, 10, "Ivanov"
    provider.FromValues AdoCommand, 45, "Simon"
    
    With AdoCommand
        .CommandType = adCmdText
        .CommandText = sSQL
        .Prepared = True
        .ActiveConnection = adoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With
        
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    Set AdoCommand.ActiveConnection = Nothing

    Debug.Print "RecordCount: " & CStr(AdoRecordset.RecordCount)
'    DbRecordset.RecordsetToQT Buffer.Range("A1"), AdoRecordset
End Sub


Private Sub SQLiteConnectionTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & ".db"
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    Dim AdoConnection As ADODB.Connection
    Set AdoConnection = New ADODB.Connection
    With AdoConnection
        .ConnectionString = adoConnStr
        .CursorLocation = adUseClient
        .Open
    End With
End Sub
