Attribute VB_Name = "ExamplesPlainADODB"
'@Folder "SecureADODB.DbManager.Examples"
'@IgnoreModule
Option Explicit


Public Sub SQLiteRecordSetOpenBasicTest()
    Dim fso As New Scripting.FileSystemObject
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim adoConnStr As String
    Dim sSQL As String
    
    sDriver = "SQLite3 ODBC Driver"
    sDatabaseExt = ".db"
    sTable = "people"
    sDatabase = ThisWorkbook.Path & Application.PathSeparator & fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
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


Public Sub CSVRecordSetOpenBasicTest()
    Dim fso As New Scripting.FileSystemObject
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
    sDatabase = ThisWorkbook.Path
    sTable = fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
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


Public Sub CSVRecordSetOpenBasicTest2()
    Dim fso As New Scripting.FileSystemObject
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
    sDatabase = ThisWorkbook.Path
    sTable = fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
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


Public Sub SQLiteRecordSetOpenBasicTest2()
    Dim fso As New Scripting.FileSystemObject
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim adoConnStr As String
    Dim sSQL As String
    
    sDriver = "SQLite3 ODBC Driver"
    sDatabaseExt = ".db"
    sTable = "people"
    sDatabase = ThisWorkbook.Path & Application.PathSeparator & fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
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


Public Sub SQLiteRecordSetOpenTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path + "\" + "SecureADODB.db"
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


Public Sub SQLiteRecordSetOpenCommandSourceTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path + "\" + "SecureADODB.db"
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
Public Sub SQLiteRecordSetOpenCommandSourceTwoParameterTest()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path + "\" + "SecureADODB.db"
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
    
    Dim mappings As ITypeMap
    Set mappings = AdoTypeMappings.Default
    Dim provider As IParameterProvider
    Set provider = AdoParameterProvider.Create(mappings)
    
    Dim adoParameter As ADODB.Parameter
    Set adoParameter = provider.FromValue(45)
    'adoParameter.name = "@category_id"
    AdoCommand.Parameters.Append adoParameter
    Set adoParameter = provider.FromValue("machinery")
    'adoParameter.name = "@section"
    AdoCommand.Parameters.Append adoParameter
    
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
