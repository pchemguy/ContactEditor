VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A thin wrapper around an ADODB connection."
'@Folder "SecureADODB.DbConnection"
'@ModuleDescription "A thin wrapper around an ADODB connection."
'@PredeclaredId
'@IgnoreModule AssignmentNotUsed
'@Exposed
Option Explicit

Implements IDbConnection

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents AdoConnection As ADODB.Connection
Attribute AdoConnection.VB_VarHelpID = -1
Private Type TDbConnection
    ExecuteStatus As ADODB.EventStatusEnum
    '''' This attribute is set to a value received by a Connection event handler
    RecordsAffected As Long
    '''' If a particular backend does not support transaction, "begin" raises
    '''' an error, and the corresponding handler sets this flag
    TransactionsDisabled As Boolean
    HasActiveTransaction As Boolean
    LogController As ILogger
    '''' This attribute is used to query affected rows count across a transaction
    TransRecordsAffected As Long
    '''' Expected count of changes over the transaction
    '''' Typically set by the manager or DbRecordset object
    ExpectedRecordsAffected As Long
    '''' This scalar query command is used to query the total changes count. Run
    '''' by the transaction begin/commit handlers to yield the count of rows
    '''' affected during the transaction
    cmdAffectedRows As ADODB.Command
    Engine As String
End Type
Private this As TDbConnection


Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    If this.HasActiveTransaction Then AdoConnection.RollbackTrans
    AdoConnection.Close
    Set AdoConnection = Nothing
    
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    LogAdoErrors
    Resume Next
End Sub


Public Function Create(ByVal connString As String, _
              Optional ByVal LogController As ILogger = Nothing) As IDbConnection
    Dim Instance As DbConnection
    Set Instance = New DbConnection
    Instance.Init connString, LogController
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal connString As String, Optional ByVal LogController As ILogger = Nothing)
Attribute Init.VB_Description = "Default constructor"
    Guard.ConnectionString connString

    Set AdoConnection = New ADODB.Connection
    Set this.LogController = LogController
    this.ExpectedRecordsAffected = -1
    
    On Error GoTo ConnectionError
    AdoConnection.Open connString
    
    Dim TransactionDDL As Long
    TransactionDDL = 0
    On Error Resume Next
    '@Ignore ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
    TransactionDDL = AdoConnection.Properties("Transaction DDL")
    On Error GoTo 0
    this.TransactionsDisabled = IIf(TransactionDDL = 0, True, False)
    
    Exit Sub
    
ConnectionError:
    If Err.Number = &H80004005 Then
        '''' Raise invalid connection string error
        Guard.ConnectionString vbNullString
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


Private Sub LogAdoErrors()
    If this.LogController Is Nothing Then Exit Sub
    
    Dim StateMap As Variant
    StateMap = Array("adStateClosed", "adStateOpen", "adStateConnecting", _
                     vbNullString, "adStateExecuting", vbNullString, _
                     vbNullString, vbNullString, "adStateFetching")
    
    Dim AdoError As ADODB.Error
    Dim Message As String
    Message = "ADODB.Connection State: " + CStr(StateMap(AdoConnection.State))
    this.LogController.Log Message
    For Each AdoError In AdoConnection.Errors
        With AdoError
            Message = "ADODB.Error #" + CStr(.Number) + ", " + _
                      "SQLState: " + CStr(.SqlState) + ", " + _
                      "Source: " + CStr(.Source) + ", " + _
                      "Description: " + CStr(.Description)
        End With
        this.LogController.Log Message
    Next AdoError
    AdoConnection.Errors.Clear
        
    With VBA.Err
        If .Number > 0 And Left$(.Source, 6) = "ADODB." Then
            Message = "VBA.Err #" + CStr(.Number) + ", " + _
                      "Source: " + CStr(.Source) + ", " + _
                      "Description: " + CStr(.Description)
        End If
    End With
    this.LogController.Log Message
    VBA.Err.Clear
End Sub


'@Description "If possible, queries the database for total changes count."
Friend Function TotalChanges() As Long
Attribute TotalChanges.VB_Description = "If possible, queries the database for total changes count."
    TotalChanges = -1
    If Not this.cmdAffectedRows Is Nothing Then
        On Error Resume Next
        TotalChanges = this.cmdAffectedRows.Execute.Fields.Item(0).Value
        On Error GoTo 0
    End If
End Function


'================================================================================'
'============================ IDbConnection INTERFACE ==========================='
'================================================================================'


Private Property Get IDbConnection_AdoConnection() As ADODB.Connection
    Set IDbConnection_AdoConnection = AdoConnection
End Property


Private Property Get IDbConnection_State() As ADODB.ObjectStateEnum
    IDbConnection_State = AdoConnection.State
End Property


Private Property Get IDbConnection_ExecuteStatus() As ADODB.EventStatusEnum
    IDbConnection_ExecuteStatus = this.ExecuteStatus
End Property


Private Property Get IDbConnection_RecordsAffected() As Long
    IDbConnection_RecordsAffected = this.RecordsAffected
End Property


Private Property Get IDbConnection_TransRecordsAffected() As Long
    IDbConnection_TransRecordsAffected = this.TransRecordsAffected
End Property


Private Property Get IDbConnection_Engine() As String
    IDbConnection_Engine = this.Engine
End Property


Private Property Get IDbConnection_ExpectedRecordsAffected() As Long
    IDbConnection_ExpectedRecordsAffected = this.ExpectedRecordsAffected
End Property


Private Property Let IDbConnection_ExpectedRecordsAffected(ByVal Value As Long)
    this.ExpectedRecordsAffected = Value
End Property


'@Description "Set database type [typically recieved from the manager]"
Private Property Let IDbConnection_Engine(ByVal EngineName As String)
Attribute IDbConnection_Engine.VB_Description = "Set database type [typically recieved from the manager]"
    this.Engine = EngineName
    '''' Set engine specific command for querying affected rows count
    If LCase$(EngineName) = "sqlite" Then
        '''' Set command for the SQLite engine
        Set this.cmdAffectedRows = New ADODB.Command
        With this.cmdAffectedRows
            .CommandType = adCmdText
            .Prepared = True
            .CommandText = "SELECT total_changes()"
            Set .ActiveConnection = AdoConnection
        End With
    End If
End Property


Private Sub IDbConnection_BeginTransaction()
    Guard.InTransaction this.HasActiveTransaction
    Guard.InvalidTransaction this.TransactionsDisabled

    On Error GoTo ErrorHandler
    '@Ignore VariableNotUsed
    Dim Result As Long
    '''' If  particular backend does not support transaction, the next line raises an error, and the corresponding handler sets this.TransactionsDisabled
    Result = AdoConnection.BeginTrans
    If Not this.TransactionsDisabled Then this.HasActiveTransaction = True
    LogAdoErrors
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    If VBA.Err.Number = 3251 Then
        this.TransactionsDisabled = True
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


Private Sub IDbConnection_CommitTransaction()
    Guard.NotInTransaction this.HasActiveTransaction
    
    AdoConnection.CommitTrans
    this.HasActiveTransaction = False
    LogAdoErrors
End Sub


Private Sub IDbConnection_RollbackTransaction()
    Guard.NotInTransaction this.HasActiveTransaction
    
    AdoConnection.RollbackTrans
    this.HasActiveTransaction = False
    LogAdoErrors
End Sub


'================================================================================'
'========================= AdoConnection EVENT HANDLERS ========================='
'================================================================================'


Private Sub AdoConnection_BeginTransComplete( _
            ByVal TransactionLevel As Long, _
            ByVal pError As ADODB.Error, _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
    '''' Queries the database for total changes count before the transaction
    this.TransRecordsAffected = TotalChanges()
End Sub

Private Sub AdoConnection_CommitTransComplete( _
            ByVal pError As ADODB.Error, _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pConnection As ADODB.Connection)
    Debug.Print "Commit transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
    With this
        '''' Queries the database for total changes count after the transaction
        '''' and calculates the difference, yielding transaction changes count
        .TransRecordsAffected = TotalChanges() - .TransRecordsAffected
        '''' If expected changes count is set, verify that it matches the value
        '''' reported by the database
        If .ExpectedRecordsAffected >= 0 Then
            Guard.Expression .ExpectedRecordsAffected = .TransRecordsAffected, _
                    "DbConnection", "Affected rows count mismatch"
            Debug.Print "Affected rows count (matched): " & CStr(.TransRecordsAffected)
        Else
            Debug.Print "Affected rows count: " & CStr(.TransRecordsAffected)
        End If
        .ExpectedRecordsAffected = -1
    End With
End Sub

Private Sub AdoConnection_RollbackTransComplete( _
            ByVal pError As ADODB.Error, _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pConnection As ADODB.Connection)
    Debug.Print "Rollback transaction completed."
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
    this.ExpectedRecordsAffected = -1
End Sub

Private Sub AdoConnection_WillConnect( _
            ByRef ConnectionString As String, _
            ByRef UserID As String, _
            ByRef Password As String, _
            ByRef Options As Long, _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin connect..."
End Sub

Private Sub AdoConnection_ConnectComplete( _
            ByVal pError As ADODB.Error, _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pConnection As ADODB.Connection)
    Debug.Print "Connect completed. Status: " & adStatus
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_Disconnect( _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pConnection As ADODB.Connection)
    Debug.Print "Disconnect completed. Status: " & adStatus
End Sub

Private Sub AdoConnection_WillExecute( _
            ByRef Source As String, _
            ByRef CursorType As ADODB.CursorTypeEnum, _
            ByRef LockType As ADODB.LockTypeEnum, _
            ByRef Options As Long, _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pCommand As ADODB.Command, _
            ByVal pRecordset As ADODB.Recordset, _
            ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin execute..."
End Sub

Private Sub AdoConnection_ExecuteComplete( _
            ByVal RecordsAffected As Long, _
            ByVal pError As ADODB.Error, _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pCommand As ADODB.Command, _
            ByVal pRecordset As ADODB.Recordset, _
            ByVal pConnection As ADODB.Connection)
    Debug.Print "Execute completed, " & RecordsAffected & " record(s) affected."
    this.ExecuteStatus = adStatus
    this.RecordsAffected = RecordsAffected
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_InfoMessage( _
            ByVal pError As ADODB.Error, _
            ByRef adStatus As ADODB.EventStatusEnum, _
            ByVal pConnection As ADODB.Connection)
    If pError.Number = 0 Then
        Debug.Print "Info: " & pError.Description
    Else
        Debug.Print "Error: " & pError.Number, pError.Description
    End If
End Sub
