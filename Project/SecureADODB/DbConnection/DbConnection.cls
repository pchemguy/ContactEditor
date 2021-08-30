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
'@ModuleDescription("A thin wrapper around an ADODB connection.")
'@PredeclaredId
'@IgnoreModule AssignmentNotUsed, IllegalAnnotation
'@Exposed
Option Explicit

Implements IDbConnection

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents AdoConnection As ADODB.Connection
Attribute AdoConnection.VB_VarHelpID = -1
Private Type TDbConnection
    ExecuteStatus As ADODB.EventStatusEnum
    RecordsAffected As Long
    '''' If a particular backend does not support transaction, "begin" raises
    '''' an error, and the corresponding handler sets this flag
    TransactionsDisabled As Boolean
    HasActiveTransaction As Boolean
    StateMap As Variant    '''' Map for ADODB.ObjectStateEnum
    LogController As ILogger
End Type
Private this As TDbConnection


'''' @Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
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
    
    Set this.LogController = LogController
    this.TransactionsDisabled = False
    On Error GoTo ConnectionError
    AdoConnection.Open connString
    On Error GoTo TransactionCheckError
    Dim TransactionDDL As Long
    '@Ignore ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
    TransactionDDL = AdoConnection.Properties("Transaction DDL")
    this.TransactionsDisabled = IIf(TransactionDDL = 0, True, False)
    On Error GoTo 0
    Exit Sub
    
TransactionCheckError:
    this.TransactionsDisabled = True
    Resume
    
ConnectionError:
    If Err.Number = &H80004005 Then
        Guard.ConnectionString Nothing
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


'@Ignore ProcedureNotUsed
Public Property Get Self() As IDbConnection
    Set Self = Me
End Property


'@Description("Returns class reference.")
'@Ignore ProcedureNotUsed
Public Property Get Class() As DbConnection
Attribute Class.VB_Description = "Returns class reference."
    Set Class = DbConnection
End Property


Private Sub Class_Initialize()
    Set AdoConnection = New ADODB.Connection
    '@Ignore EmptyStringLiteral
    this.StateMap = Array("adStateClosed", "adStateOpen", "adStateConnecting", "", "adStateExecuting", "", "", "", "adStateFetching")
End Sub


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


Private Sub LogAdoErrors()
    If this.LogController Is Nothing Then Exit Sub
    
    Dim AdoError As ADODB.Error
    Dim Message As String
    Message = "ADODB.Connection State: " + CStr(this.StateMap(AdoConnection.State))
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


'================================================================================'
'============================ IDbConnection INTERFACE ==========================='
'================================================================================'

Private Property Get IDbConnection_Self() As IDbConnection
    Set IDbConnection_Self = Me
End Property


'@Description("Returns class reference.")
Private Property Get IDbConnection_Class() As DbConnection
Attribute IDbConnection_Class.VB_Description = "Returns class reference."
    Set IDbConnection_Class = DbConnection
End Property


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
    If this.TransactionsDisabled Then Exit Sub
    Guard.NotInTransaction this.HasActiveTransaction
    
    AdoConnection.CommitTrans
    this.HasActiveTransaction = False
    LogAdoErrors
End Sub


Private Sub IDbConnection_RollbackTransaction()
    If this.TransactionsDisabled Then Exit Sub
    Guard.NotInTransaction this.HasActiveTransaction
    
    AdoConnection.RollbackTrans
    this.HasActiveTransaction = False
    LogAdoErrors
End Sub


'================================================================================'
'========================= AdoConnection EVENT HANDLERS ========================='
'================================================================================'

Private Sub AdoConnection_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Commit transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Rollback transaction completed."
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin connect..."
End Sub

Private Sub AdoConnection_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Connect completed. Status: " & adStatus
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Disconnect completed. Status: " & adStatus
End Sub

Private Sub AdoConnection_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin execute..."
End Sub

Private Sub AdoConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    this.ExecuteStatus = adStatus
    this.RecordsAffected = RecordsAffected
    Debug.Print "Execute completed, " & RecordsAffected & " record(s) affected."
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If pError.Number = 0 Then
        Debug.Print "Info: " & pError.Description
    Else
        Debug.Print "Error: " & pError.Number, pError.Description
    End If
End Sub
