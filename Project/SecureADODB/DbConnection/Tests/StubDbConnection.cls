VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbConnection implementation."
'@Folder "SecureADODB.DbConnection.Tests"
'@ModuleDescription("A stub acting as a IDbConnection implementation.")
'@IgnoreModule ProcedureNotUsed
Option Explicit

Implements IDbConnection

Private stubState As ADODB.ObjectStateEnum
Private stubConnection As ADODB.Connection
Private beganTransaction As Boolean
Private committedTransaction As Boolean
Private rolledBackTransaction As Boolean


Public Property Get State() As ADODB.ObjectStateEnum
    State = stubState
End Property

Public Property Let State(ByVal Value As ADODB.ObjectStateEnum)
    stubState = Value
End Property

Public Property Get DidBeginTransaction() As Boolean
    DidBeginTransaction = beganTransaction
End Property

Public Property Get DidCommitTransaction() As Boolean
    DidCommitTransaction = committedTransaction
End Property

Public Property Get DidRollBackTransaction() As Boolean
    DidRollBackTransaction = rolledBackTransaction
End Property

Private Sub Class_Initialize()
    'needed because we don't want IDbConnection.AdoConnection to return Nothing.
    Set stubConnection = New ADODB.Connection
    stubState = adStateOpen
End Sub

Private Sub Class_Terminate()
    stubState = adStateClosed
    Set stubConnection = Nothing
End Sub

Private Property Get IDbConnection_AdoConnection() As ADODB.Connection
    Set IDbConnection_AdoConnection = stubConnection
End Property

Private Property Get IDbConnection_Self() As IDbConnection
    Set IDbConnection_Self = Me
End Property

Private Property Get IDbConnection_Class() As DbConnection
    Set IDbConnection_Class = DbConnection
End Property

Private Property Get IDbConnection_State() As ADODB.ObjectStateEnum
    IDbConnection_State = State
End Property

Private Property Get IDbConnection_ExecuteStatus() As ADODB.EventStatusEnum
    IDbConnection_ExecuteStatus = ADODB.EventStatusEnum.adStatusOK
End Property

Private Property Get IDbConnection_RecordsAffected() As Long
    IDbConnection_RecordsAffected = 1
End Property

Private Sub IDbConnection_BeginTransaction()
    Guard.InTransaction beganTransaction
    beganTransaction = True
End Sub

Private Sub IDbConnection_CommitTransaction()
    Guard.NotInTransaction beganTransaction And Not committedTransaction And Not rolledBackTransaction
    committedTransaction = True
End Sub

Private Sub IDbConnection_RollbackTransaction()
    Guard.NotInTransaction beganTransaction And Not committedTransaction And Not rolledBackTransaction
    rolledBackTransaction = True
End Sub
