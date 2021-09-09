VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object that wraps an active ADODB connection."
'@Folder "SecureADODB.DbConnection.Abstract"
'@ModuleDescription("Represents an object that wraps an active ADODB connection.")
'@Interface
'@Exposed
Option Explicit

'@Description("Gets the wrapped ADODB connection.")
Public Property Get AdoConnection() As ADODB.Connection
Attribute AdoConnection.VB_Description = "Gets the wrapped ADODB connection."
End Property

'@Description("Gets the state of the wrapped ADODB connection.")
Public Property Get State() As ADODB.ObjectStateEnum
Attribute State.VB_Description = "Gets the state of the wrapped ADODB connection."
End Property

'@Description("Gets the execute completed status of the wrapped ADODB connection.")
Public Property Get ExecuteStatus() As ADODB.EventStatusEnum
Attribute ExecuteStatus.VB_Description = "Gets the execute completed status of the wrapped ADODB connection."
End Property

'@Description("Gets the number of affected records [recieved from ADODB]")
Public Property Get RecordsAffected() As Long
Attribute RecordsAffected.VB_Description = "Gets the number of affected records [recieved from ADODB]"
End Property

'@Description("Gets the number of affected records [via a query]")
Public Property Get TransRecordsAffected() As Long
Attribute TransRecordsAffected.VB_Description = "Gets the number of affected records [via a query]"
End Property

Public Property Get ExpectedRecordsAffected() As Long
End Property

Public Property Let ExpectedRecordsAffected(ByVal Value As Long)
End Property

Public Property Get Engine() As String
End Property

Public Property Let Engine(ByVal EngineName As String)
End Property

Public Sub BeginTransaction()
End Sub

Public Sub CommitTransaction()
End Sub

Public Sub RollbackTransaction()
End Sub
