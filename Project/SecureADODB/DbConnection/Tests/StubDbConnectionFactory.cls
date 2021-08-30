VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbConnectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbConnectionFactory implementation."
'@Folder "SecureADODB.DbConnection.Tests"
'@ModuleDescription("A stub acting as a IDbConnectionFactory implementation.")
Option Explicit

Implements IDbConnectionFactory

Private Type TInvokeState
    CreateConnectionInvokes As Long
End Type
Private this As TInvokeState


Private Function IDbConnectionFactory_CreateInstance(ByVal connString As String, Optional ByVal LogController As ILogger = Nothing) As IDbConnection
    this.CreateConnectionInvokes = this.CreateConnectionInvokes + 1
    Set IDbConnectionFactory_CreateInstance = New StubDbConnection
End Function

'@Ignore ProcedureNotUsed
Public Property Get CreateConnectionInvokes() As Long
    CreateConnectionInvokes = this.CreateConnectionInvokes
End Property
