VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbCommandBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbCommandBase implementation."
'@Folder "SecureADODB.DbCommand.Tests"
'@ModuleDescription("A stub acting as a IDbCommandBase implementation.")
Option Explicit

Implements IDbCommandBase

Private Type TInvokeState
    CreateCommandInvokes As Long
    ValidateOrdinalArgumentsInvokes As Long
End Type
Private this As TInvokeState


'@Ignore ProcedureNotUsed
Public Property Get CreateCommandInvokes() As Long
    CreateCommandInvokes = this.CreateCommandInvokes
End Property

'@Ignore ProcedureNotUsed
Public Property Get ValidateOrdinalArgumentsInvokes() As Long
    ValidateOrdinalArgumentsInvokes = this.ValidateOrdinalArgumentsInvokes
End Property

Private Function IDbCommandBase_CreateCommand( _
            ByVal db As IDbConnection, _
            ByVal CommandType As ADODB.CommandTypeEnum, _
            ByVal SQL As String, _
            ADODBParamsValues() As Variant) As ADODB.Command
    this.CreateCommandInvokes = this.CreateCommandInvokes + 1
    Dim Result As ADODB.Command
    Set Result = New ADODB.Command
    Result.CommandType = CommandType
    Result.CommandText = SQL
    Set IDbCommandBase_CreateCommand = Result
End Function
