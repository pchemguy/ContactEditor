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
'@IgnoreModule ProcedureNotUsed
Option Explicit

Implements IDbCommandBase

Private Type TInvokeState
    CreateCommandInvokes As Long
    ValidateOrdinalArgumentsInvokes As Long
End Type
Private this As TInvokeState


Public Property Get CreateCommandInvokes() As Long
    CreateCommandInvokes = this.CreateCommandInvokes
End Property

Public Property Get ValidateOrdinalArgumentsInvokes() As Long
    ValidateOrdinalArgumentsInvokes = this.ValidateOrdinalArgumentsInvokes
End Property

Private Property Get IDbCommandBase_Self() As IDbCommandBase
    Set IDbCommandBase_Self = Me
End Property

'@Description("Returns class reference.")
Private Property Get IDbCommandBase_Class() As Object
Attribute IDbCommandBase_Class.VB_Description = "Returns class reference."
    Set IDbCommandBase_Class = DbCommandBase
End Property

Private Function IDbCommandBase_CreateCommand(ByVal db As IDbConnection, ByVal CommandType As ADODB.CommandTypeEnum, ByVal SQL As String, args() As Variant) As ADODB.Command
    this.CreateCommandInvokes = this.CreateCommandInvokes + 1
    Dim Result As ADODB.Command
    Set Result = New ADODB.Command
    Result.CommandType = CommandType
    Result.CommandText = SQL
    Set IDbCommandBase_CreateCommand = Result
End Function
