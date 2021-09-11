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
    UpdateCommandInvokes As Long
    ParameterProviderInvokes As Long
    ValidateOrdinalArgumentsInvokes As Long
End Type
Private this As TInvokeState


Public Property Get CreateCommandInvokes() As Long
    CreateCommandInvokes = this.CreateCommandInvokes
End Property

Public Property Get UpdateCommandInvokes() As Long
    UpdateCommandInvokes = this.UpdateCommandInvokes
End Property

Public Property Get ParameterProviderInvokes() As Long
    ParameterProviderInvokes = this.ParameterProviderInvokes
End Property

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


Private Function IDbCommandBase_UpdateCommand(ByVal cmd As ADODB.Command, _
                                              ByVal SQL As String, _
                                              ByRef ADODBParamsValues() As Variant) As ADODB.Command
    this.UpdateCommandInvokes = this.UpdateCommandInvokes + 1
    If Len(SQL) > 0 And cmd.CommandText <> SQL Then cmd.CommandText = SQL
    If UBound(ADODBParamsValues) - LBound(ADODBParamsValues) + 1 > 0 Then
        this.ParameterProviderInvokes = this.ParameterProviderInvokes + 1
    End If
    Set IDbCommandBase_UpdateCommand = cmd
End Function

Private Function IDbCommandBase_GetSQL(ByVal AdoCommand As ADODB.Command) As String
    IDbCommandBase_GetSQL = AdoCommand.CommandText
End Function
