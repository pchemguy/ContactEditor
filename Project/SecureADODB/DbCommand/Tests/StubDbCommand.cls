VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbCommand implementation."
'@Folder "SecureADODB.DbCommand.Tests"
'@ModuleDescription("A stub acting as a IDbCommand implementation.")
'@IgnoreModule ProcedureNotUsed
Option Explicit

Implements IDbCommand

Private Type TInvokeState
    stubCommand As ADODB.Command
    ExecuteNonQueryInvokes As Long
End Type
Private this As TInvokeState


Public Property Get ExecuteNonQueryInvokes() As Long
    ExecuteNonQueryInvokes = this.ExecuteNonQueryInvokes
End Property

Public Function IDbCommand_AdoCommand(ByVal SQL As String, ParamArray args()) As ADODB.Command
    Set IDbCommand_AdoCommand = this.stubCommand
End Function

Private Sub IDbCommand_ExecuteNonQuery(ByVal SQL As String, ParamArray args() As Variant)
    this.ExecuteNonQueryInvokes = this.ExecuteNonQueryInvokes + 1
End Sub

Private Property Get IDbCommand_Self() As IDbCommand
    Set IDbCommand_Self = Me
End Property

'@Description("Returns class reference.")
Private Property Get IDbCommand_Class() As Object
Attribute IDbCommand_Class.VB_Description = "Returns class reference."
    Set IDbCommand_Class = DbCommand
End Property
