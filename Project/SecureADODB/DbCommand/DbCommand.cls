VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close or disconnected recordset."
'@Folder "SecureADODB.DbCommand"
'@ModuleDescription("An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close or disconnected recordset.")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDbCommand

Private Type TCommand
    cmd As ADODB.Command
    Base As IDbCommandBase
    Connection As IDbConnection
End Type
Private this As TCommand


Public Function Create(ByVal db As IDbConnection, ByVal baseCommand As IDbCommandBase) As IDbCommand
    Dim Instance As DbCommand
    Set Instance = New DbCommand
    Instance.Init db, baseCommand
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal db As IDbConnection, ByVal baseCommand As IDbCommandBase)
Attribute Init.VB_Description = "Default constructor"
    Guard.NullReference db
    Guard.NullReference baseCommand

    Set this.Connection = db
    Set this.Base = baseCommand
End Sub


Private Function IDbCommand_AdoCommand( _
            ByVal SQL As String, _
            ParamArray ADODBParamsValues()) As ADODB.Command
    '''' Recordset routines pass their (sql, ParamArray) for processing here.
    '''' If the first item in args is also an array, assume ParamArray to
    '''' ParamArray delegation. Should unfold the outer array.
    Dim localArgs() As Variant
    localArgs = UnfoldParamArray(ADODBParamsValues)
    
    If this.cmd Is Nothing Then
        Set this.cmd = this.Base.CreateCommand( _
                           this.Connection, adCmdText, SQL, localArgs)
    Else
        Set this.cmd = this.Base.UpdateCommand( _
                           this.cmd, SQL, localArgs)
    End If
    Set IDbCommand_AdoCommand = this.cmd
End Function


Private Sub IDbCommand_ExecuteNonQuery(ByVal SQL As String, ParamArray ADODBParamsValues() As Variant)
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = UnfoldParamArray(ADODBParamsValues)
    
    Set this.cmd = IDbCommand_AdoCommand(SQL, localArgs)
    '@Ignore FunctionReturnValueDiscarded
    this.cmd.Execute Options:=ADODB.ExecuteOptionEnum.adExecuteNoRecords
End Sub


Private Property Get IDbCommand_Connection() As IDbConnection
    Set IDbCommand_Connection = this.Connection
End Property


Private Property Get IDbCommand_SQL() As String
    IDbCommand_SQL = this.Base.GetSQL(this.cmd)
End Property
