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
'@IgnoreModule IllegalAnnotation
Option Explicit

Implements IDbCommand

Private Type TCommand
    cmd As ADODB.Command
    Base As IDbCommandBase
    Connection As IDbConnection
End Type
Private this As TCommand


'''' @Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
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


'@Ignore ProcedureNotUsed
Public Property Get Self() As IDbCommand
    Set Self = Me
End Property


'@Description("Returns class reference.")
'@Ignore ProcedureNotUsed
Public Property Get Class() As DbCommand
Attribute Class.VB_Description = "Returns class reference."
    Set Class = DbCommand
End Property


Private Property Get IDbCommand_Self() As IDbCommand
    Set IDbCommand_Self = Me
End Property


'@Description("Returns class reference.")
Private Property Get IDbCommand_Class() As Object
Attribute IDbCommand_Class.VB_Description = "Returns class reference."
    Set IDbCommand_Class = DbCommand
End Property


Public Function IDbCommand_AdoCommand(ByVal SQL As String, ParamArray args()) As ADODB.Command
    '''' Recordset routines pass their (sql, ParamArray) for processing here.
    '''' If the first item in args is also an array, assume ParamArray to ParamArray
    '''' delegation. Should unfold the outer array.
    Dim localArgs() As Variant
    localArgs = UnfoldParamArray(args)
    
    Set this.cmd = this.Base.CreateCommand(this.Connection, adCmdText, SQL, localArgs)
    Set IDbCommand_AdoCommand = this.cmd
End Function


Private Sub IDbCommand_ExecuteNonQuery(ByVal SQL As String, ParamArray args() As Variant)
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = UnfoldParamArray(args)
    
    Set this.cmd = this.Base.CreateCommand(this.Connection, adCmdText, SQL, localArgs)
    '@Ignore FunctionReturnValueDiscarded
    this.cmd.Execute , ADODB.ExecuteOptionEnum.adExecuteNoRecords
End Sub
