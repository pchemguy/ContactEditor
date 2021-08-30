VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCommandBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object intended to be used as a private instance field in a class implementing the IDbCommand interface (composition in place of inheritance)."
'@Folder "SecureADODB.DbCommand"
'@ModuleDescription("An object intended to be used as a private instance field in a class implementing the IDbCommand interface (composition in place of inheritance).")
'@PredeclaredId
'@IgnoreModule IllegalAnnotation
'@Exposed
Option Explicit

Implements IDbCommandBase

Private Type TCommand
    ParameterProvider As IParameterProvider
End Type
Private this As TCommand


'''' @Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal provider As IParameterProvider) As IDbCommandBase
    Dim Instance As DbCommandBase
    Set Instance = New DbCommandBase
    Instance.Init provider
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal provider As IParameterProvider)
Attribute Init.VB_Description = "Default constructor"
    Guard.NullReference provider
    Set this.ParameterProvider = provider
End Sub


Private Function CreateCommand(ByVal db As IDbConnection, _
                               ByVal CommandType As ADODB.CommandTypeEnum, _
                               ByVal SQL As String, _
                               ByRef args() As Variant) As ADODB.Command
    Guard.NullReference db
    Guard.Expression db.State = adStateOpen, Source:="DbCommandBase", Message:="Connection is not open."
    Guard.EmptyString SQL
    Guard.Expression ValidateOrdinalArguments(SQL, args), _
                     Source:="DbCommandBase", _
                     Message:="Arguments supplied are inconsistent with the provided command string parameters."
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = db.AdoConnection
    cmd.CommandType = CommandType
    cmd.CommandText = SQL
    cmd.Prepared = True
    
    Dim arg As ADODB.Parameter
    For Each arg In this.ParameterProvider.FromValues(args)
        cmd.Parameters.Append arg
    Next
    
    Set CreateCommand = cmd
End Function


Public Function ValidateOrdinalArguments(ByVal SQL As String, ByRef args() As Variant) As Boolean
    Dim Expected As Long
    Expected = Len(SQL) - Len(Replace(SQL, "?", vbNullString))
    
    Dim Actual As Long
    Actual = UBound(args) + (1 - LBound(args))
    
    ValidateOrdinalArguments = (Expected = Actual)
End Function


'@Ignore ProcedureNotUsed
Public Property Get Self() As IDbCommandBase
    Set Self = Me
End Property


'@Description("Returns class reference.")
'@Ignore ProcedureNotUsed
Public Property Get Class() As DbCommandBase
Attribute Class.VB_Description = "Returns class reference."
    Set Class = DbCommandBase
End Property


Private Property Get IDbCommandBase_Self() As IDbCommandBase
    Set IDbCommandBase_Self = Me
End Property


'@Description("Returns class reference.")
Private Property Get IDbCommandBase_Class() As Object
Attribute IDbCommandBase_Class.VB_Description = "Returns class reference."
    Set IDbCommandBase_Class = DbCommandBase
End Property


Private Function IDbCommandBase_CreateCommand(ByVal db As IDbConnection, _
                                              ByVal CommandType As ADODB.CommandTypeEnum, _
                                              ByVal SQL As String, _
                                              ByRef args() As Variant) As ADODB.Command
    Set IDbCommandBase_CreateCommand = CreateCommand(db, CommandType, SQL, args)
End Function
