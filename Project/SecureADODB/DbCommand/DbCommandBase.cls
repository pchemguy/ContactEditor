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
'@Exposed
Option Explicit

Implements IDbCommandBase

Private Type TCommand
    ParameterProvider As IDbParameters
End Type
Private this As TCommand


Public Function Create(ByVal provider As IDbParameters) As IDbCommandBase
    Dim Instance As DbCommandBase
    Set Instance = New DbCommandBase
    Instance.Init provider
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal provider As IDbParameters)
Attribute Init.VB_Description = "Default constructor"
    Guard.NullReference provider
    Set this.ParameterProvider = provider
End Sub


Private Function IDbCommandBase_CreateCommand(ByVal db As IDbConnection, _
                                              ByVal CommandType As ADODB.CommandTypeEnum, _
                                              ByVal SQL As String, _
                                              ByRef ADODBParamsValues() As Variant) As ADODB.Command
    Guard.NullReference db
    Guard.Expression db.State = adStateOpen, Source:="DbCommandBase", Message:="Connection is not open."
    Guard.EmptyString SQL
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = db.AdoConnection
        .CommandType = CommandType
        .CommandText = SQL
        .Prepared = True
    End With
    this.ParameterProvider.FromValues cmd, ADODBParamsValues
    
    Set IDbCommandBase_CreateCommand = cmd
End Function
