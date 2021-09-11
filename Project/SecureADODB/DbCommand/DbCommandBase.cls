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

Private Type TDbCommandBase
    ParameterProvider As IDbParameters
End Type
Private this As TDbCommandBase


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


Private Function IDbCommandBase_CreateCommand( _
            ByVal db As IDbConnection, _
            ByVal CommandType As ADODB.CommandTypeEnum, _
            ByVal SQL As String, _
            ByRef ADODBParamsValues() As Variant) As ADODB.Command
    Guard.NullReference db
    Guard.Expression db.State = adStateOpen, _
                     Source:="DbCommandBase", _
                     Message:="Connection is not open."
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = db.AdoConnection
    cmd.Prepared = True
    If Len(SQL) > 0 Then
        cmd.CommandType = CommandType
        cmd.CommandText = SQL
        this.ParameterProvider.FromValues cmd, ADODBParamsValues
    End If
    
    Set IDbCommandBase_CreateCommand = cmd
End Function


Private Function IDbCommandBase_UpdateCommand( _
            ByVal cmd As ADODB.Command, _
            ByVal SQL As String, _
            ByRef ADODBParamsValues() As Variant) As ADODB.Command
    Guard.NullReference cmd
    
    If Len(SQL) > 0 And cmd.CommandText <> SQL Then cmd.CommandText = SQL
    If UBound(ADODBParamsValues) - LBound(ADODBParamsValues) + 1 > 0 Then
        this.ParameterProvider.FromValues cmd, ADODBParamsValues
    End If
    Set IDbCommandBase_UpdateCommand = cmd
End Function


Private Function IDbCommandBase_GetSQL(ByVal AdoCommand As ADODB.Command) As String
    IDbCommandBase_GetSQL = this.ParameterProvider.GetSQL(AdoCommand)
End Function
