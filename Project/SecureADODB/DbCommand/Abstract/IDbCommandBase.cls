VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbCommandBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Defines an explicit public interface for the DbCommandBase class."
'@Folder "SecureADODB.DbCommand.Abstract"
'@ModuleDescription("Defines an explicit public interface for the DbCommandBase class.")
'@Interface
'@Exposed
Option Explicit


'@Description("Creates and returns an ADODB.Command, without executing it.")
Public Function CreateCommand(ByVal db As IDbConnection, ByVal CommandType As ADODB.CommandTypeEnum, ByVal SQL As String, ByRef args() As Variant) As ADODB.Command
Attribute CreateCommand.VB_Description = "Creates and returns an ADODB.Command, without executing it."
End Function

Public Property Get Self() As IDbCommandBase
End Property

'@Description("Returns class reference.")
Public Property Get Class() As Object
Attribute Class.VB_Description = "Returns class reference."
End Property
