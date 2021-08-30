VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object abstracting a database command."
'@Folder "SecureADODB.DbCommand.Abstract"
'@ModuleDescription("Represents an object abstracting a database command.")
'@Interface
'@Exposed
Option Explicit


'@Description("Returns an ADODB.Command, without executing it.")
Public Function AdoCommand(ByVal SQL As String, ParamArray args()) As ADODB.Command
Attribute AdoCommand.VB_Description = "Returns an ADODB.Command, without executing it."
End Function

'@Description("Executes the specified parameterized SQL statement/command.")
Public Sub ExecuteNonQuery(ByVal SQL As String, ParamArray args())
Attribute ExecuteNonQuery.VB_Description = "Executes the specified parameterized SQL statement/command."
End Sub

Public Property Get Self() As IDbCommand
End Property

'@Description("Returns class reference.")
Public Property Get Class() As Object
Attribute Class.VB_Description = "Returns class reference."
End Property
