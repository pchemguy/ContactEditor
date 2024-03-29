VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbCommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An abstract factory that creates IDbCommand objects."
'@Folder "SecureADODB.DbCommand.Abstract"
'@ModuleDescription("An abstract factory that creates IDbCommand objects.")
'@Interface
'@Exposed
Option Explicit

'@Description("Creates a new IDbCommand object.")
Public Function CreateInstance(ByVal db As IDbConnection) As IDbCommand
Attribute CreateInstance.VB_Description = "Creates a new IDbCommand object."
End Function
