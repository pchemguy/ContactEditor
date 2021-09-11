VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbConnectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An abstract factory that creates IDbConnection objects."
'@Folder "SecureADODB.DbConnection.Abstract"
'@ModuleDescription("An abstract factory that creates IDbConnection objects.")
'@Interface
'@Exposed
Option Explicit

'@Description("Creates and returns a new database connection.")
Public Function CreateInstance(ByVal connString As String, Optional ByVal LogController As ILogger = Nothing) As IDbConnection
Attribute CreateInstance.VB_Description = "Creates and returns a new database connection."
End Function
