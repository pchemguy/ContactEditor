VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An implementation of an abstract factory that creates DbConnection objects."
'@Folder "SecureADODB.DbConnection"
'@ModuleDescription("An implementation of an abstract factory that creates DbConnection objects.")
'@Exposed
Option Explicit

Implements IDbConnectionFactory

Private Function IDbConnectionFactory_CreateInstance( _
        ByVal connString As String, _
        Optional ByVal LogController As ILogger = Nothing) As IDbConnection
    Set IDbConnectionFactory_CreateInstance = DbConnection.Create(connString, LogController)
End Function
