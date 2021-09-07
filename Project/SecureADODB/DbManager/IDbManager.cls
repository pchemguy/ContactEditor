VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object encapsulating a database transaction."
'@Folder "SecureADODB.DbManager"
'@ModuleDescription("Represents an object encapsulating a database transaction.")
'@Interface
'@Exposed
Option Explicit


'@Description("Returns logger reference")
Public Property Get LogController() As ILogger
Attribute LogController.VB_Description = "Returns logger reference"
End Property

'@Description("Returns connection reference")
Public Property Get Connection() As IDbConnection
Attribute Connection.VB_Description = "Returns connection reference"
End Property

'@Description("Returns DbConnectionString reference")
Public Property Get DbConnStr() As DbConnectionString
Attribute DbConnStr.VB_Description = "Returns DbConnectionString reference"
End Property

'@Description("Returns DbMetaData reference")
Public Property Get DbMeta() As DbMetaData
Attribute DbMeta.VB_Description = "Returns DbMetaData reference"
End Property

'@Description("Creates a new command to execute.")
Public Function Command() As IDbCommand
Attribute Command.VB_Description = "Creates a new command to execute."
End Function

'@Description("Creates a new recordset to execute.")
Public Function Recordset(Optional ByVal Disconnected As Boolean = True, _
                          Optional ByVal CacheSize As Long = 10, _
                          Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
                          Optional ByVal LockType As ADODB.LockTypeEnum = adLockReadOnly, _
                          Optional ByVal AsyncMode As Boolean = False, _
                          Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0) As IDbRecordset
Attribute Recordset.VB_Description = "Creates a new recordset to execute."
End Function

'@Description("Begins the transaction.")
Public Sub Begin()
Attribute Begin.VB_Description = "Begins the transaction."
End Sub

'@Description("Commits the transaction.")
Public Sub Commit()
Attribute Commit.VB_Description = "Commits the transaction."
End Sub

'@Description("Rolls back the transaction.")
Public Sub Rollback()
Attribute Rollback.VB_Description = "Rolls back the transaction."
End Sub
