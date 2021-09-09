VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that encapsulates a database transaction."
'@Folder "SecureADODB.DbManager"
'@ModuleDescription "An object that encapsulates a database transaction."
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDbManager

Private Type TDbManager
    DbMeta As DbMetaData
    DbConnStr As DbConnectionString
    Connection As IDbConnection
    CommandFactory As IDbCommandFactory
    LogController As ILogger
End Type
Private this As TDbManager


'@Description("Default factory")
Public Function Create(ByVal db As IDbConnection, _
                       ByVal Factory As IDbCommandFactory, _
              Optional ByVal LogController As ILogger = Nothing) As IDbManager
Attribute Create.VB_Description = "Default factory"
    Dim Instance As DbManager
    Set Instance = New DbManager
    Instance.Init db, Factory, LogController
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal db As IDbConnection, _
                ByVal Factory As IDbCommandFactory, _
       Optional ByVal LogController As ILogger = Nothing)
Attribute Init.VB_Description = "Default constructor"
              
    Guard.NullReference Factory
    Guard.NullReference db
    Guard.Expression db.State = adStateOpen, Source:="DbManager", Message:="Connection should be open."

    Set this.LogController = LogController
    Set this.Connection = db
    Set this.CommandFactory = Factory
End Sub


'@Description("Initialize additional attributes")
Friend Sub InitExtra(Optional ByVal DbConnStr As DbConnectionString = Nothing)
Attribute InitExtra.VB_Description = "Initialize additional attributes"
    With this
        If Not DbConnStr Is Nothing Then Set .DbConnStr = DbConnStr
        If Not .DbConnStr Is Nothing Then Set .DbMeta = DbMetaData.Create(.DbConnStr)
    End With
End Sub


'''' Factory for file-based databases.
''''
'''' Args:
''''   DbType (string):
''''     Type of the database: "sqlite", "csv", "xls"
''''
''''   DbFileName (string, optional, ""):
''''     Database file name. If not provided, ThisWorkbook.VBProject.Name
''''     will be used. Extension is added based on the database type:
''''       "sqlite" - "db" or "sqlite"
''''       "csv"    - "xsv" or "csv"
''''       "xls"    - "xls
''''
''''   ConnectionOptions (variant, optional, Empty):
''''     Connection options. If Empty, default values is selected based on the
''''     database type (see DbConnectionString constructor code for details).
''''
''''  N.B.: "xls" backend support is not currently implemented
''''
'@Description "Factory for file-based databases"
Public Function CreateFileDb( _
                 ByVal DbType As String, _
        Optional ByVal DbFileName As String = vbNullString, _
        Optional ByVal ConnectionOptions As String = vbNullString, _
        Optional ByVal LoggerType As LoggerTypeEnum = LoggerTypeEnum.logGlobal _
        ) As IDbManager
Attribute CreateFileDb.VB_Description = "Factory for file-based databases"
    Dim LogController As ILogger
    Select Case LoggerType
        Case LoggerTypeEnum.logDisabled
            Set LogController = Nothing
        Case LoggerTypeEnum.logGlobal
            Set LogController = Logger
        Case LoggerTypeEnum.logPrivate
            Set LogController = Logger.Create
    End Select
    
    '''' CSV fails if String -> adVarWChar mapping is used
    ''''              String -> adVarChar must be used for CSV instead
    Dim provider As IDbParameters
    Set provider = DbParameters.Create( _
            IIf(LCase$(DbType) <> "csv", AdoTypeMappings.Default, AdoTypeMappings.CSV))
    
    Dim baseCommand As IDbCommandBase
    Set baseCommand = DbCommandBase.Create(provider)
    
    Dim Factory As IDbCommandFactory
    Set Factory = DbCommandFactory.Create(baseCommand)
    
    Dim DbConnStr As DbConnectionString
    Set DbConnStr = DbConnectionString.CreateFileDb(DbType, DbFileName, , ConnectionOptions)
    Dim db As IDbConnection
    Set db = DbConnection.Create(DbConnStr.ConnectionString, LogController)
    db.Engine = DbType
    
    Dim Instance As DbManager
    Set Instance = DbManager.Create(db, Factory, LogController)
    Instance.InitExtra DbConnStr
    
    Set CreateFileDb = Instance
End Function


'@Ignore ProcedureNotUsed
Public Property Get DbConnStr() As DbConnectionString
    Set DbConnStr = this.DbConnStr
End Property

'@Ignore ProcedureNotUsed
Public Property Get DbMeta() As DbMetaData
    Set DbMeta = this.DbMeta
End Property

Private Property Get IDbManager_LogController() As ILogger
    Set IDbManager_LogController = this.LogController
End Property

Private Property Get IDbManager_DbConnStr() As DbConnectionString
    Set IDbManager_DbConnStr = this.DbConnStr
End Property

Private Property Get IDbManager_DbMeta() As DbMetaData
    Set IDbManager_DbMeta = this.DbMeta
End Property

Private Property Get IDbManager_Connection() As IDbConnection
    Set IDbManager_Connection = this.Connection
End Property

Private Function IDbManager_Command() As IDbCommand
    Set IDbManager_Command = this.CommandFactory.CreateInstance(this.Connection)
End Function

Private Function IDbManager_Recordset( _
            Optional ByVal Disconnected As Boolean = True, _
            Optional ByVal CacheSize As Long = 10, _
            Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
            Optional ByVal LockType As ADODB.LockTypeEnum = adLockReadOnly, _
            Optional ByVal AsyncMode As Boolean = False, _
            Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0) As IDbRecordset
    Set IDbManager_Recordset = DbRecordset.Create( _
        IDbManager_Command, Disconnected, CacheSize, CursorType, LockType, AsyncMode, AsyncOption)
End Function

Private Sub IDbManager_Begin()
    this.Connection.BeginTransaction
End Sub

Private Sub IDbManager_Commit()
    this.Connection.CommitTransaction
End Sub

'@Ignore ProcedureNotUsed
Private Sub IDbManager_Rollback()
    this.Connection.RollbackTransaction
End Sub
