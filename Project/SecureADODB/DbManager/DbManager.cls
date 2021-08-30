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
'@ModuleDescription("An object that encapsulates a database transaction.")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDbManager

Private Type TDbManager
    DbMeta As DbMetaData
    DbConnStr As DbConnectionString
    Connection As IDbConnection
    CommandFactory As IDbCommandFactory
    UseTransactions As Boolean
    LogController As ILogger
End Type
Private this As TDbManager


'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As DbManager
Attribute Class.VB_Description = "Returns class reference"
    Set Class = DbManager
End Property


'@Description("Default factory")
Public Function Create(ByVal db As IDbConnection, _
                       ByVal factory As IDbCommandFactory, _
              Optional ByVal UseTransactions As Boolean = True, _
              Optional ByVal LogController As ILogger = Nothing) As IDbManager
Attribute Create.VB_Description = "Default factory"
    Dim Instance As DbManager
    Set Instance = New DbManager
    Instance.Init db, factory, UseTransactions, LogController
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal db As IDbConnection, _
                ByVal factory As IDbCommandFactory, _
       Optional ByVal UseTransactions As Boolean = True, _
       Optional ByVal LogController As ILogger = Nothing)
Attribute Init.VB_Description = "Default constructor"
              
    Guard.NullReference factory
    Guard.NullReference db
    Guard.Expression db.State = adStateOpen, Source:="DbManager", Message:="Connection should be open."

    Set this.LogController = LogController
    Set this.Connection = db
    Set this.CommandFactory = factory
    this.UseTransactions = UseTransactions
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
        Optional ByVal UseTransactions As Boolean = True, _
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
    
    Dim provider As IParameterProvider
    Set provider = AdoParameterProvider.Create(AdoTypeMappings.Default)
    
    Dim baseCommand As IDbCommandBase
    Set baseCommand = DbCommandBase.Create(provider)
    
    Dim factory As IDbCommandFactory
    Set factory = DbCommandFactory.Create(baseCommand)
    
    Dim DbConnStr As DbConnectionString
    Set DbConnStr = DbConnectionString.CreateFileDb(DbType, DbFileName, , ConnectionOptions)
    Dim db As IDbConnection
    Set db = DbConnection.Create(DbConnStr.ConnectionString, LogController)
    
    Dim Instance As DbManager
    Set Instance = DbManager.Create(db, factory, UseTransactions, LogController)
    Set Instance.DbConnStr = DbConnStr
    Set Instance.DbMeta = DbMetaData.Create(DbConnStr)
    Set CreateFileDb = Instance
End Function


'@Ignore ProcedureNotUsed
Public Property Get DbConnStr() As DbConnectionString
    Set DbConnStr = this.DbConnStr
End Property


Public Property Set DbConnStr(ByVal Instance As DbConnectionString)
    Set this.DbConnStr = Instance
End Property


'@Ignore ProcedureNotUsed
Public Property Get DbMeta() As DbMetaData
    Set DbMeta = this.DbMeta
End Property


Public Property Set DbMeta(ByVal Instance As DbMetaData)
    Set this.DbMeta = Instance
End Property


'@Description("Returns class reference")
Public Property Get IDbManager_Class() As DbManager
Attribute IDbManager_Class.VB_Description = "Returns class reference"
    Set IDbManager_Class = DbManager
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
        Optional ByVal Scalar As Boolean = False, _
        Optional ByVal Disconnected As Boolean = True, _
        Optional ByVal CacheSize As Long = 10, _
        Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
        Optional ByVal AsyncMode As Boolean = False, _
        Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0) As IDbRecordset
    Dim cmd As IDbCommand
    Set cmd = this.CommandFactory.CreateInstance(this.Connection)
    Set IDbManager_Recordset = DbRecordset.Create( _
        cmd, Scalar, Disconnected, CacheSize, CursorType, AsyncMode, AsyncOption)
End Function


Private Sub IDbManager_Begin()
    Guard.Expression this.UseTransactions, Source:="DbManager", Message:="Transactions are disabled by the caller."
    this.Connection.BeginTransaction
End Sub


Private Sub IDbManager_Commit()
    Guard.Expression this.UseTransactions, Source:="DbManager", Message:="Transactions are disabled by the caller."
    this.Connection.CommitTransaction
End Sub


Private Sub IDbManager_Rollback()
    Guard.Expression this.UseTransactions, Source:="DbManager", Message:="Transactions are disabled by the caller."
    this.Connection.RollbackTransaction
End Sub
