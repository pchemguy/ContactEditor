VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Helper routines for building connection strings"
'@Folder "SecureADODB.DbManager.DbConnectionString"
'@ModuleDescription "Helper routines for building connection strings"
'@PredeclaredId
'@Exposed
''''
'''' The module incorporates routines, which facilitate construction of
'''' connection strings for the ADODB library. Presently, a distinction is made
'''' between file-based databases, such as sqlite, csv, xls, etc., and network
'''' based databases. A file-based database is accessed based on its type and
'''' file pathname.
'''' CreateFileDB/InitFileDB pair is used for file based databases.
''''
Option Explicit
Option Compare Text

Private Type TDbConnectionString
    DbType As String
    DbPath As String
    Options As String
    Driver As String
    ConnectionString As String
End Type
Private this As TDbConnectionString


'''' Factory for file-based databases.
''''
'''' Args:
''''   FileType (string, optional, "sqlite"):
''''     Type of the database: "sqlite", "csv", "xls"
''''
''''   FileName (string, optional, ""):
''''     Database file name. If not provided, ThisWorkbook.VBProject.Name
''''     will be used. Extension is added based on the database type:
''''       "sqlite" - "db" or "sqlite"
''''       "csv"    - "xsv" or "csv"
''''       "xls"    - "xls
''''
''''   Driver (variant, optional, Empty):
''''     Database driver. If Empty, default values is selected based on
''''     the database type (see constructor code for details).
''''
''''   ConnectionOptions (variant, optional, Empty):
''''     Connection options. See above for details.
''''
''''  N.B.: "xls" backend support is not currently implemented
''''
'''' Examples:
''''   >>> ?DbConnectionString.CreateFileDB("sqlite").ConnectionString
''''   "Driver=SQLite3 ODBC Driver;Database=<Thisworkbook.Path>\SecureADODB.db;SyncPragma=NORMAL;FKSupport=True;"
''''
''''   >>> ?DbConnectionString.CreateFileDB("sqlite").QTConnectionString
''''   "OLEDB;Driver=SQLite3 ODBC Driver;Database=<Thisworkbook.Path>\SecureADODB.db;SyncPragma=NORMAL;FKSupport=True;"
''''
''''   >>> ?DbConnectionString.CreateFileDB("csv").ConnectionString
''''   "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=<Thisworkbook.Path>;"
''''
''''   >>> ?DbConnectionString.CreateFileDB("xls").ConnectionString
''''   NotImplementedErr
''''
'@Description "Factory for file-based databases"
Public Function CreateFileDb(Optional ByVal FileType As String = "sqlite", _
                             Optional ByVal FileName As String = vbNullString, _
                             Optional ByVal Driver As Variant = Empty, _
                             Optional ByVal ConnectionOptions As Variant = Empty _
                            ) As DbConnectionString
Attribute CreateFileDb.VB_Description = "Factory for file-based databases"
    Dim Instance As DbConnectionString
    Set Instance = New DbConnectionString
    Instance.InitFileDB FileType, FileName, Driver, ConnectionOptions
    Set CreateFileDb = Instance
End Function


'@Description "Constructor for file-based databases"
Friend Sub InitFileDB(Optional ByVal FileType As String = "sqlite", _
                      Optional ByVal FileName As String = vbNullString, _
                      Optional ByVal Driver As String = vbNullString, _
                      Optional ByVal ConnectionOptions As Variant = Empty)
Attribute InitFileDB.VB_Description = "Constructor for file-based databases"
    With this
        .DbType = LCase$(FileType)
        .Driver = Driver
        .Options = ConnectionOptions
        Select Case .DbType
            Case "sqlite"
                If Len(Driver) = 0 Then
                    .Driver = "SQLite3 ODBC Driver"
                End If
                If IsEmpty(ConnectionOptions) Then
                    .Options = "SyncPragma=NORMAL;FKSupport=True;"
                End If
                .DbPath = VerifyOrGetDefaultPath(FileName, Array("db", "sqlite"))
                .ConnectionString = "Driver=" + .Driver + ";" + _
                                    "Database=" + .DbPath + ";" + _
                                    .Options
            Case "worksheet", "wsheet"
                .DbPath = VerifyOrGetDefaultPath(FileName, Array("xls"))
                .ConnectionString = .DbPath
            Case "csv"
                .DbPath = VerifyOrGetDefaultPath(FileName, Array("xsv", "csv"))
                Dim DbFileName As String
                DbFileName = Dir$(.DbPath, vbArchive + vbNormal + vbHidden + vbReadOnly + vbSystem)
                .DbPath = Left$(.DbPath, Len(.DbPath) - Len(DbFileName) - 1)
                #If Win64 Then
                    .Driver = "Microsoft Access Text Driver (*.txt, *.csv)"
                #Else
                    .Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
                #End If
                .ConnectionString = "Driver=" + .Driver + ";" + _
                                    "DefaultDir=" + .DbPath + ";"
            Case Else
                .ConnectionString = vbNullString
        End Select
    End With

    If this.ConnectionString = vbNullString Then
        Dim errorDetails As TError
        With errorDetails
            .Number = ErrNo.NotImplementedErr
            .Name = "NotImplementedErr"
            .Source = "DbConnectionString"
            .Description = "Unsupported backend: " & FileType
            .Message = .Description
        End With
        RaiseError errorDetails
    End If
End Sub


Public Property Get ConnectionString() As String
    ConnectionString = this.ConnectionString
End Property


Public Property Get QTConnectionString() As String
    QTConnectionString = "OLEDB;" & this.ConnectionString
End Property
