VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Logging facility using a dictionary object as the database."
'@Folder "Common.Logger"
'@ModuleDescription "Logging facility using a dictionary object as the database."
'@PredeclaredId
'@Exposed
Option Explicit

Implements ILogger

Public Enum LoggerTypeEnum
    logDisabled = -1
    logGlobal = 1
    logPrivate = 2
End Enum

Private Type TLogger
    Database As Scripting.Dictionary
End Type
Private this As TLogger


Private Sub Class_Initialize()
    Set this.Database = New Scripting.Dictionary
    this.Database.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.Database = Nothing
End Sub


'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As Logger
Attribute Class.VB_Description = "Returns class reference"
    Set Class = Logger
End Property


'@Ignore ProcedureNotUsed
Public Property Get LogDatabase() As Scripting.Dictionary
    Set LogDatabase = this.Database
End Property


'@Description("Default class factory")
Public Function Create() As ILogger
Attribute Create.VB_Description = "Default class factory"
    Guard.NonDefaultInstance Me
    
    Dim Result As Logger
    Set Result = New Logger
    Set Create = Result
End Function


'@Description "Adds item to the log database"
Public Sub Log(ByVal LogLine As String, Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
Attribute Log.VB_Description = "Adds item to the log database"
    Dim LogDb As Scripting.Dictionary
    Set LogDb = IIf(LogDatabase Is Nothing, this.Database, LogDatabase)
    
    Dim key As String
    Do
        key = GetTimeStampMs
    Loop While LogDb.Exists(key)
    LogDb.Item(key) = LogLine
End Sub


'@Description "Prints the log database to the immediate"
Public Sub PrintLog(Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
Attribute PrintLog.VB_Description = "Prints the log database to the immediate"
    Dim LogDb As Scripting.Dictionary
    Set LogDb = IIf(LogDatabase Is Nothing, this.Database, LogDatabase)
    
    Dim key As Variant
    For Each key In LogDb.Keys
        Debug.Print key & vbTab & vbTab & LogDb.Item(key)
    Next key
End Sub


Public Sub ClearLog(Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
    Dim LogDb As Scripting.Dictionary
    Set LogDb = IIf(LogDatabase Is Nothing, this.Database, LogDatabase)
    
    LogDb.RemoveAll
End Sub


'===================================================='
'================= ILogger INTERFACE ================'
'===================================================='

Private Property Get ILogger_Self() As ILogger
    Set ILogger_Self = Me
End Property


Private Property Get ILogger_Class() As ILogger
    Set ILogger_Class = Class
End Property


Private Property Get ILogger_LogDatabase() As Scripting.Dictionary
    Set ILogger_LogDatabase = this.Database
End Property


Private Sub ILogger_Log(ByVal LogLine As String, Optional ByVal LogDatabase As Scripting.Dictionary = Nothing)
    Log LogLine, LogDatabase
End Sub


Private Sub ILogger_PrintLog(Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
    PrintLog LogDatabase
End Sub


Private Sub ILogger_ClearLog(Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
    ClearLog LogDatabase
End Sub
