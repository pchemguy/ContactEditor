VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADOlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides ADO related helpers"
'@Folder "ContactEditor.Storage.Table.ADOlib"
'@ModuleDescription "Provides ADO related helpers"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


Public Function Create() As ADOlib
    Dim ADOTool As ADOlib
    Set ADOTool = New ADOlib
    'ADOTool.Init
    Set Create = ADOTool
End Function


Public Function GetSQLiteConnectionString(Optional ByVal DBPath As String = vbNullString, _
                                          Optional ByVal Features As String = vbNullString) As Scripting.Dictionary
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    
    Dim Database As String
    Database = Right$(DBPath, Abs(Len(DBPath) - 7))
    Database = VerifyOrGetDefaultPath(Database, Array("db", "sqlite"))
    
    Dim Options As String
    Options = IIf(Len(Features) > 0, Features, "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;")
    
    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = New Scripting.Dictionary
    ConnectionString.CompareMode = TextCompare
    
    ConnectionString("ADO") = "Driver=" & Driver & ";" & "Database=" & Database & ";" & Options
    ConnectionString("QT") = "OLEDB;" + ConnectionString("ADO")
    
    Set GetSQLiteConnectionString = ConnectionString
End Function

 
' Takes and array of field names and types (ADODB.DataTypeEnum) collected by the
' DataTableADODB.CollectTableMetadata routine.
' Returns an array of ADODB.Parameter.
' If AdoParams is provided, the collection is emptied and populated with created parameters.
'
' LIMITATIONS:
' The first field in FieldNames/FieldTypes is assumed to be primary key, which is added last,
' as it is expected to be used in the WHERE clause.
'
'@Description "Prepares ADODB.Parameters or array of ADODB.Parameter for record update."
Public Function SetAdoParamsForRecordUpdate(ByVal FieldNames As Variant, _
                                            ByVal FieldTypes As Variant, _
                                   Optional ByVal AdoParamsCol As ADODB.Parameters = Nothing) As Variant
Attribute SetAdoParamsForRecordUpdate.VB_Description = "Prepares ADODB.Parameters or array of ADODB.Parameter for record update."
    '''' Make sure that both FieldNames and FieldTypes are arrays and have the same base and size
    Dim TypeCheck As Boolean
    TypeCheck = ((VarType(FieldNames) And vbArray) = vbArray)
    Guard.Expression TypeCheck, "ADOlib", "FieldNames is not array"
    TypeCheck = ((VarType(FieldTypes) And vbArray) = vbArray)
    Guard.Expression TypeCheck, "ADOlib", "FieldTypes is not array"
    TypeCheck = (VarType(FieldNames(LBound(FieldNames))) = vbString)
    Guard.Expression TypeCheck, "ADOlib", "FieldNames must be strings"
    Guard.Expression LBound(FieldNames) = LBound(FieldTypes), "ADOlib", "FieldNames and FieldTypes arrays base mismatch"
    Guard.Expression UBound(FieldNames) = UBound(FieldTypes), "ADOlib", "FieldNames and FieldTypes arrays size mismatch"
    Guard.Expression LBound(FieldNames) = 1, "ADOlib", "FieldNames and FieldTypes arrays should be 1-based"
    
    Dim FieldCount As Long: FieldCount = UBound(FieldNames) - LBound(FieldNames) + 1
    Dim AdoParamsArray() As ADODB.Parameter: ReDim AdoParamsArray(1 To FieldCount)
    Dim AdoParam As ADODB.Parameter
    Dim FieldIndex As Long
    ' The first field is assumed to be the primary key, should be added last
    For FieldIndex = 2 To FieldCount
        Set AdoParam = New ADODB.Parameter
        With AdoParam
            .Name = FieldNames(FieldIndex)
            .Type = FieldTypes(FieldIndex)
            ' Value and size are required for ADODB.Command.Parameters.Append
            ' After Type is set (above), zero is automatically cast by Parameter.Value
            ' setter as necessary, so there is no need to worry about matching Value
            ' and Type.
            .Value = 0
            .Size = Len(.Value)
        End With
        Set AdoParamsArray(FieldIndex - 1) = AdoParam
    Next FieldIndex
    ' Add the primary key
    Set AdoParam = New ADODB.Parameter
    AdoParam.Name = FieldNames(LBound(FieldNames))
    AdoParam.Type = FieldTypes(LBound(FieldNames))
    Set AdoParamsArray(FieldCount) = AdoParam
    
    ' If AdoParamsCol is supplied, empty it and populate from AdoParamsArray
    If Not AdoParamsCol Is Nothing Then
        Dim ParamIndex As Long
        For ParamIndex = 1 To AdoParamsCol.Count
            AdoParamsCol.Delete ParamIndex
        Next ParamIndex
        For FieldIndex = 1 To FieldCount
            AdoParamsCol.Append AdoParamsArray(FieldIndex)
        Next FieldIndex
    End If
    SetAdoParamsForRecordUpdate = AdoParamsArray
End Function


Public Sub RecordValuesToAdoParams(ByVal RecordValues As Variant, ByVal AdoParams As Variant)
    Dim TypeCheck As Boolean
    TypeCheck = TypeOf AdoParams Is ADODB.Parameters
    If Not TypeCheck Then
        TypeCheck = (VarType(AdoParams) = vbArray + vbObject) And (TypeOf AdoParams(1) Is ADODB.Parameter)
    End If
    Guard.Expression TypeCheck, "ADOlib", "AdoParams should be either ADODB.Parameters or an array of ADODB.Parameter's"
    TypeCheck = (VarType(RecordValues) And vbArray = vbArray)
    Guard.Expression TypeCheck, "ADOlib", "RecordValues should be an array"
    
    Dim ParamCount As Long
    Dim ParamMinIndex As Long
    TypeCheck = (VarType(AdoParams) = vbArray + vbObject)
    If TypeCheck Then ' AdoParams - array of ADODB.Parameter's
        ParamMinIndex = LBound(AdoParams)
        ParamCount = UBound(AdoParams) - LBound(AdoParams) + 1
    Else ' AdoParams - ADODB.Parameters
        ParamMinIndex = 0 ' ADODB.Parameters is 0-based
        ParamCount = AdoParams.Count
    End If
    Dim FieldCount As Long: FieldCount = UBound(RecordValues) - LBound(RecordValues) + 1
    Guard.Expression FieldCount = ParamCount, "ADOlib", "RecordValues and AdoParams - size mismatch"
    
    Dim ParamIndex As Long: ParamIndex = ParamMinIndex
    Dim FieldIndex As Long
    Dim ParamSize As Long
    For FieldIndex = LBound(RecordValues) + 1 To UBound(RecordValues)  ' Skip the first field - primary key
        With AdoParams(ParamIndex)
            ParamSize = Len(RecordValues(FieldIndex))
            .Size = ParamSize ' .Size is not set by ADODB.Parameter.Value setter, and it must be adjusted before setting .Value
            .Value = RecordValues(FieldIndex)
        End With
        ParamIndex = ParamIndex + 1
    Next FieldIndex
    FieldIndex = LBound(RecordValues)
    With AdoParams(ParamIndex)
        ParamSize = Len(RecordValues(FieldIndex))
        .Size = ParamSize
        .Value = RecordValues(FieldIndex)
    End With
End Sub
