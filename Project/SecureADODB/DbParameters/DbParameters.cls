VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Wraps ADODB.Parameters collection"
'@Folder "SecureADODB.DbParameters"
'@ModuleDescription "Wraps ADODB.Parameters collection"
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDbParameters

Private Type TDbParameters
    AdoCommand As ADODB.Command
    TypeMap As ITypeMap
End Type
Private this As TDbParameters

Private Type TAdoParam
    Name As String
    DataType As ADODB.DataTypeEnum
    Direction As ADODB.ParameterDirectionEnum
    Size As Long
    Value As Variant
End Type


'''' This permits running DebugFromValue on the predeclared instance
Private Sub Class_Initialize()
    Set this.TypeMap = AdoTypeMappings.Default
End Sub


Public Function Create(ByVal AdoCommand As ADODB.Command, _
              Optional ByVal TypeMap As ITypeMap = Nothing) As IDbParameters
    Dim Instance As DbParameters
    Set Instance = New DbParameters
    Instance.Init AdoCommand, TypeMap
    Set Create = Instance
End Function


Friend Sub Init(ByVal AdoCommand As ADODB.Command, _
       Optional ByVal TypeMap As ITypeMap = Nothing)
    Guard.NullReference AdoCommand
    
    With this
        Set .AdoCommand = AdoCommand
        '''' Default is set in Class_Initialize
        If Not TypeMap Is Nothing Then Set .TypeMap = TypeMap
    End With
End Sub


'@Description "Takes parameter value, validates, and returns properties for ADODB.Parameter"
Private Function FromValue(ByVal Value As Variant, _
                  Optional ByVal Name As String = vbNullString, _
                  Optional ByVal DataType As String = vbNullString) As TAdoParam
Attribute FromValue.VB_Description = "Takes parameter value, validates, and returns properties for ADODB.Parameter"
    Dim AdoParam As TAdoParam
    
    Dim DataTypeName As String
    DataTypeName = IIf(DataType <> vbNullString, DataType, TypeName(Value))
    Guard.Expression this.TypeMap.IsMapped(DataTypeName), _
                     Source:="DbParameters", _
                     Message:="The data type '" & DataType _
                              & "' has no ADODB.DataTypeEnum mapping."
    With AdoParam
        .DataType = this.TypeMap.Mapping(DataTypeName)
        .Direction = ADODB.ParameterDirectionEnum.adParamInput
        If AdoTypeMappings.IsCharMapping(.DataType) Then
            '''' For vbNullString and Null set .Size = 1
            .Size = IIf(IsNull(Value), 1, Len(Value) + 1)
        End If
        .Value = IIf(IsEmpty(.DataType), Null, Value)
        .Name = Name
    End With
    
    FromValue = AdoParam
End Function


'@Description "Creates ADODB.Parameter from prepared TAdoParam structure"
Friend Function CreateParameter(ByVal Value As Variant, _
                       Optional ByVal Name As String = vbNullString, _
                       Optional ByVal DataTypeName As String = vbNullString _
                ) As ADODB.Parameter
Attribute CreateParameter.VB_Description = "Creates ADODB.Parameter from prepared TAdoParam structure"
    Dim AdoParam As TAdoParam
    AdoParam = FromValue(Value, Name, DataTypeName)
    
    With AdoParam
        Set CreateParameter = this.AdoCommand.CreateParameter( _
            .Name, .DataType, .Direction, .Size, .Value)
    End With
End Function


'@Description "Validates ValueCount, ParamCount (if >0), and PlaceHolderCount (in SQL)"
Friend Function ValidateParameterValues(ParamArray ValuesArgs()) As Boolean
Attribute ValidateParameterValues.VB_Description = "Validates ValueCount, ParamCount (if >0), and PlaceHolderCount (in SQL)"
    Dim Values() As Variant
    Values = UnfoldParamArray(ValuesArgs)

    '''' Debug.Assert ValueCount > 0
    Dim ValueCount As Long
    ValueCount = UBound(Values) - LBound(Values) + 1
    Guard.Expression ValueCount > 0, "DbParameters", "Invalid parameter array"
    
    '''' Debug.Assert PlaceHolderCount = ValueCount
    Dim PlaceHolderCount As Long
    Dim SQLQuery As String
    SQLQuery = this.AdoCommand.CommandText
    If Len(SQLQuery) > 0 Then
        PlaceHolderCount = Len(SQLQuery) - Len(Replace(SQLQuery, "?", vbNullString))
        Guard.Expression PlaceHolderCount = ValueCount, _
                         "DbParameters", _
                         "Number of <?> placeholders does not match parameter value count"
    Else
        Debug.Print "WARNING: AdoCommand.CommandText is not set, skipping check"
    End If

    '''' Debug.Assert ParameterCount = 0 Or ParameterCount = ValueCount
    Dim ParameterCount As Long
    ParameterCount = this.AdoCommand.Parameters.Count
    Guard.Expression ParameterCount = 0 Or ParameterCount = ValueCount, _
                     "DbParameters", _
                     "AdoCommand.Parameters.Count does not match parameter value count"
    
    ValidateParameterValues = True
End Function


'@Description "Permits testing FromValue without moving TAdoParam to a regular module"
Friend Function DebugFromValue(ByVal Value As Variant, _
                      Optional ByVal Name As String = vbNullString, _
                      Optional ByVal DataType As String = vbNullString) As Scripting.Dictionary
Attribute DebugFromValue.VB_Description = "Permits testing FromValue without moving TAdoParam to a regular module"
    Dim AdoParam As TAdoParam
    AdoParam = FromValue(Value, Name, DataType)
    Dim Props As Scripting.Dictionary
    Set Props = New Scripting.Dictionary
    Props.CompareMode = TextCompare
    With Props
        .Item("Name") = AdoParam.Name
        .Item("Type") = AdoParam.DataType
        .Item("Direction") = AdoParam.Direction
        .Item("Size") = AdoParam.Size
        .Item("Value") = AdoParam.Value
    End With
    Set DebugFromValue = Props
End Function


'@Ignore ProcedureNotUsed: For use in immediate
'@Description "Permits testing FromValue without moving TAdoParam to a regular module"
Friend Function DebugFromValuePrintable( _
                         ByVal Value As Variant, _
                Optional ByVal Name As String = vbNullString, _
                Optional ByVal DataType As String = vbNullString) As Variant
Attribute DebugFromValuePrintable.VB_Description = "Permits testing FromValue without moving TAdoParam to a regular module"
    Dim AdoParam As TAdoParam
    AdoParam = FromValue(Value, Name, DataType)
    
    Dim ValueStr As String
    With AdoParam
        If IsEmpty(.Value) Then
            ValueStr = "<Empty>"
        ElseIf IsNull(.Value) Then
            ValueStr = "<Null>"
        Else
            ValueStr = CStr(.Value)
        End If
        
        DebugFromValuePrintable = Array( _
            "Name:  " & .Name, _
            "Type:  " & AdoTypeMappings.DataTypeEnumAsText(.DataType), _
            "I/O:   " & .Direction, _
            "Size:  " & .Size, _
            "Value: " & ValueStr _
        )
    End With
End Function


'@Description "Creates or updates ADODB.Parameters collection in the Command object from an array of values"
Private Sub IDbParameters_FromValues(ParamArray ValuesArgs())
Attribute IDbParameters_FromValues.VB_Description = "Creates or updates ADODB.Parameters collection in the Command object from an array of values"
    Dim Values() As Variant
    Values = UnfoldParamArray(ValuesArgs)
    
    '@Ignore FunctionReturnValueDiscarded: Raises error on failure
    ValidateParameterValues Values

    Dim ParameterCount As Long
    ParameterCount = this.AdoCommand.Parameters.Count
    Dim UpdateParams As Boolean
    UpdateParams = ParameterCount > 0
    
    Dim AdoParam As TAdoParam
    Dim Param As ADODB.Parameter
    Dim ValueIndex As Long
    Dim ParameterIndex As Long
    If Not UpdateParams Then
        For ValueIndex = LBound(Values) To UBound(Values)
            this.AdoCommand.Parameters.Append CreateParameter(Values(ValueIndex))
        Next ValueIndex
    Else
        ParameterIndex = 0
        For ValueIndex = LBound(Values) To UBound(Values)
            AdoParam = FromValue(Values(ValueIndex))
            '@Ignore IndexedDefaultMemberAccess
            Set Param = this.AdoCommand.Parameters(ParameterIndex)
            With AdoParam
                Param.Type = .DataType
                Param.Size = .Size
                Param.Value = .Value
            End With
            ParameterIndex = ParameterIndex + 1
        Next ValueIndex
    End If
End Sub
