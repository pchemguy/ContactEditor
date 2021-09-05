VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoParameterProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for converting a value of a mapped data type to an ADODB.Parameter object."
'@Folder "SecureADODB.DbParameterProvider"
'@ModuleDescription("An object responsible for converting a value of a mapped data type to an ADODB.Parameter object.")
'@PredeclaredId
'@IgnoreModule IllegalAnnotation
'@Exposed
Option Explicit

Implements IParameterProvider

Private Mappings As ITypeMap


'@Description("Creates a new parameter provider using the specified type mappings.")
Public Function Create(ByVal TypeMap As ITypeMap) As IParameterProvider
Attribute Create.VB_Description = "Creates a new parameter provider using the specified type mappings."
    Dim Instance As AdoParameterProvider
    Set Instance = New AdoParameterProvider
    Instance.Init TypeMap
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal TypeMap As ITypeMap)
Attribute Init.VB_Description = "Default constructor"
    Guard.NullReference TypeMap
    Set Mappings = TypeMap
End Sub


'@Ignore ProcedureNotUsed
Public Property Get Self() As IParameterProvider
    Set Self = Me
End Property


'@Description("Returns class reference.")
'@Ignore ProcedureNotUsed
Public Property Get Class() As AdoParameterProvider
Attribute Class.VB_Description = "Returns class reference."
    Set Class = AdoParameterProvider
End Property


Private Property Get IParameterProvider_Self() As IParameterProvider
    Set IParameterProvider_Self = Me
End Property


'@Description("Returns class reference.")
Private Property Get IParameterProvider_Class() As Object
Attribute IParameterProvider_Class.VB_Description = "Returns class reference."
    Set IParameterProvider_Class = AdoParameterProvider
End Property


Private Sub SetParam(ByVal AdoParam As ADODB.Parameter, _
                     ByVal Value As Variant, _
            Optional ByVal Name As String, _
            Optional ByVal dataTypeName As String)
    Dim localDataTypeName As String
    localDataTypeName = dataTypeName
    If localDataTypeName = vbNullString Then localDataTypeName = TypeName(Value)
    Guard.Expression Mappings.IsMapped(localDataTypeName), _
                     Source:="AdoParameterProvider", _
                     Message:="The data type '" & localDataTypeName & "' has no ADODB.DataTypeEnum mapping."
    
    AdoParam.Type = Mappings.Mapping(localDataTypeName)
    AdoParam.Direction = ADODB.ParameterDirectionEnum.adParamInput
    
    If AdoParam.Type = AdoTypeMappings.DefaultNullMapping Then
        AdoParam.Attributes = ADODB.ParameterAttributesEnum.adParamNullable
        AdoParam.Value = Null
    Else
        AdoParam.Value = Value
    End If
    
    '''' Textual parameters must have positive .Size. The .Size property can
    '''' be set to any value >= Len(Value). However, for vbNullString
    '''' setting .Size to Len(Value) = 0 or not setting it, which means
    '''' leaving it at 0, will fail. (Parameter should be created, but
    '''' an attempt to add it to .Parameters would fail.) "+1" will fix
    '''' the issue for vbNullString and will have no effect on other values.
    If IsCharMapping(AdoParam.Type) Then AdoParam.Size = Len(Value) + 1
    If Name <> vbNullString Then AdoParam.Name = Name
End Sub


Private Function IParameterProvider_FromValues(ByRef Values() As Variant) As VBA.Collection
    Dim AdoParams As VBA.Collection
    Set AdoParams = New VBA.Collection
    Dim ValueIndex As Long
    For ValueIndex = LBound(Values) To UBound(Values)
        AdoParams.Add IParameterProvider_FromValue(Values(ValueIndex))
    Next
    Set IParameterProvider_FromValues = AdoParams
End Function


Private Function IParameterProvider_FromValue(ByVal Value As Variant, _
                                     Optional ByVal Name As String, _
                                     Optional ByVal dataTypeName As String) As ADODB.Parameter
    Dim AdoParam As ADODB.Parameter
    Set AdoParam = New ADODB.Parameter
    SetParam AdoParam, Value, Name, dataTypeName
    Set IParameterProvider_FromValue = AdoParam
End Function


Private Sub IParameterProvider_UpdateFromValues(ByVal AdoCommand As ADODB.Command, ByRef Values() As Variant)
    Guard.NullReference AdoCommand
    
    Dim ParamCount As Long
    ParamCount = AdoCommand.Parameters.Count
    Guard.Expression ParamCount > 0, "AdoParameterProvider", "Parameters collection is empty"
    
    Dim ValueCount As Long
    ValueCount = UBound(Values) - LBound(Values) + 1
    Guard.Expression ValueCount = ParamCount, "AdoParameterProvider", "Parameter/Value counts mismatch"
    
    Dim dataTypeName As String
    Dim Value As Variant
    Dim AdoParam As ADODB.Parameter
    Dim ValueIndex As Long
    Dim ParamIndex As Long
    ParamIndex = 0
    For ValueIndex = LBound(Values) To UBound(Values)
        Set AdoParam = AdoCommand.Parameters(ParamIndex)
        Value = Values(ValueIndex)
        dataTypeName = TypeName(Value)
        Guard.Expression Mappings.IsMapped(dataTypeName), _
                         "AdoParameterProvider", _
                         "The data type '" & dataTypeName & "' has no ADODB.DataTypeEnum mapping."
        Guard.Expression AdoParam.Type = Mappings.Mapping(dataTypeName), _
                         "AdoParameterProvider", _
                         "Parameter/Value type mismatch for parameter: " & CStr(ParamIndex)

        '''' Textual parameters must have positive .Size. The .Size property can
        '''' be set to any value >= Len(Value). However, for vbNullString
        '''' setting .Size to Len(Value) = 0 or not setting it, which means
        '''' leaving it at 0, will fail. (Parameter should be created, but
        '''' an attempt to add it to .Parameters would fail.) "+1" will fix
        '''' the issue for vbNullString and will have no effect on other values.
        If IsCharMapping(AdoParam.Type) Then AdoParam.Size = Len(Value) + 1
        
        If AdoParam.Type = AdoTypeMappings.DefaultNullMapping Then
            AdoParam.Attributes = ADODB.ParameterAttributesEnum.adParamNullable
            AdoParam.Value = Null
        Else
            AdoParam.Value = Value
        End If
        
        ParamIndex = ParamIndex + 1
    Next ValueIndex
End Sub


Private Function IsCharMapping(ByVal Value As ADODB.DataTypeEnum) As Boolean
    IsCharMapping = ((Value = adChar) _
                  Or (Value = adWChar) _
                  Or (Value = adVarChar) _
                  Or (Value = adVarWChar))
End Function
