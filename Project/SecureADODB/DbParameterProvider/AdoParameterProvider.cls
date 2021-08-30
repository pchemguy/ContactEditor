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

Private mappings As ITypeMap


'''' @Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
'@Description("Creates a new parameter provider using the specified type mappings.")
Public Function Create(ByVal typeMap As ITypeMap) As IParameterProvider
Attribute Create.VB_Description = "Creates a new parameter provider using the specified type mappings."
    Dim Instance As AdoParameterProvider
    Set Instance = New AdoParameterProvider
    Instance.Init typeMap
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal typeMap As ITypeMap)
Attribute Init.VB_Description = "Default constructor"
    Guard.NullReference typeMap
    Set mappings = typeMap
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


Private Function IParameterProvider_FromValues(ByRef values() As Variant) As VBA.Collection
    Dim results As VBA.Collection
    Set results = New VBA.Collection
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = LBound(values) To UBound(values)
        results.Add IParameterProvider_FromValue(values(i))
    Next
    Set IParameterProvider_FromValues = results
End Function


Private Function IParameterProvider_FromValue(ByVal Value As Variant, _
                                     Optional ByVal Name As String, _
                                     Optional ByVal dataTypeName As String) As ADODB.Parameter
    If TypeOf Value Is ADODB.Parameter Then
        Set IParameterProvider_FromValue = Value
        Exit Function
    End If

    Dim localDataTypeName As String
    localDataTypeName = dataTypeName
    If localDataTypeName = vbNullString Then localDataTypeName = TypeName(Value)
    Guard.Expression mappings.IsMapped(localDataTypeName), _
                     Source:="AdoParaeterProvider", _
                     Message:="The data type '" & localDataTypeName & "' has no ADODB.DataTypeEnum mapping."
    
    Dim Result As ADODB.Parameter
    Set Result = New ADODB.Parameter
    Result.Type = mappings.Mapping(localDataTypeName)
    Result.Direction = ADODB.ParameterDirectionEnum.adParamInput
    
    If Result.Type = AdoTypeMappings.DefaultNullMapping Then
        Result.Attributes = ADODB.ParameterAttributesEnum.adParamNullable
        Result.Value = Null
    Else
        Result.Value = Value
    End If
    
    If IsCharMapping(Result.Type) And Len(Value) > 0 Then Result.Size = Len(Value)
    If Name <> vbNullString Then Result.Name = Name
    
    Set IParameterProvider_FromValue = Result
End Function


Private Function IsCharMapping(ByVal Value As ADODB.DataTypeEnum) As Boolean
    IsCharMapping = ((Value = adChar) _
                  Or (Value = adWChar) _
                  Or (Value = adVarChar) _
                  Or (Value = adVarWChar))
End Function
