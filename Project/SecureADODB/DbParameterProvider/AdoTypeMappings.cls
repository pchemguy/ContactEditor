VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoTypeMappings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for mapping programmatic data types to ADODB.DataTypeEnum values."
'@Folder "SecureADODB.DbParameterProvider"
'@ModuleDescription("An object responsible for mapping programmatic data types to ADODB.DataTypeEnum values.")
'@PredeclaredId
'@Exposed
Option Explicit

Implements ITypeMap

Private mappings As Scripting.Dictionary


'@Description("Creates a default mapping for intrinsic data types.")
Public Function Default() As ITypeMap
Attribute Default.VB_Description = "Creates a default mapping for intrinsic data types."
    Guard.NonDefaultInstance Me
    Set Default = New AdoTypeMappings
End Function

'@Ignore ProcedureNotUsed: false positive for indexed property accessors?
'@Description("Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name.")
Public Property Get Mapping(ByVal Name As String) As ADODB.DataTypeEnum
Attribute Mapping.VB_Description = "Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name."
    Guard.Expression IsMapped(Name), _
                     Source:="AdoTypeMappings", _
                     Message:="Argument 'name' must be the name of an intrinsic data type; data type '" & Name & "' is not supported."
    Mapping = mappings.Item(Name)
End Property

'@Ignore ProcedureNotUsed: false positive for indexed property accessors?
'@Description("Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name.")
Public Property Let Mapping(ByVal Name As String, ByVal Value As ADODB.DataTypeEnum)
Attribute Mapping.VB_Description = "Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name."
    Guard.Expression IsMapped(Name), _
                     Source:="AdoTypeMappings", _
                     Message:="Argument 'name' must be the name of an intrinsic data type; data type '" & Name & "' is not supported."
    mappings.Item(Name) = Value
End Property

'@Description("True if the specified data type name is mapped to an ADODB.DataTypeEnum value.")
Public Property Get IsMapped(ByVal Name As String) As Boolean
Attribute IsMapped.VB_Description = "True if the specified data type name is mapped to an ADODB.DataTypeEnum value."
    IsMapped = mappings.Exists(Name)
End Property

'@Description("Gets the ADODB.DataTypeEnum value to map NULL values.")
Public Property Get DefaultNullMapping() As ADODB.DataTypeEnum
Attribute DefaultNullMapping.VB_Description = "Gets the ADODB.DataTypeEnum value to map NULL values."
    DefaultNullMapping = adEmpty
End Property

Private Sub Class_Initialize()
    Set mappings = New Scripting.Dictionary
    With mappings
        .Add "Boolean", adBoolean
        .Add "Byte", adInteger
        .Add "Currency", adCurrency
        .Add "Date", adDate
        .Add "Double", adDouble
        .Add "Integer", adInteger
        .Add "Long", adInteger
        .Add "Single", adSingle
        .Add "String", adVarWChar
        .Add "Null", AdoTypeMappings.DefaultNullMapping
        .Add "Empty", AdoTypeMappings.DefaultNullMapping
    End With
End Sub

'@Ignore ProcedureNotUsed
Public Property Get Self() As ITypeMap
    Set Self = Me
End Property

'@Description("Returns class reference.")
'@Ignore ProcedureNotUsed
Public Property Get Class() As AdoTypeMappings
Attribute Class.VB_Description = "Returns class reference."
    Set Class = AdoTypeMappings
End Property

Private Property Get ITypeMap_Self() As ITypeMap
    Set ITypeMap_Self = Me
End Property

'@Description("Returns class reference.")
Private Property Get ITypeMap_Class() As Object
Attribute ITypeMap_Class.VB_Description = "Returns class reference."
    Set ITypeMap_Class = AdoTypeMappings
End Property

Private Property Get ITypeMap_IsMapped(ByVal Name As String) As Boolean
    ITypeMap_IsMapped = Me.IsMapped(Name)
End Property

Private Property Get ITypeMap_Mapping(ByVal Name As String) As ADODB.DataTypeEnum
    ITypeMap_Mapping = Me.Mapping(Name)
End Property

Private Property Let ITypeMap_Mapping(ByVal Name As String, ByVal Value As ADODB.DataTypeEnum)
    Me.Mapping(Name) = Value
End Property
