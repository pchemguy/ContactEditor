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
'@Folder "SecureADODB.DbParameters"
'@ModuleDescription("An object responsible for mapping programmatic data types to ADODB.DataTypeEnum values.")
'@PredeclaredId
'@Exposed
Option Explicit

Implements ITypeMap

Private Mappings As Scripting.Dictionary
Private DataTypeEnumNames As Scripting.Dictionary


Private Sub Class_Initialize()
    Set Mappings = New Scripting.Dictionary
    Mappings.CompareMode = TextCompare
    With Mappings
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
    
    Set DataTypeEnumNames = New Scripting.Dictionary
    DataTypeEnumNames.CompareMode = TextCompare
    With DataTypeEnumNames
        .Item(CStr(adBoolean)) = "adBoolean"
        .Item(CStr(adInteger)) = "adInteger"
        .Item(CStr(adCurrency)) = "adCurrency"
        .Item(CStr(adDate)) = "adDate"
        .Item(CStr(adDouble)) = "adDouble"
        .Item(CStr(adSingle)) = "adSingle"
        .Item(CStr(adVarWChar)) = "adVarWChar"
        .Item(CStr(adVarChar)) = "adVarChar"
    End With
End Sub


Public Function DataTypeEnumAsText(ByVal DataTypeValue As DataTypeEnum) As String
    '@Ignore IndexedDefaultMemberAccess
    DataTypeEnumAsText = DataTypeEnumNames(CStr(DataTypeValue))
End Function


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
    Mapping = Mappings.Item(Name)
End Property

'@Ignore ProcedureNotUsed: false positive for indexed property accessors?
'@Description("Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name.")
Public Property Let Mapping(ByVal Name As String, ByVal Value As ADODB.DataTypeEnum)
Attribute Mapping.VB_Description = "Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name."
    Guard.Expression IsMapped(Name), _
                     Source:="AdoTypeMappings", _
                     Message:="Argument 'name' must be the name of an intrinsic data type; data type '" & Name & "' is not supported."
    Mappings.Item(Name) = Value
End Property

Public Function IsCharMapping(ByVal DataTypeName As ADODB.DataTypeEnum) As Boolean
    IsCharMapping = ((DataTypeName = adChar) _
                  Or (DataTypeName = adWChar) _
                  Or (DataTypeName = adVarChar) _
                  Or (DataTypeName = adVarWChar))
End Function

'@Description("True if the specified data type name is mapped to an ADODB.DataTypeEnum value.")
Public Property Get IsMapped(ByVal Name As String) As Boolean
Attribute IsMapped.VB_Description = "True if the specified data type name is mapped to an ADODB.DataTypeEnum value."
    IsMapped = Mappings.Exists(Name)
End Property

'@Description("Gets the ADODB.DataTypeEnum value to map NULL values.")
Public Property Get DefaultNullMapping() As ADODB.DataTypeEnum
Attribute DefaultNullMapping.VB_Description = "Gets the ADODB.DataTypeEnum value to map NULL values."
    DefaultNullMapping = adVarChar
End Property

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
