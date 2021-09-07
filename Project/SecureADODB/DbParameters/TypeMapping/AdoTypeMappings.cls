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
'@Folder "SecureADODB.DbParameters.TypeMapping"
'@ModuleDescription("An object responsible for mapping programmatic data types to ADODB.DataTypeEnum values.")
'@PredeclaredId
'@Exposed
Option Explicit

Implements ITypeMap

Private Type TAdoTypeMappings
    Mapping As Scripting.Dictionary
    DataTypeEnumNames As Scripting.Dictionary
End Type
Private this As TAdoTypeMappings


'@Description("Creates a default mapping for intrinsic data types.")
Public Function Default() As ITypeMap
Attribute Default.VB_Description = "Creates a default mapping for intrinsic data types."
    Dim Instance As AdoTypeMappings
    Set Instance = New AdoTypeMappings
    Instance.InitDefault
    Set Default = Instance
End Function


'@Description("Creates a CSV mapping for intrinsic data types.")
Public Function CSV() As ITypeMap
Attribute CSV.VB_Description = "Creates a CSV mapping for intrinsic data types."
    Dim Instance As AdoTypeMappings
    Set Instance = New AdoTypeMappings
    Instance.InitCSV
    Set CSV = Instance
End Function


Friend Sub InitDefault()
    Set this.Mapping = New Scripting.Dictionary
    With this.Mapping
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
    
    Set this.DataTypeEnumNames = New Scripting.Dictionary
    With this.DataTypeEnumNames
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


'@Description("Adjusts mapping for the CSV backend: adVarWChar -> adVarChar")
Friend Sub InitCSV()
Attribute InitCSV.VB_Description = "Adjusts mapping for the CSV backend: adVarWChar -> adVarChar"
    InitDefault
    '@Ignore IndexedDefaultMemberAccess
    this.Mapping("String") = adVarChar
End Sub


Public Function DataTypeEnumAsText(ByVal DataTypeValue As DataTypeEnum) As String
    '@Ignore IndexedDefaultMemberAccess
    DataTypeEnumAsText = this.DataTypeEnumNames(CStr(DataTypeValue))
End Function

'@Ignore ProcedureNotUsed: false positive for indexed property accessors?
'@Description("Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name.")
Public Property Get Mapping(ByVal Name As String) As ADODB.DataTypeEnum
Attribute Mapping.VB_Description = "Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name."
    Guard.Expression IsMapped(Name), _
                     Source:="AdoTypeMappings", _
                     Message:="Argument 'name' must be the name of an intrinsic data type; data type '" & Name & "' is not supported."
    '@Ignore IndexedDefaultMemberAccess
    Mapping = this.Mapping(Name)
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
    IsMapped = this.Mapping.Exists(Name)
End Property

'@Description("Gets the ADODB.DataTypeEnum value to map NULL values.")
Public Property Get DefaultNullMapping() As ADODB.DataTypeEnum
Attribute DefaultNullMapping.VB_Description = "Gets the ADODB.DataTypeEnum value to map NULL values."
    DefaultNullMapping = adVarChar
End Property

Private Property Get ITypeMap_IsMapped(ByVal Name As String) As Boolean
    ITypeMap_IsMapped = Me.IsMapped(Name)
End Property

Private Property Get ITypeMap_Mapping(ByVal Name As String) As ADODB.DataTypeEnum
    ITypeMap_Mapping = Mapping(Name)
End Property
