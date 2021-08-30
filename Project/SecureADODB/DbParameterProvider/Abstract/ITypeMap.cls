VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITypeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for mapping programmatic data types to ADODB.DataTypeEnum values."
'@Folder "SecureADODB.DbParameterProvider.Abstract"
'@ModuleDescription("Represents an object responsible for mapping programmatic data types to ADODB.DataTypeEnum values.")
'@Interface
'@Exposed
Option Explicit

'@Description("Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name.")
Public Property Get Mapping(ByVal Name As String) As ADODB.DataTypeEnum
Attribute Mapping.VB_Description = "Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name."
End Property

Public Property Let Mapping(ByVal Name As String, ByVal Value As ADODB.DataTypeEnum)
End Property

'@Description("True if the specified data type name is mapped to an ADODB.DataTypeEnum value.")
Public Property Get IsMapped(ByVal Name As String) As Boolean
Attribute IsMapped.VB_Description = "True if the specified data type name is mapped to an ADODB.DataTypeEnum value."
End Property

Public Property Get Self() As ITypeMap
End Property

'@Description("Returns class reference.")
Public Property Get Class() As Object
Attribute Class.VB_Description = "Returns class reference."
End Property
