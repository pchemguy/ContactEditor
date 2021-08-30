VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubParameterProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IParameterProvider implementation."
'@Folder "SecureADODB.DbParameterProvider.Tests"
'@ModuleDescription("A stub acting as a IParameterProvider implementation.")
Option Explicit

Implements IParameterProvider


Private Function IParameterProvider_FromValue(ByVal Value As Variant, Optional ByVal Name As String, Optional ByVal dataTypeName As String) As ADODB.Parameter
    Set IParameterProvider_FromValue = New ADODB.Parameter
End Function

Private Function IParameterProvider_FromValues(values() As Variant) As Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Set IParameterProvider_FromValues = Result
End Function


Private Property Get IParameterProvider_Self() As IParameterProvider
    Set IParameterProvider_Self = Me
End Property


'@Description("Returns class reference.")
Private Property Get IParameterProvider_Class() As Object
Attribute IParameterProvider_Class.VB_Description = "Returns class reference."
    Set IParameterProvider_Class = AdoParameterProvider
End Property
