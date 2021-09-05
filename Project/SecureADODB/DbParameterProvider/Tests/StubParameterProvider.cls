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


Private Function IParameterProvider_FromValue(ByVal Value As Variant, _
                                     Optional ByVal Name As String, _
                                     Optional ByVal DataTypeName As String) As ADODB.Parameter
    Set IParameterProvider_FromValue = New ADODB.Parameter
End Function

Private Function IParameterProvider_FromValues(ByRef Values() As Variant) As Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Set IParameterProvider_FromValues = Result
End Function

Private Sub IParameterProvider_UpdateFromValues(ByVal AdoCommand As ADODB.Command, ByRef Values() As Variant)
    Debug.Print "Stub IParameterProvider_UpdateFromValues in StubParameterProvider"
End Sub
