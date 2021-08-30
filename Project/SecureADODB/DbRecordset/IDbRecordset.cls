VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object that wraps an ADODB recordset."
'@IgnoreModule FunctionReturnValueDiscarded
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription("Represents an object that wraps an ADODB recordset.")
'@Interface
'@Exposed
Option Explicit


Public Property Get Self() As IDbRecordset
End Property

Public Property Get Class() As DbRecordset
End Property

'@Description("Gets the wrapped ADODB recordset.")
Public Function AdoRecordset(ByVal SQL As String, ParamArray args() As Variant) As ADODB.Recordset
Attribute AdoRecordset.VB_Description = "Gets the wrapped ADODB recordset."
End Function

Public Function OpenRecordset(ByVal SQL As String, ParamArray args() As Variant) As ADODB.Recordset
End Function

Public Function OpenScalar(ByVal SQL As String, ParamArray args() As Variant) As Variant
End Function

Public Sub RecordsetToQT(ByVal OutputRange As Excel.Range)
End Sub
