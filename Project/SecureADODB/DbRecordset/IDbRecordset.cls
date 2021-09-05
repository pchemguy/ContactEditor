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
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription "Represents an object that wraps an ADODB recordset."
'@Interface
'@Exposed
'@IgnoreModule FunctionReturnValueDiscarded
Option Explicit


Public Property Get Self() As IDbRecordset
End Property

Public Property Get Class() As DbRecordset
End Property

Public Property Get AdoRecordset() As ADODB.Recordset
End Property

'@Description("Sets the source of the wrapped ADODB recordset.")
Public Function InitRecordset(ByVal SQL As String, ParamArray args() As Variant) As ADODB.Recordset
Attribute InitRecordset.VB_Description = "Sets the source of the wrapped ADODB recordset."
End Function

Public Function OpenRecordset(ByVal SQL As String, ParamArray args() As Variant) As ADODB.Recordset
End Function

Public Function OpenScalar(ByVal SQL As String, ParamArray args() As Variant) As Variant
End Function

Public Sub RecordsetToQT(ByVal OutputRange As Excel.Range)
End Sub
