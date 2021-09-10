VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Wraps ADODB.Parameters collection"
'@Folder "SecureADODB.DbParameters"
'@ModuleDescription "Wraps ADODB.Parameters collection"
'@Exposed
Option Explicit

Public Sub FromValues(ByVal AdoCommand As ADODB.Command, ParamArray ADODBParamsValues())
End Sub

Public Function GetSQL(ByVal AdoCommand As ADODB.Command) As String
End Function
