VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbRecordsetFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An abstract factory that creates IDbRecordset objects."
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription "An abstract factory that creates IDbRecordset objects."
'@Interface
'@Exposed
Option Explicit


'@Description("Creates and returns a new DbRecordset.")
Public Function CreateInstance() As IDbRecordset
Attribute CreateInstance.VB_Description = "Creates and returns a new DbRecordset."
End Function
