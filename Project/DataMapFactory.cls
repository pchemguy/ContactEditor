VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataMapFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Storage.Map.Factory"
'@PredeclaredId
Option Explicit
Option Compare Text

Implements IDataMapFactory

Private Type TDataMapFactory
    ClassName As String
End Type
Private this As TDataMapFactory


Public Function Create(ByVal ClassName As String) As IDataMapFactory
    Dim result As DataMapFactory
    Set result = New DataMapFactory
    result.Init ClassName
    Set Create = result
End Function


Public Sub Init(ByVal ClassName As String)
    this.ClassName = ClassName
End Sub


Public Function CreateInstance(ByVal ClassName As String, ByVal model As DataMapModel, ByVal ConnectionString As String) As IDataMap
    Select Case ClassName
        Case "SheetMap"
            Set CreateInstance = SheetMap.Create(model, ConnectionString)
        Case Else
            VBA.Err.Raise vbObjectError, "IDataMapFactory", "Unsupported backend: " & ClassName
    End Select
End Function


Private Function IDataMapFactory_CreateInstance(ByVal model As DataMapModel, ByVal ConnectionString As String) As IDataMap
    IDataMapFactory_CreateInstance = CreateInstance(this.ClassName, model, ConnectionString)
End Function
