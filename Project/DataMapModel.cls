VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataMapModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a set of property-value pairs."
'@Folder "Storage.Map.Model"
'@ModuleDescription "Represents a set of property-value pairs."
'@Exposed
Option Explicit


Private Type TDataMapModel
    Properties As Scripting.Dictionary
End Type
Private this As TDataMapModel


Private Sub Class_Initialize()
    Set this.Properties = New Scripting.Dictionary
    this.Properties.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.Properties = Nothing
End Sub

Public Property Get Properties() As Scripting.Dictionary
    Set Properties = this.Properties
End Property
