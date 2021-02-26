VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRecordModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a set of property-value pairs."
'@Folder "Storage.Record.Model"
'@ModuleDescription "Represents a set of property-value pairs."
'@Exposed
Option Explicit

Private Type TDataRecordModel
    Record As Scripting.Dictionary
End Type
Private this As TDataRecordModel


Private Sub Class_Initialize()
    Set this.Record = New Scripting.Dictionary
    this.Record.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.Record = Nothing
End Sub

Public Property Get Record() As Scripting.Dictionary
    Set Record = this.Record
End Property
