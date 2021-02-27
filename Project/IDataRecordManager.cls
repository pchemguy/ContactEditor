VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataRecordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A composition of a data model and a storage class responsible for loading/saving the model data."
'@Folder "Storage.Record.Manager"
'@ModuleDescription "A composition of a data model and a storage class responsible for loading/saving the model data."
'@Interface
'@Exposed
Option Explicit


Public Property Get Model() As DataRecordModel
End Property

Public Sub LoadDataIntoModel()
End Sub

Public Sub SaveDataFromModel()
End Sub
