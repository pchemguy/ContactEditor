Attribute VB_Name = "RecordEditorRunner"
'@Folder "WorkbookProxyExampleDialog.Record Editor"
Option Explicit

Public Sub ExampleMacro()
    Dim proxy As StorageManager
    Set proxy = New StorageManager
    
    With New RecordEditorPresenter
        .Show proxy
    End With
End Sub
