Attribute VB_Name = "RecordEditorRunner"
'@Folder "Forms.Record Editor"
Option Explicit

Public Sub ExampleMacro()
    With New RecordEditorPresenter
        .Show
    End With
End Sub
