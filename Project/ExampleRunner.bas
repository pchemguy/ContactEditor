Attribute VB_Name = "ExampleRunner"
'@Folder "Forms.Example"
Option Explicit

Public Sub ExampleMacro()
    Dim proxy As StorageManager
    Set proxy = New StorageManager
    With New ExamplePresenter
        .Show proxy
    End With
End Sub
