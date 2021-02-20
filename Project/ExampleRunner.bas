Attribute VB_Name = "ExampleRunner"
'@Folder "Forms.Example"
Option Explicit

Public Sub ExampleMacro()
    Dim proxy As WorkbookProxy
    Set proxy = New WorkbookProxy
    With New ExamplePresenter
        .Show proxy
    End With
End Sub
