Attribute VB_Name = "Module1"
Option Explicit

Public Sub ExampleMacro()
    Dim proxy As WorkbookProxy
    Set proxy = New WorkbookProxy
    With New ExamplePresenter
        .Show proxy
    End With
End Sub
