Attribute VB_Name = "ContactEditorRunner"
'@Folder("Forms.Contact Editor")
Option Explicit

'@Ignore MoveFieldCloserToUsage
Private presenter As ContactEditorPresenter


Public Sub RunContactEditor()
    Set presenter = New ContactEditorPresenter
    presenter.Show
End Sub
