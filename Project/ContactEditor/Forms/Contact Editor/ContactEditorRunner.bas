Attribute VB_Name = "ContactEditorRunner"
'@Folder "ContactEditor.Forms.Contact Editor"
Option Explicit

'@Ignore MoveFieldCloserToUsage
Private presenter As ContactEditorPresenter


Public Sub RunContactEditor()
    Set presenter = New ContactEditorPresenter
    presenter.Show "ADODB"
End Sub
