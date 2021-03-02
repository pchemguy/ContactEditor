Attribute VB_Name = "ContactBrowserRunner"
'@Folder "ContactEditor.Forms.Contact Browser"
Option Explicit

'@Ignore MoveFieldCloserToUsage
Private presenter As ContactBrowserPresenter


Public Sub RunContactEditor()
    Set presenter = New ContactBrowserPresenter
    presenter.Show
End Sub
