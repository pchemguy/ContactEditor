Attribute VB_Name = "ProjectUtilsSnippets"
'@Folder "Common.Project Utils"
'@IgnoreModule ProcedureNotUsed, EmptyStringLiteral
Option Explicit


Private Sub ReferencesSaveToFile()
    Dim Project As ProjectUtils
    Set Project = New ProjectUtils
    Project.ReferencesSaveToFile
End Sub

Private Sub ReferencesAddFromFile()
    Dim Project As ProjectUtils
    Set Project = New ProjectUtils
    Project.ReferencesAddFromFile
End Sub

Private Sub ProjectStructureParse()
    Dim Project As ProjectUtils
    Set Project = New ProjectUtils
    Project.ProjectStructureParse
End Sub

Private Sub ProjectStructureExport()
    Dim Project As ProjectUtils
    Set Project = New ProjectUtils
    Dim ExportFolder As String
    ExportFolder = "" '"Storage\Record"
    Project.ProjectStructureExport ExportFolder
End Sub

Private Sub ProjectFilesExport()
    Dim Project As ProjectUtils
    Set Project = New ProjectUtils
    Dim ExportFolder As String
    ExportFolder = "" '"Storage\Record"
    Project.ProjectFilesExport ExportFolder
End Sub

Private Sub WalkTree()
    Dim Project As ProjectUtils
    Set Project = New ProjectUtils
    Dim ImportFolder As String
    ImportFolder = "" '"Storage\Record"
    Project.WalkTree ImportFolder
End Sub

Private Sub CollectFiles()
    Dim Project As ProjectUtils
    Set Project = New ProjectUtils
    Dim ImportFolder As String
    ImportFolder = "" '"Storage\Record"
    Project.CollectFiles ImportFolder
End Sub

Private Sub ProjectFilesImport()
    Dim Project As ProjectUtils
    Set Project = New ProjectUtils
    Dim ImportFolder As String
    ImportFolder = "" '"Storage\Record"
    Project.ProjectFilesImport ImportFolder
End Sub
