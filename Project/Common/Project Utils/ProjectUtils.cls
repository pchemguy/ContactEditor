VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Common.Project Utils"
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess
Option Explicit

Private Const PROJECT_FOLDER As String = "Project"
Private Const COMMON_FOLDER As String = "Common"
Private Const REFERENCES_FILE As String = "References.xsv"

Private Type TProjectUtils
    Project As VBIDE.VBProject
    ProjectPath As String
    Paths As Scripting.Dictionary
    Files As Scripting.Dictionary
    fso As Scripting.FileSystemObject
    wsh As IWshRuntimeLibrary.WshShell
    EnvVarNames As Variant
End Type
Private this As TProjectUtils


Private Sub Class_Initialize()
    With this
        Select Case Application.Name
            Case "Microsoft Excel"
                Set .Project = Application.ActiveWorkbook.VBProject
            Case "Microsoft Access", "Microsoft Word"
                Set .Project = Application.VBE.ActiveVBProject
            Case "Microsoft PowerPoint"
                '@Ignore IndexedDefaultMemberAccess
                Set .Project = Application.VBE.VBProjects(1)
        End Select
        
        Set .Paths = New Scripting.Dictionary
        .Paths.CompareMode = TextCompare
        Set .Files = New Scripting.Dictionary
        .Files.CompareMode = TextCompare
        Set .fso = New Scripting.FileSystemObject
        Set .wsh = New IWshRuntimeLibrary.WshShell
        .ProjectPath = .fso.GetParentFolderName(.Project.FileName) & Application.PathSeparator & PROJECT_FOLDER
        If Dir$(.ProjectPath, vbDirectory) = vbNullString Then MkDir .ProjectPath
        .EnvVarNames = Array("CommonProgramFiles(x86)", "CommonProgramFiles", "ProgramFiles (x86)", "ProgramFiles", "SystemRoot")
    End With
End Sub


Private Sub Class_Terminate()
    With this
        Set .Project = Nothing
        Set .Paths = Nothing
        Set .Files = Nothing
        Set .fso = Nothing
        Set .wsh = Nothing
    End With
End Sub


'@Description "Creates path, including any non-existing subdirectories."
Private Sub MkPath(ByVal FullPath As String)
Attribute MkPath.VB_Description = "Creates path, including any non-existing subdirectories."
    Dim Folders As Variant
    Folders = Split(FullPath, Application.PathSeparator)
    
    Dim Path As String
    Path = Folders(0)
    If Dir$(Path, vbDirectory) = vbNullString Then Err.Raise 76, "ProjectUtils", "Path not found"
    
    Dim FolderIndex As Long
    For FolderIndex = 1 To UBound(Folders, 1)
        Path = Path & Application.PathSeparator & Folders(FolderIndex)
        If Dir$(Path, vbDirectory) = vbNullString Then MkDir Path
    Next FolderIndex
End Sub


'@Description "Save references info for the Project of the ActiveWorkbook into tsv/csv file."
Public Sub ReferencesSaveToFile()
Attribute ReferencesSaveToFile.VB_Description = "Save references info for the Project of the ActiveWorkbook into tsv/csv file."
    Dim References As VBIDE.References
    Set References = this.Project.References
    Dim ReferenceString As String
    Dim ReferencesArray() As Variant
    ReDim ReferencesArray(0 To References.Count)
    Dim Reference As VBIDE.Reference
    Dim ReferenceIndex As Long: ReferenceIndex = 0
    
    ReferenceString = "Name" & _
              vbTab & "GUID" & _
              vbTab & "Major" & _
              vbTab & "Minor" & _
              vbTab & "FullPath"
    ReferencesArray(ReferenceIndex) = ReferenceString
    ReferenceIndex = ReferenceIndex + 1
    
    Dim EnvVarIndex As Long
    Dim FullPath As String
    Dim FullPathLen As Long
    Dim EnvVarCount As Long: EnvVarCount = UBound(this.EnvVarNames, 1) + 1
    For Each Reference In References
        '''' Replace path prefix with corresponding common environment variable if exists
        FullPath = Reference.FullPath
        FullPathLen = Len(FullPath)
        EnvVarIndex = 0
        Do While (Len(FullPath) = FullPathLen) And (EnvVarIndex < EnvVarCount)
            FullPath = Replace(FullPath, Environ$(this.EnvVarNames(EnvVarIndex)), "%" & this.EnvVarNames(EnvVarIndex) & "%", 1, 1, vbTextCompare)
            EnvVarIndex = EnvVarIndex + 1
        Loop
        If Len(FullPath) = FullPathLen Then
            FullPath = Replace(FullPath, "C:\ProgramFiles", "%ProgramFiles%", 1, 1, vbTextCompare)
        End If
        
        ReferenceString = Reference.Name & _
                  vbTab & Reference.GUID & _
                  vbTab & CStr(Reference.Major) & _
                  vbTab & CStr(Reference.Minor) & _
                  vbTab & FullPath
        ReferencesArray(ReferenceIndex) = ReferenceString
        ReferenceIndex = ReferenceIndex + 1
    Next Reference
    
    '''' Save references to a tab separated file in the project folder
    Dim PathName As String
    PathName = this.ProjectPath & Application.PathSeparator & REFERENCES_FILE
    Dim FileHandle As Long: FileHandle = FreeFile
    Open PathName For Output As #FileHandle
    Print #FileHandle, Join(ReferencesArray, vbNewLine)
    Close #FileHandle
End Sub


'''' Errors during reference addition are ignored (expect the major source due to already
'''' activated references). Alternatively, read activated refernces first and skip addition
'''' of activated references.
'@Description "Add references from the tsv/csv file to the Project of the ActiveWorkbook ."
Public Sub ReferencesAddFromFile(Optional ByVal PathName As String = vbNullString, Optional ByVal UseGUID As Boolean = True)
Attribute ReferencesAddFromFile.VB_Description = "Add references from the tsv/csv file to the Project of the ActiveWorkbook ."
    Dim PathNameL As String
    PathNameL = IIf(PathName <> vbNullString, PathName, this.ProjectPath & Application.PathSeparator & REFERENCES_FILE)
    Dim ReadBuffer As String
    
    Dim FileHandle As Long: FileHandle = FreeFile
    Open PathNameL For Input As #FileHandle
    ReadBuffer = Input$(LOF(FileHandle), #FileHandle)
    Close #FileHandle
    
    Dim refj As Long
    If Len(ReadBuffer) > 0 Then
        '''' Split buffer into record lines
        Dim ReadLines() As String
        ReadLines = Split(ReadBuffer, vbNewLine)
        Dim ReferencesArray() As Variant
        ReDim ReferencesArray(0 To UBound(ReadLines, 1) - 2)
        For refj = 0 To UBound(ReadLines, 1) - 2
            '''' Split records into fields
            ReferencesArray(refj) = Split(ReadLines(refj + 1), vbTab)
            ReferencesArray(refj)(4) = this.wsh.ExpandEnvironmentStrings(ReferencesArray(refj)(4))
        Next refj
    End If
    
    '''' Add all references ignoring errors
    Debug.Print "Adding references from " & PathName & vbNewLine & _
                "For refernces that are already activated, a warning" & vbNewLine & _
                UCase$("Name conflicts with existing module, project, or object library") & vbNewLine & _
                "will be reported. Please ignore it." & vbNewLine
    On Error Resume Next
    If UseGUID Then
        For refj = 0 To UBound(ReferencesArray, 1)
            this.Project.References.AddFromGuid ReferencesArray(refj)(1), CLng(ReferencesArray(refj)(2)), CLng(ReferencesArray(refj)(3))
            If VBA.Information.Err.Number > 0 Then Debug.Print VBA.Information.Err.Description
        Next refj
    Else
        For refj = 0 To UBound(ReferencesArray, 1)
            this.Project.References.AddFromFile ReferencesArray(refj)(4)
            If VBA.Information.Err.Number > 0 Then Debug.Print VBA.Information.Err.Description
        Next refj
    End If
    On Error GoTo 0
End Sub


'@Description "Iterates over VBComponents, extracts @Folder annotation and collects File/Folder hierarchy information."
Public Sub ProjectStructureParse()
Attribute ProjectStructureParse.VB_Description = "Iterates over VBComponents, extracts @Folder annotation and collects File/Folder hierarchy information."
    this.Paths.RemoveAll
    this.Files.RemoveAll
    
    Dim Component As VBComponent
    Dim AnnotateStart As Long
    Dim AnnotateStop As Long
    Dim Path As String
    Dim Ext As String
    Dim ComponentType As String
    Dim ComponentDeclareLines As String

    For Each Component In this.Project.VBComponents
        Select Case Component.Type
            Case vbext_ComponentType.vbext_ct_StdModule
                ComponentType = "Module"
                Ext = ".bas"
            Case vbext_ComponentType.vbext_ct_ClassModule
                ComponentType = "Class"
                Ext = ".cls"
            Case vbext_ComponentType.vbext_ct_MSForm
                ComponentType = "Form"
                Ext = ".frm"
            Case vbext_ComponentType.vbext_ct_Document
                ComponentType = "Document"
                Ext = ".doccls"
        End Select
        
        ComponentDeclareLines = Component.CodeModule.Lines(1, Component.CodeModule.CountOfDeclarationLines)
        AnnotateStart = InStr(1, ComponentDeclareLines, "'@Folder", vbTextCompare)
        If AnnotateStart > 0 Then
            AnnotateStart = InStr(AnnotateStart, ComponentDeclareLines, """")
            AnnotateStop = InStr(AnnotateStart + 1, ComponentDeclareLines, """")
            Path = Mid$(ComponentDeclareLines, AnnotateStart + 1, AnnotateStop - AnnotateStart - 1)
        Else
            Path = COMMON_FOLDER & "." & ComponentType
            Component.CodeModule.InsertLines 1, "'@Folder """ & Path & """"
        End If
        Path = Replace$(Path, ".", Application.PathSeparator)
        this.Paths(Path) = vbNullString
        this.Files(Component.Name) = Array(Ext, ComponentType, Path)
    Next Component
End Sub


'If provided, prefix must use system path separator
'@Description "Recreates project folder structure in the Project folder (for the entire project or just a path matching provided prefix)."
Public Sub ProjectStructureExport(Optional ByVal Prefix As String = vbNullString)
Attribute ProjectStructureExport.VB_Description = "Recreates project folder structure in the Project folder (for the entire project or just a path matching provided prefix)."
    ProjectStructureParse
    
    Dim Path As Variant
    Dim PrefixMatch As String
    For Each Path In this.Paths
        Select Case Len(Path) - Len(Prefix)
            Case Is > 0
                PrefixMatch = Prefix & Application.PathSeparator
            Case 0
                PrefixMatch = Prefix
            Case Is < 0
                PrefixMatch = vbNullString
        End Select
            
        If (Len(PrefixMatch) > 0) And (Left$(Path, Len(PrefixMatch)) = PrefixMatch) Or (Len(Prefix) = 0) Then
            MkPath this.ProjectPath & Application.PathSeparator & Path
        End If
    Next Path
End Sub


'If provided, prefix must use system path separator
'@Description "Exports code modules (entire project or prefix matching path) to Project folder."
Public Sub ProjectFilesExport(Optional ByVal Prefix As String = vbNullString)
Attribute ProjectFilesExport.VB_Description = "Exports code modules (entire project or prefix matching path) to Project folder."
    ProjectStructureExport Prefix
    
    Dim File As Variant
    Dim Path As String
    Dim PrefixMatch As String
    For Each File In this.Files.Keys
        Path = this.Files(File)(2)
        Select Case Len(Path) - Len(Prefix)
            Case Is > 0
                PrefixMatch = Prefix & Application.PathSeparator
            Case 0
                PrefixMatch = Prefix
            Case Is < 0
                PrefixMatch = vbNullString
        End Select
            
        If (Len(PrefixMatch) > 0) And (Left$(Path, Len(PrefixMatch)) = PrefixMatch) Or (Len(Prefix) = 0) Then
            this.Project.VBComponents(File).Export this.ProjectPath & Application.PathSeparator & Path & Application.PathSeparator & File & this.Files(File)(0)
        End If
    Next File
End Sub


Private Sub WalkTreeCore(ByVal Folder As Folder)
    Dim SubFolder As Folder
    
    this.Paths(Folder.Path) = vbNullString
    For Each SubFolder In Folder.SubFolders
        WalkTreeCore SubFolder
    Next SubFolder
End Sub


Public Sub WalkTree(Optional ByVal Prefix As String = vbNullString, Optional ByVal DoNotImportFilesFromRoot As Variant = Empty)
    this.Paths.RemoveAll
    
    Dim RootPrefix As String
    RootPrefix = IIf(Prefix = vbNullString, this.ProjectPath, this.ProjectPath & Application.PathSeparator & Prefix)
    
    Dim Root As Folder
    Set Root = this.fso.GetFolder(RootPrefix)
    
    WalkTreeCore Root
    Dim SkipRootDir As Boolean
    SkipRootDir = IIf(IsEmpty(DoNotImportFilesFromRoot), (Prefix = vbNullString), DoNotImportFilesFromRoot)
    If SkipRootDir Then this.Paths.Remove RootPrefix
End Sub


Public Sub CollectFiles(Optional ByVal Prefix As String = vbNullString, Optional ByVal DoNotImportFilesFromRoot As Variant = Empty)
    WalkTree Prefix, DoNotImportFilesFromRoot
    
    this.Files.RemoveAll
    
    Dim File As IWshRuntimeLibrary.File
    Dim Path As Variant
    For Each Path In this.Paths
        For Each File In this.fso.GetFolder(Path).Files
            this.Files(this.fso.GetBaseName(File.Name)) = Array(this.fso.GetExtensionName(File.Name), Path)
        Next File
    Next Path
End Sub


Public Sub ProjectFilesImport(Optional ByVal Prefix As String = vbNullString, Optional ByVal DoNotImportFilesFromRoot As Variant = Empty)
    Dim DestructiveWarning As Boolean
    Dim Message As String
    Message = "Warning, you are about to overwrite existing project modules in project " & UCase$(this.Project.Name) & "! Continue?"
    DestructiveWarning = MsgBox(Message, vbYesNo + vbExclamation + vbDefaultButton2)
    If Not DestructiveWarning Then Exit Sub
    
    CollectFiles Prefix, DoNotImportFilesFromRoot
    
    Dim Path As String
    Dim FileName As Variant
    Dim Ext As String
    Dim PathName As String
    
    Dim Module As VBIDE.VBComponent
    Dim DummyModule As VBIDE.VBComponent
    On Error Resume Next
    For Each FileName In this.Files.Keys
        Set Module = this.Project.VBComponents(FileName)
        Path = this.Files(FileName)(1)
        Ext = this.Files(FileName)(0)
        If Ext = "frx" Then
            Ext = "frm"
            this.Files(FileName)(0) = Ext
        End If
        PathName = Path & Application.PathSeparator & FileName & "." & Ext
        
        If Not Module Is Nothing Then this.Project.VBComponents.Remove Module
        Select Case Ext
            Case "cls", "bas", "frm"
                this.Project.VBComponents.Import PathName
            Case "doccls"
                Set DummyModule = this.Project.VBComponents.Import(PathName)
                Module.CodeModule.DeleteLines 1, Module.CodeModule.CountOfLines
                Module.CodeModule.InsertLines 1, DummyModule.CodeModule.Lines(1, DummyModule.CodeModule.CountOfLines)
                this.Project.VBComponents.Remove DummyModule
        End Select
    Next FileName
    On Error GoTo 0
    
    '''' If "References.xsv" exists in the prefix root, add the references.
    PathName = this.ProjectPath _
             & IIf(Prefix <> vbNullString, Application.PathSeparator & Prefix, vbNullString) _
             & Application.PathSeparator & REFERENCES_FILE
    If Dir$(PathName) <> vbNullString Then ReferencesAddFromFile PathName
End Sub


'Document modules are skipped
'@Description "Deletes all modules from the project node subtree identified by the prefix"
Public Sub ProjectNodeSubtreeDelete(ByVal Prefix As String)
Attribute ProjectNodeSubtreeDelete.VB_Description = "Deletes all modules from the project node subtree identified by the prefix"
    If Prefix = vbNullString Then
        Err.Raise 17, "RDVBAProjectUtils", "Empty prefix is not allowed"
    End If
    
    ProjectStructureParse
    
    Dim File As Variant
    Dim Path As String
    For Each File In this.Files.Keys
        Path = this.Files(File)(2)
            
        If (Left$(Path, Len(Prefix)) = Prefix) Then
            If this.Project.VBComponents(File).Type <> vbext_ct_Document Then
                this.Project.VBComponents.Remove this.Project.VBComponents(File)
            End If
        End If
    Next File
End Sub
