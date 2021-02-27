VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRecordWSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Abstracts Excel worksheet"
'@Folder "Storage.Record.Backend"
'@ModuleDescription "Abstracts Excel worksheet"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
'@Exposed
Option Explicit

Implements IDataRecordStorage

Private Type TDataRecord
    Model As DataRecordModel
    WSheet As Excel.Worksheet
End Type
Private this As TDataRecord


'@Description "Returns a new IDataRecord object."
Public Function Create(ByVal Model As DataRecordModel, ByVal ConnectionString As String, ByVal TableName As String) As IDataRecordStorage
Attribute Create.VB_Description = "Returns a new IDataRecord object."
    Dim result As DataRecordWSheet
    Set result = New DataRecordWSheet
    result.Init Model, ConnectionString, TableName
    Set Create = result
End Function


'@Ignore ParameterNotUsed
Public Sub Init(ByVal Model As DataRecordModel, ByVal ConnectionString As String, ByVal TableName As String)
    Set this.Model = Model
    Dim BookSheet As Variant: BookSheet = Split(ConnectionString, "!")
    Set this.WSheet = Excel.Workbooks(BookSheet(0)).Worksheets(BookSheet(1))
End Sub


Private Sub Class_Terminate()
    Set this.Model = Nothing
End Sub


'''' Collect all named ranges, keep only those referring to a single cell on a proper worksheet
'''' and the adjacent label above or to the left matches the range name.
Private Sub InitModel()
    Dim WSheetName As String: WSheetName = this.WSheet.Name
    Dim Prefix As String: Prefix = "=" & WSheetName & "!"
    Dim PrefixLength As Long: PrefixLength = Len(Prefix)
    
    Dim RangeName As Excel.Name
    Dim TargetRange As Excel.Range
    Dim FieldName As String
    Dim RefersTo As String
    Dim CellCount As Long
    Dim WBook As Excel.Workbook: Set WBook = this.WSheet.Parent
    For Each RangeName In WBook.Names
        ' Check that this range name refers to the desired Worksheet (RefersTo starts with "=<Worksheet Name>!
        RefersTo = RangeName.RefersTo
        If Left$(RefersTo, PrefixLength) <> Prefix Then GoTo RANGENAME_CONTINUE
        
        ' Check that this range name refers to a single cell
        Set TargetRange = this.WSheet.Range(Right$(RefersTo, Len(RefersTo) - PrefixLength))
        CellCount = TargetRange.Cells.Count
        If CellCount > 1 Then GoTo RANGENAME_CONTINUE
        
        ' For Worksheet scoped range names, strip Worksheet name and check that it matches the label in adjacent top or left cell
        FieldName = RangeName.Name
        If Left$(FieldName, PrefixLength - 1) = WSheetName & "!" Then FieldName = Right$(FieldName, Len(FieldName) - (PrefixLength - 1))
        
        ' Check that the range name matches the label (the value of an adjacent cell above or to the left of the candidate cell)
        If (TargetRange.Offset(-1, 0).Value <> FieldName) And (TargetRange.Offset(0, -1).Value <> FieldName) Then GoTo RANGENAME_CONTINUE
        
        'Add FieldName -> TargetRange.Value to the dictionary
        this.Model.Record(CStr(FieldName)) = TargetRange.Value
RANGENAME_CONTINUE:
    Next RangeName
End Sub


Private Sub IDataRecordStorage_LoadDataIntoModel()
    If this.Model.Record.Count > 0 Then
        Dim FieldName As Variant
        For Each FieldName In this.Model.Record.Keys
             this.Model.Record(CStr(FieldName)) = this.WSheet.Range(FieldName).Value
        Next FieldName
    Else
        InitModel
    End If
End Sub


Private Sub IDataRecordStorage_SaveDataFromModel()
    With this.Model
        If Not .IsDirty Then Exit Sub
        
        Dim FieldName As Variant
        For Each FieldName In .Record.Keys
            this.WSheet.Range(FieldName).Value = .Record(CStr(FieldName))
        Next FieldName
        
        .IsNotDirty
    End With
End Sub
