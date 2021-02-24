VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Abstracts Excel worksheet"
'@Folder "Storage.Map.Backend"
'@ModuleDescription "Abstracts Excel worksheet"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
'@Exposed
Option Explicit

Implements IDataMap

Private Type TDataMap
    model As DataMapModel
    MapSheet As Excel.Worksheet
End Type
Private this As TDataMap


'@Description "Returns a new IDataMap object."
Public Function Create(ByVal model As DataMapModel, ByVal ConnectionString As String) As IDataMap
Attribute Create.VB_Description = "Returns a new IDataMap object."
    Dim result As SheetMap
    Set result = New SheetMap
    result.Init model, ConnectionString
    Set Create = result
End Function


Public Sub Init(ByVal model As DataMapModel, ByVal ConnectionString As String)
    Set this.model = model
    Dim BookSheet As Variant: BookSheet = Split(ConnectionString, "!")
    Set this.MapSheet = Excel.Workbooks(BookSheet(0)).Worksheets(BookSheet(1))
End Sub


Private Sub Class_Terminate()
    Set this.model = Nothing
End Sub


'''' Collect all named ranges, keep only those referring to a single cell on a proper worksheet
'''' and the adjacent label above or to the left matches the range name.
Private Sub IDataMap_GetData()
    this.model.Properties.RemoveAll
    
    Dim WSheetName As String: WSheetName = this.MapSheet.Name
    Dim Prefix As String: Prefix = "=" & WSheetName & "!"
    Dim PrefixLength As Long: PrefixLength = Len(Prefix)
    
    Dim RangeName As Excel.Name
    Dim TargetRange As Excel.Range
    Dim RefersTo As String
    Dim CellCount As Long
    Dim FieldName As String
    For Each RangeName In Application.Names
        ' Check that this range name refers to the desired Worksheet (RefersTo starts with "=<Worksheet Name>!
        RefersTo = RangeName.RefersTo
        If Left$(RefersTo, PrefixLength) <> Prefix Then GoTo RANGENAME_CONTINUE
        
        ' Check that this range name refers to a single cell
        Set TargetRange = this.MapSheet.Range(Right$(RefersTo, Len(RefersTo) - PrefixLength))
        CellCount = TargetRange.Cells.Count
        If CellCount > 1 Then GoTo RANGENAME_CONTINUE
        
        ' For Worksheet scoped range names, strip Worksheet name and check that it matches the label in adjacent top or left cell
        FieldName = RangeName.Name
        If Left$(FieldName, PrefixLength - 1) = WSheetName & "!" Then FieldName = Right$(FieldName, Len(FieldName) - (PrefixLength - 1))
        If (TargetRange.Offset(-1, 0).Value <> FieldName) And (TargetRange.Offset(0, -1).Value <> FieldName) Then GoTo RANGENAME_CONTINUE
        
        'Add FieldName -> TargetRange.Value to the dictionary
        this.model.Properties.Item(FieldName) = TargetRange.Value
RANGENAME_CONTINUE:
    Next RangeName
End Sub


Private Sub IDataMap_PersistData()
    Dim PropertyName As Variant
    For Each PropertyName In this.model.Properties
        this.MapSheet.Range(PropertyName).Value = this.model.Properties(PropertyName)
    Next PropertyName
End Sub
