VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetBProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Forms.Workbook"
'@IgnoreModule IndexedDefaultMemberAccess

Option Explicit

Implements IDataSheet

Private Type TData
    TableAMap As Scripting.Dictionary
End Type
Private this As TData


Private Sub Class_Initialize()
    Set this.TableAMap = New Scripting.Dictionary
    this.TableAMap.CompareMode = TextCompare
    
    Dim TableHeader As Excel.Range: Set TableHeader = CodeNameTable.Rows(1)
    Dim FieldCount As Long: FieldCount = TableHeader.Columns.Count
    Dim FieldIndex As Long
    For FieldIndex = 1 To FieldCount
        this.TableAMap.Item(TableHeader.Cells(ColumnIndex:=FieldIndex).Value) = FieldIndex
    Next FieldIndex
End Sub

Private Sub Class_Terminate()
    Set this.TableAMap = Nothing
End Sub

Private Property Get CodeNameTable() As Excel.Range
    Set CodeNameTable = CodeNameData.Range("CodeNameTable").Rows
End Property

Private Property Get CodeColumnIndex() As Long
    CodeColumnIndex = this.TableAMap("Code")
End Property

Private Property Get NameColumnIndex() As Long
    NameColumnIndex = this.TableAMap("Name")
End Property

Private Property Get IDataSheet_Data() As Collection
    Dim result As Collection
    Set result = New Collection
    
    Dim RowIndex As Long
    For RowIndex = 2 To CodeNameTable.Count
        Dim currentItem As SheetBModel
        Set currentItem = New SheetBModel
        
        currentItem.Code = CodeNameTable.Item(RowIndex).Cells(ColumnIndex:=CodeColumnIndex).Value
        currentItem.Name = CodeNameTable.Item(RowIndex).Cells(ColumnIndex:=NameColumnIndex).Value
        
        result.Add currentItem
    Next
    
    Set IDataSheet_Data = result
End Property
