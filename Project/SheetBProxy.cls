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
    
    Dim TableHeader As Excel.Range: Set TableHeader = TableA.Rows(1)
    Dim FieldCount As Long: FieldCount = TableHeader.Columns.Count
    Dim FieldIndex As Long
    For FieldIndex = 1 To FieldCount
        this.TableAMap.Item(TableHeader.Cells(ColumnIndex:=FieldIndex).value) = FieldIndex
    Next FieldIndex
End Sub

Private Sub Class_Terminate()
    Set this.TableAMap = Nothing
End Sub

Private Property Get TableA() As Excel.Range
    Set TableA = SheetB.Range("TableA").Rows
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
    For RowIndex = 2 To TableA.Count
        Dim currentItem As SheetBModel
        Set currentItem = New SheetBModel
        
        currentItem.Code = TableA.Item(RowIndex).Cells(ColumnIndex:=CodeColumnIndex).value
        currentItem.Name = TableA.Item(RowIndex).Cells(ColumnIndex:=NameColumnIndex).value
        
        result.Add currentItem
    Next
    
    Set IDataSheet_Data = result
End Property
