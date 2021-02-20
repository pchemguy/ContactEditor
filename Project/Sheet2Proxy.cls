VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2Proxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IDataSheet

Private Property Get Table() As ListObject
    Set Table = Sheet2.ListObjects(1)
End Property

Private Property Get CodeColumnIndex() As Long
    CodeColumnIndex = Table.ListColumns("Code").Index
End Property

Private Property Get NameColumnIndex() As Long
    NameColumnIndex = Table.ListColumns("Name").Index
End Property

Private Property Get IDataSheet_Data() As Collection
    
    Dim result As Collection
    Set result = New Collection
    
    Dim currentRow As ListRow
    For Each currentRow In Table.ListRows
        
        Dim currentItem As SomeModel
        Set currentItem = New SomeModel
        
        currentItem.Code = currentRow.Range(ColumnIndex:=CodeColumnIndex).value
        currentItem.Name = currentRow.Range(ColumnIndex:=NameColumnIndex).value
        
        result.Add currentItem
    
    Next
    
    Set IDataSheet_Data = result
    
End Property
