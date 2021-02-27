Attribute VB_Name = "StorageDescription"
'@Folder("Storage")
Option Explicit

'''' Storage package provides storage management for two types of models:
''''    - Record subpackage
''''    - Table subpackage
'''' Both packages have very similar structure, including the following classes:
''''    - Model
''''    - Backends with an interface
''''    - Abstract factory for backends with an interface
''''    - Manager with an interface
''''
'''' Record subpackage models a single table record/row, and DataRecordModel uses
'''' a Dictionary object to store FieldName -> CStr(Value) mapping. Such a model
'''' can be used to represent a UserForm data displaying a single record data.
''''
'''' Table subpackage models a single table and stores
''''    - an array of field names
''''    - a mapping FieldName -> ColumnIndex in a Dictionary object
''''    - a 2D Variant array in the row-wise order.
''''
'''' Each backend class provides two methods to move the data between the model and
'''' persistent storage (Load and Save).
''''
''''
'''' Manager is a composition of a model and a backend, exposing the model getter and
'''' Load/Save methods.
''''


Public Function GetTopLeftCell(ByVal RangeName As String) As Variant
    '@Ignore ImplicitActiveSheetReference
    GetTopLeftCell = Range(RangeName).Range("A1").Value
End Function

