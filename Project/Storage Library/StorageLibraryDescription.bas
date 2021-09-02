Attribute VB_Name = "StorageLibraryDescription"
'@Folder "Storage Library"
'@IgnoreModule EmptyModule: This module acts as package description
Option Explicit

'''' Storage package provides storage management for two types of models:
''''    - Record subpackage
''''    - Table subpackage
'''' Both packages have very similar structure, including the following classes:
''''    - Model
''''    - Backends implementing a common interface
''''    - Abstract factory for backends implementing a common interface
''''    - Manager implementing a common interface
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
'''' Each backend class provides two methods to move the data between the model
'''' and persistent storage (Load and Save).
''''
'''' Manager is a composition of a model and a backend pair from either Table
'''' or Record subpackage), exposing the model getter and Load/Save methods.
''''
'''' DataCompositeManager class incorporates one Table and one Record model with
'''' their backends. Record submodel is used to represent a row from the Table.
''''
'''' While basic single subpackage managers are implememted as predeclared
'''' classes with factories, DataCompositeManager is a regular class without
'''' a factory. While this design deviates from the common project pattern,
'''' employing the factory pattern would result in fairly complex factory
'''' signature. Instead, the class provides to individual methods replacing
'''' the factory/constructor pattern, InitRecord and InitTable, to be called
'''' directly.
''''
