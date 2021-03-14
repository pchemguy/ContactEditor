
### Overview

At the basis of the project model are two data data model classes, `DataRecordModel`, representing a single Record (table row) shown to the user on the UserForm, and `DataTableModel`, abstracting persistent storage. Model classes know nothing about data storage, which is the responsibility of the backend classes.  

Presently, there is one backend for DataRecordModel, `DataRecordWSheet`, which saves the last saved record to Excel Worksheet, populating the UserForm with this record at startup.

DataTableModel has three backends, `DataTableWSheet`, `DataTableCSV`, and `DataTableADODB`, responsible for handling a Worksheet, a delimited text file, and a generic RDBMS respectively.  
 
The backends correspondingly implement either IDataRecordStorage or IDataTableStorage interface are instantiated by their respective abstract factories: DataRecordFactory\IDataRecordFactory and  DataTableFactory\IDataTableFactory. In turn, DataRecordManager\IDataRecordManager and DataTableManager\IDataTableManager incorporate by composition one model and one backend class, yielding a "backend-managed" model.

![Base classes][Base classes]

DataRecordModel and DataTableModel work in concert, with DataRecordModel holding a piece of information from the DataTableModel. Data chunks need to travel between the two models; hence, the need for additional functionality, which should a part of a derived manager. Thus, `DataCompositeManager` class has been added. In VBA, composition is used for such purposes, and the composite manager can be implemented either via a composition of "backend-managed" classes, or via a composition of two model and two backend classes directly. The latter pathway has been chosen for  DataCompositeManager class, and it exposes the necessary features of the constituent classes and implements the "inter-model" functionality. Finally, the viewModel class ContactEditorModel incorporates DataCompositeManager.

![Composite classes][Composite classes]

### Implementation Details

#### DataTableModel

`Values` Variant is the main data field in the DataTableModel. This field is populated with a record-wise 2D 1-based (in both dimensions) array (the faster changing index is the field index).  The first column is assumed to be the id/primary key (PK) field (multi column PKs are not supported by the present implementation).  

`FieldNames` - 1D 1-based Variant array containing the names of the fields in the order they appear in the table from left to right. This array is automatically populated by the Table backend, so no field names need to be hardcoded for this purpose.  

The main record id is its primary key (which at present to be assumed to be a single left-most column). Since id can, in general, be numeric (not necessarily continuous) or textual, the current approach is to cast the id field as String (`RecordId`) and use the the record identifier. Similarly, field names act as field/column id's, and both needs to be mapped the respective indices in the Values array. Hence, two dictionary-based mapping fields have been added to the DataTableModel.  

`IdIndices` maps RecordId to RecordIndex.  

`FieldIndices` maps FieldName to FieldIndex.  

`DirtyRecords` is the last field of the DataTableModel, used to store RecordId's for modified records. Only these records need to be saved to the backend when user requests to save the changes.

#### DataRecordModel

`Record` field wraps a Scripting.Dictionary object, storing the current record data as a "Field Name"<>"Value" map. Additionally, a dirty flag is also included to indicated that the user made changes to the data.  

#### Data Transfer

DataTableModel also provides two important methods, `UpdateRecordFromDictionary` and `CopyRecordToDictionary`. The former takes a dictionary "Field Name"<>"Value" map, such as provided by the DataRecordModel, and updates data in the corresponding record in the Values array using the FieldIndices map. The latter does the opposite, based on supplied RecordId and IdIndices map.

[Composite classes]: https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram.svg
[Base classes]: https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram%20-%20Table%20and%20Record.svg
