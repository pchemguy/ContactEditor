### Overview

"Contact Editor" uses two base model classes. *DataRecordModel* holds a single record (data table row) and is behind the "record editor" user form. *DataTableModel* represents a whole data table or a subset of rows, abstracting persistent storage. Storage backends transfer data between data models and persistent storage.

*DataRecordModel* has one backend, *DataRecordWSheet*,  which saves the record to Excel Worksheet and populates the model at application startup. "Record backends" implement the *IDataRecordStorage* interface. Abstract factory *DataRecordFactory*, implementing the *IDataRecordFactory* interface, instantiates record backends. See [Fig. 1](#FigBaseClassDiagram), left schematic.

*DataTableModel* has three backends, including *DataTableWSheet*, *DataTableCSV*, and *DataTableADODB*. They handle a Worksheet, a delimited text file, and a relational database, respectively. "Table backends" implement the *IDataTableStorage* interface. Abstract factory *DataTableFactory*, implementing the *IDataTableFactory* interface, instantiates table backends. See [Fig. 1](#FigBaseClassDiagram), right schematic.

*DataRecordManager* and *DataTableManager* are composite classes, implementing *IDataRecordManager* and *IDataTableManager* interfaces. These classes incorporate by composition one model and one backend class, yielding "backend-managed" models.

<a name="FigBaseClassDiagram"></a>

<img src="https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram%20-%20Table%20and%20Record.svg?raw=true" alt="Overview" width="100%" />

<p align="center"><b>Figure 1. Base class diagram</b></p>

*DataRecordModel* and *DataTableModel* work cooperatively: *DataRecordModel* holds row from the row set held in *DataTableModel*. Since data needs to be transferred between the two model classes, a composite manager, *DataCompositeManager*, handling such transfers is necessary. The composite manager can either incorporate two "backend-managed" classes or two model and two backend classes directly. *DataCompositeManager* uses the latter option. Finally, the *ContactEditorModel*, the main application backend-managed model, incorporates *DataCompositeManager*. See [Fig. 2](#FigCompositeManager).

<a name="FigCompositeManager"></a>

<img src="https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram.svg?raw=true" alt="Overview" width="100%" />

<p align="center"><b>Figure 2. Composite manager</b></p>

### Implementation Details

#### DataTableModel

`Values` Variant is the main data field in the DataTableModel. This field is populated with a record-wise 2D 1-based (in both dimensions) array (the faster changing index is the field index).  The first column is assumed to be the id/primary key (PK) field (multi-column PKs are not supported by the present implementation).  

`FieldNames` - 1D 1-based Variant array containing the names of the fields in the order they appear in the table from left to right. This array is automatically populated by the Table backend, so no field names need to be hardcoded for this purpose.  

The main record id is its primary key (which at present is assumed to be a single left-most column). Since id can, in general, be numeric (not necessarily continuous) or textual, the current approach is to cast the id field as String (`RecordId`) and use the record identifier. Similarly, field names act as field/column IDs, and both need to be mapped to the respective indices in the Values array. Hence, two dictionary-based mapping fields have been added to the DataTableModel.  

`IdIndices` maps RecordId to RecordIndex.  

`FieldIndices` maps FieldName to FieldIndex.  

`DirtyRecords` is the last field of the DataTableModel, used to store RecordId's for modified records. Only these records need to be saved to the backend when the user requests to save the changes.

#### DataRecordModel

`Record` field wraps a Scripting.Dictionary object, storing the current record data as a "Field Name"<>"Value" map. Additionally, a dirty flag is also included to indicated that the user made changes to the data.  

#### Data Transfer

DataTableModel also provides two important methods, `UpdateRecordFromDictionary` and `CopyRecordToDictionary`. The former takes a dictionary "Field Name"<>"Value" map, such as provided by the DataRecordModel, and updates data in the corresponding record in the Values array using the FieldIndices map. The latter does the opposite, based on supplied RecordId and IdIndices map.

[Composite classes]: https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram.svg
[Base classes]: https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram%20-%20Table%20and%20Record.svg
