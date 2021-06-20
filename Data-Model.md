---
layout: default
title: Data model
nav_order: 2
permalink: /data-model
---

### Overview

"Contact Editor" uses two base model classes. *DataRecordModel* holds a single record (data table row) and is behind the "record editor" user form. *DataTableModel* represents a whole data table or a subset of rows, abstracting persistent storage. Storage backends transfer data between data models and persistent storage.

*DataRecordModel* has one backend, *DataRecordWSheet*,  which saves the record to Excel Worksheet and populates the model at application startup. "Record backends" implement the *IDataRecordStorage* interface. Abstract factory *DataRecordFactory*, implementing the *IDataRecordFactory* interface, instantiates record backends. See [Fig. 1](#FigDataRecordModel).

<a name="FigDataRecordModel"></a>

<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/develop/Assets/Diagrams/Class%20Diagram%20-%20Record.svg" alt="FigDataRecordModel" width="100%" />

<p align="center"><b>Figure 1. DataRecordModel class diagram</b></p>

*DataTableModel* has three backends, including *DataTableWSheet*, *DataTableCSV*, and *DataTableADODB*. They handle a Worksheet, a delimited text file, and a relational database, respectively. "Table backends" implement the *IDataTableStorage* interface. Abstract factory *DataTableFactory*, implementing the *IDataTableFactory* interface, instantiates table backends. See [Fig. 2](#FigDataTableModel).

<a name="FigDataTableModel"></a>

<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/develop/Assets/Diagrams/Class%20Diagram%20-%20Table.svg" alt="FigDataTableModel" width="100%" />

<p align="center"><b>Figure 2. DataTableModel class diagram</b></p>

*DataRecordManager* ([Fig. 1](#FigDataRecordModel)) and *DataTableManager* ([Fig. 2](#FigDataTableModel)) are composite classes, implementing *IDataRecordManager* and *IDataTableManager* interfaces. These classes incorporate by composition one model and one backend class, yielding "backend-managed" models.

*DataRecordModel* and *DataTableModel* work cooperatively: *DataRecordModel* holds a row from the row set held in *DataTableModel*. Since data needs to be transferred between the two model classes, a composite manager, *DataCompositeManager*, handling such transfers is necessary. The composite manager can either incorporate two "backend-managed" classes or two model and two backend classes directly. 

*DataCompositeManager* uses the latter option. Finally, the *ContactEditorModel*, the main application backend-managed model, incorporates *DataCompositeManager*. See [Fig. 3](#FigCompositeManager).

<a name="FigCompositeManager"></a>

<img src="https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram.svg?raw=true" alt="Overview" width="100%" />

<p align="center"><b>Figure 3. Composite manager</b></p>


### Implementation Details

#### DataRecordModel

*Record* field wraps a Dictionary object, storing the current record as a "Field Name"<>"Value" map. *ContactEditorPresenter* takes this model's data and populates *ContactEditorForm*. The  "Change" events of the controls in *ContactEditorForm* perform the reverse update. Additionally, a dirty flag is also included to indicated that the user made changes to the data.  

#### DataTableModel

The *DataTableModel* class handles tabular data and acts as an intermediary between the user and persistent storage. A typical data flow in this arrangement involves three essential tasks:

1. saving received records in the *DataTableModel*,
2. enabling the user access to specific data elements, and
3. persisting modified data if necessary.

Naturally, *DataTableModel* uses a 2D record-wise Variant array called *Values* for internal storage of a set of table rows. (The field index is the faster changing index, and both indices are 1-based.) Accessing a specific data element is about addressing it. For a 2D array, the address of a data element is its (record/field) indices. Finally, to persist a record, the backend must map record/field indices to the primary key and field name coordinates used by persistent storage. Further, field indices are meaningless to the user, so the GUI layer also needs the field index to field name map.

Since *DataTableModel* deals with tabular data only, saving the mapping metadata does not introduce additional coupling. However, since relational databases should be the dominant data source, the following assumptions about the field names and primary keys have simplified this implementation.

It is safe to assume that field names are strings (textual). Therefore, the model includes two structures to support conversion between the field name and its index. *FieldNames* is a 1D 1-based Variant array containing the names of the fields in the order they appear in the table from left to right and providing the index to name mapping. Table backend populates this field automatically, avoiding the use of hardcoded names. *FieldIndices* is a Dictionary providing the reverse name to index mapping. 

The assumption that the first record field (called *RecordId*) is a scalar (single field) primary key (PK) is not always valid, but it simplifies code logic. Further, string casting *RecordId* removes type-related ambiguity and provides two other benefits. *IdIndices* uses the string *RecordId* key in the Dictionary mapping RecordId to record index; the *Values* array provides the reverse mapping.

The other GUI-related benefit is more subtle and, in general, is not the model's concern. *ContactEditorForm* uses PKs for record selection via a drop-down combo list control. The control's 1D Variant array attribute takes PKs from the *Values* array, and its other attribute holds the "current value" of this control, which is always a string. If the populated PK array is numeric, typing in RecordId will not match an existing list element because the combo control does not do typecasting (String("1") <> Integer(1)).

*DirtyRecords* is a Dictionary, collecting (*RecordId*, *RecordIndex*) pairs of modified records. Only these records need to be saved by the backend when the user requests to save the changes.

*UpdateRecordFromDictionary* takes "Field Name"<>"Value" Dictionary and updates data in the corresponding record in the *Values* array using the *FieldIndices* map.

*CopyRecordToDictionary* does the opposite operation based on supplied *RecordId* and *IdIndices* map.
