---
layout: default
title: Implementation details
nav_order: 2
parent: Storage Library overview
permalink: /storagelibrary/implementationdetails
---

#### DataRecordModel

*Record* field wraps a Dictionary object, storing the current record as a "Field&nbsp;Name"&nbsp;&rightarrow;&nbsp;"Value" map. *ContactEditorPresenter* takes this model's data and populates *ContactEditorForm*. The  "Change" events of the controls in *ContactEditorForm* perform the reverse update. Additionally, a dirty flag is included to indicated that the user made changes to the data.  

#### DataTableModel

The *DataTableModel* class handles tabular data and acts as an intermediary between the user and persistent storage. A typical data flow in this arrangement involves three essential tasks:

1. saving received records in the *DataTableModel*,
2. enabling the user access to specific data elements, and
3. persisting modified data if necessary.

*DataTableModel* uses a 2D record-wise Variant array called *Values* for internal storage of a set of table rows. (The field index is the faster changing index, and both indices are 1-based.) Accessing a specific data element is about addressing it. For a 2D array, the address of a data element is its record/field indices, and persistent storage typically uses PK/FieldName addressing. Since *DataTableModel* stores record sets, it can also store metadata necessary for translating the two addressing schemes without introducing additional coupling.

For simplicity, only scalar (single field) PK is supported, and it should be the first field in the table. PK is cast as String so that the same logic could be applied to both integer and string keys, and a string variable called *RecordId* is used to store particular values. The *Values* array provides \<record&nbsp;index\>&nbsp;&rightarrow;&nbsp;\<PK\> mapping, and the *IdIndices* dictionary provides the reverse mapping.

Typecasting PK also has a subtle GUI-related benefit. In the current implementation, *ContactEditorForm* uses PKs for record selection via a drop-down combo list control. Combo's *List* attribute is populated with PKs from the *Values* array, and another relevant attribute *Value* is initialized from *DataRecordModel*. Without proper care, the two attributes may end up being of different types (e.g., Double and String). Since the *Value* attribute should match one of the elements from the *List* array, such a type mismatch may cause difficult to debug glitches.

*FieldNames* is a 1D 1-based array. It contains the names of the fields in the order matching the structure of the table from left to right and provides \<field&nbsp;index\>&nbsp;&rightarrow;&nbsp;\<field&nbsp;name\> mapping. *FieldIndices* dictionary provides the reverse mapping. Table backend populates this field automatically using introspection, avoiding the use of hardcoded names.

*DirtyRecords* is a Dictionary, collecting (*RecordId*, *RecordIndex*) pairs of modified records. Only these records need to be saved by the backend when the user requests to save the changes.

*UpdateRecordFromDictionary* takes "Field&nbsp;Name"&nbsp;&rightarrow;&nbsp;"Value" Dictionary and updates data in the corresponding record in the *Values* array using the *FieldIndices* map.

*CopyRecordToDictionary* does the opposite operation based on supplied *RecordId* and *IdIndices* map.

#### DataCompositeManager

While Storage Library classes typically employ the factory/constructor design pattern combined with the *Predeclared* attribute, using this pattern for the *DataCompositeManager* class would result in an overly complicated factory signature. Instead, *DataCompositeManager* is a regular class, and its two methods, *InitRecord* and *InitTable*, replace the factory/constructor pattern. The user should use the *New* operator and then call these methods directly to complete manager initialization.
