---
layout: default
title: ADODB
nav_order: 4
parent: Backends
permalink: /backends/adodb
---

### Overview

*DataTableADODB* enables the application to load data into *DataTableModel* from a relational database via the ADODB library. The backend consists of the main module, the *DataTableADODB* class, and two additional classes.  The *ADOlib* class contains ADODB related helper routines, and the *SQLlib* class provides SQL templates and generates typical SQL queries. To make internal class methods accessible to unit testing, they are declared as Public. Additionally, *DataTableADODB* uses one generic helper routine from the "CommonRoutines" module.

The current implementation provides limited functionality using ADODB directly. *DataTableADODB* assumes that

- the first record field is a single field primary key, and
- each *DataTableADODB* instance accesses a single database table.

This functionality may be extended in the future, e.g., via the [SecureADODB][SecureADODB] or its [fork][SecureADODB fork].

### Core attributes and methods

DataTableADODB constructor needs a proper *ConnectionString* to complete initialization. To simplify the code, ADOlib provides connection string building helpers. Currently, only an SQLite helper is available. The constructor checks whether the provided connection string has a database-specific prefix and calls the appropriate helper if necessary. In the case of SQLite, the GetSQLiteConnectionString helper, in turn, calls *CommonRoutines.VerifyOrGetDefaultPath*. The latter takes a *FilePathName* candidate and an array containing default extensions. If *FilePathName* is not a valid path-name to an existing file, this helper also checks the folder containing the Workbook for a file named \<workbook name\>.\<supplied extension\>. If any such file exists, the connection string helper uses its path for connection string construction.

Database introspection reduces the need to hardcode table metadata, with basic metadata being available via either the ADOX extension library or dummy ADODB queries. During instantiation, the backend calls *ADOLib.GetTableMeta* and populates *FieldNames*, *FieldTypes*, and *FieldMap* (renamed *FieldIndicies* from [DataTableModel][DataTableModel]).

The *AdoCommandInit* method sets the *AdoCommand* attribute storing a reusable reference to an ADODB.Command object. *AdoCommandInit* optionally takes an SQL query (defaults to basic SELECT) and *CursorLocation* (defaults to client-side). After initialization, the command can be reused multiple times either as the source for an ADODB.Recordset or on its own for modifying queries.

The *AdoRecordset* method executes the query saved in the *AdoCommand* and returns a Recordset. If provided an optional *SQLQuery*, *AdoRecordset* calls *AdoCommandInit* updating the query before execution. 

The *Records* method uses the *GetRows* method on the Recordset returned by *AdoRecordset*, and the *RecordsAsText* method additionally requests the backend to cast all fields as text. The returned result is a column-wise 2D Variant array, and "WorksheetFunction.Transpose" yields a row-wise array.

Warning: it turned out that "WorksheetFunction.Transpose" is limited and should not be used for anything serious. It has a hardcoded limit on the size of the transposed array, which, at least in Excel 2002, is too small. More importantly, it converts Variant/Integer (from Recordset.GetRows) to Variant/Double. This silent conversion had caused subtle difficult to trace annoying issues with the ID field before the addition of string-casting (see the GUI note in the [Data Model section][DataTableModel]). This function needs to be replaced, for example, with the routine provided by the [Chip Pearson's VBA Array library][VBAArrayLib].

As discussed in the [Data Model section][DataTableModel], string-casting the "ID" column is desirable. The tests module illustrates the use of *SQLlib.SelectIdAsText* for the generation of a "SELECT" query template with the typecasting request. Similarly, *SQLlib.SelectAllAsText* requests string-casting for all fields.

### Persisting changes

When I had been working on the current version of the backend, I could not make Recordset.UpdateBatch to work, so I had to resort to generating SQL UPDATEs. Two helper routines, *MakeAdoParamsForRecordUpdate* and *RecordToAdoParams*, from the ADOlib library and *UpdateSingleRecord* from SQLlib help prepare UPDATE statements within the same assumptions as discussed above. 

*UpdateSingleRecord* generates a single record UPDATE statement fully parametrized with respect to all fields, including the ID column in the WHERE clause.

*MakeAdoParamsForRecordUpdate* takes the *FieldNames* and *FieldTypes* arrays and the *AdoCommand*. It clears AdoCommand.Parameters and repopulates it with dummy Parameter objects. For a new Parameter, three attributes must be provided explicitly, including *type*, *length*, and *value*. The name attribute is also set to field name for easier matching with the corresponding record field value. The Parameter type is set to the actual value collected during introspection. Length and value dummies must be provided; otherwise, they remain *Empty*, triggering an error. At the same time, dummy integers (length=1 and value=0) can be supplied to the factory regardless of type, and the factory performs type coercion as necessary. The first field (PK) goes last in the UPDATE statement (in the WHERE clause), so it is added to the Parameters collection at the end.

*RecordToAdoParams* takes a record dictionary and updates the Parameters collection by matching field and parameter names. It updates the length attribute before the value to prevent potential errors. Again, the string-cast primary key value can be added to the corresponding parameter, and the setter will do type coercion if necessary.

*IDataTableStorage_SaveDataFromModel* interface from DataTableADODB takes the *DirtyRecords* Dictionary from the table model and loops through it inside a transaction. Inside the loop, *CopyRecordToDictionary* copies individual records from the *Values* array to a Dictionary; the helper updates field values into the Parameters collection of *AdoCommand*, and *AdoCommand* is executed.


[SecureADODB]: https://github.com/rubberduck-vba/examples/tree/master/SecureADODB
[SecureADODB fork]: https://github.com/pchemguy/RDVBA-examples
[Multiple interfaces]: https://pchemguy.github.io/ContactEditor/class-design
[DataTableModel]: https://pchemguy.github.io/ContactEditor/data-model#datatablemodel
[VBAArrayLib]: http://cpearson.com/excel/vbaarrays.htm