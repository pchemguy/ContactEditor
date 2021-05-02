---
layout: default
title: ADODB
nav_order: 3
parent: Backends
permalink: /backends/adodb
---

DataTableADODB constructor needs a proper *ConnectionString* to complete initialization. To simplify the code, ADOlib provides connection string building helpers. Currently, only an SQLite helper is available. The constructor checks whether the provided connection string has a database-specific prefix and calls the appropriate helper if necessary. In the case of SQLite, the GetSQLiteConnectionString helper, in turn, calls *CommonRoutines.VerifyOrGetDefaultPath*. The latter takes a *FilePathName* candidate and an array containing default extensions. If *FilePathName* is not a valid path-name to an existing file, this helper also checks the folder containing the Workbook for a file named <workbook name>.<supplied extension>. If any such file exists, the connection string helper uses its path for connection string construction.

Database introspection reduces the need to hardcode table metadata, with basic metadata being available via either the ADOX extension library or dummy ADODB queries. During instantiation, the backend calls *ADOLib.GetTableMeta* and populates *FieldNames*, *FieldTypes*, and *FieldMap* (renamed *FieldIndicies* from [DataTableModel][DataTableModel]).

The *AdoCommandInit* method sets the *AdoCommand* attribute storing a reusable reference to an ADODB.Command object. *AdoCommandInit* optionally takes an SQL query (defaults to basic SELECT) and *CursorLocation* (defaults to client-side). After initialization, the command can be reused multiple times either as the source for an ADODB.Recordset or on its own for modifying queries.

The *AdoRecordset* method executes the query saved in the *AdoCommand* and returns a Recordset. If provided an optional *SQLQuery*, *AdoRecordset* calls *AdoCommandInit* updating the query before execution. 

The *Records* method uses the *GetRows* method on the Recordset returned by *AdoRecordset*, and the *RecordsAsText* method additionally requests the backend to cast all fields as text. The returned result is a column-wise 2D Variant array, and "WorksheetFunction.Transpose" yields a row-wise array.

Warning: it turned out that "WorksheetFunction.Transpose" is limited and should not be used for anything serious. It has a hardcoded limit on the size of the transposed array, which, at least in Excel 2002, is too small. More importantly, it converts Variant/Integer (from Recordset.GetRows) to Variant/Double. This silent conversion had caused subtle difficult to trace annoying issues with the ID field before the addition of string-casting (see the GUI note in the [Data Model section][DataTableModel]). This function needs to be replaced, for example, with the routine provided by the [Chip Pearson's VBA Array library][VBAArrayLib].

As discussed in the [Data Model section][DataTableModel], string-casting the "ID" column is desirable. The tests module illustrates the use of *SQLlib.SelectIdAsText* for the generation of a "SELECT" query template with the typecasting request. Similarly, *SQLlib.SelectAllAsText* requests string-casting for all fields. Also, note that apart from string-casting, these two routines also spell out each field name and perform aliasing (\<FieldName\> AS \<FieldName\>). Without it, the returned field names follow the verbose template \<TableName\>.\<FieldName\>.

**Update query helpers**. Two other routines from ADOlib, *MakeAdoParamsForRecordUpdate* and *RecordToAdoParams* are used for performing database update. 

*MakeAdoParamsForRecordUpdate* take a list of FieldNames, FieldTypes, and an ADODB.Command. It constructs a basic single record update query, assuming that the first field is the primary key used in the WHERE clause. The update query is parameterized with respect to all field values, and for each field, a parameter objected is created and added to ADODB.Command.Parameters. Then this command can be used to update multiple records by setting the values of Parameter members to the corresponding field values. *RecordToAdoParams* takes a record dictionary and updates parameters by matching the names of the fields with the names of parameters. *IDataTableStorage_SaveDataFromModel* interface from DataTableADODB takes the list of dirty records from the table model, it uses its routine that copies individual records to a dictionary object, then it calls the helper to update parameter values, and executes update query. This process loops through the dirty record list, reusing the same *prepared* command, and the loop is placed inside a transaction.




[SecureADODB]: https://github.com/rubberduck-vba/examples/tree/master/SecureADODB
[SecureADODB fork]: https://github.com/pchemguy/RDVBA-examples
[Multiple interfaces]:  https://github.com/pchemguy/ContactEditor/wiki/Class-Module-Design-Convention
[DataTableModel]: https://pchemguy.github.io/ContactEditor/data-model#datatablemodel
[VBAArrayLib]: http://cpearson.com/excel/vbaarrays.htm