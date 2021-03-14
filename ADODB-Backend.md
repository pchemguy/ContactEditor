### Overview

While there is a plan to integrate [SecureADODB][SecureADODB] or its [fork][SecureADODB fork], a plain ADODB backend has been added to the project. Apart from the main class module, DataTableADODB.cls, two additional modules, ADOlib and SQLlib have also been added, with one helper routine added to CommonRoutines.bas module.

**DataTableADODB.cls**, like all other backends implements IDataTableStorage interface, and also contains several backend specific routines supporting the interface code. In order to make these supporting routines accessible to unit testing all such routines have been declared public and the [hybrid interface pattern][Hybrid Interface] has been implemented.

Present implementation of the DataTableADODB is fairly basic and it assumes that its instance queries one particular table supplied to the class factory. During instantiation, the factory runs a basic wild card select query against the target table requesting one record. The returned ADODB.Recordset object is then used for table introspection. Its Fields collection property is enumerated, and `FieldNames`, `FieldTypes`, and `FieldMap` (renamed FieldIndicies, [see][DataTableModel]) of a DataTableADODB instance are populated by the `CollectTableMetadata` routine.  

`AdoCommand` field stores a reusable reference to ADODB.Command object. This property is set by the `AdoCommandInit` method, which optionally takes an SQL query (basic SELECT is used by default) and *CursorLocation* (default - client side). The minimum set of properties necessary for execution are initialized, including *ConnectionString* and *SQLQuery* are set. After initialization, command can be used multiple times either as the source for an ADODB.Recordset or on its own for queries not returning data until it is changed by another call to AdoCommandInit.  

`AdoRecordset` optionally takes an *SQLQuery* and, if provided, calls AdoCommandInit, then executes the query and retunes data as Recordset.  

`Records` method is similar to AdoRecordset, except that it returns a row-wise 2D Variant array, whereas `RecordsAsText` additionally requests the backend to cast all fields as text.

**ConnectionString helpers**. DataTableADODB checks if ConnectionString has a certain prefix (for now just for SQLite. If so, it calls a corresponding helper routine from ADOlib that builds certain default connection string. GetSQLiteConnectionString calls VerifyOrGetDefaultPath, which in turn checks if provided argument is a valid file path. If not, by default, it looks for a file with the same name as the Workbook, in the same folder and tries provided extensions. If such file is found, its path is returned.

**Update query helpers**. Two other routines from ADOlib, *MakeAdoParamsForRecordUpdate* and *RecordToAdoParams* are used for performing database update. 

*MakeAdoParamsForRecordUpdate* take a list of FieldNames, FieldTypes, and an ADODB.Command. It constructs a basic single record update query, assuming that the first field is the primary key used in the WHERE clause. The update query is parametrized with respect to all field values, and for each field a parameter objected is created and added to ADODB.Command.Parameters. Then this command can be used to update multiple records by setting the values of Parameter members to the corresponding field values. *RecordToAdoParams* takes a record dictionary and updates parameters by matching the names of the fields with names of parameters. *IDataTableStorage_SaveDataFromModel* interface from DataTableADODB take the list of dirty records from the table model, it uses its routine that copies individual records to a dictionary object, then it calls the helper to update parameter values, and executes update query. This process loops through the dirty record list, reusing the same *prepared* command, and the loop is placed inside a transaction.

*SQLlib* module is used as a query provider and, at present, contains several basic SELECT queries and an UPDATE query. Optionally, an AS TEXT type casting can be added to the query, to that the data returned in the form of a Variant array contains numeric fields cast as text. Otherwise, Excel's Transpose function presently used to transpose column wise data from the Recordset, casts integer fields as Double. This is conversion is problematic since the demo uses dropdown prepopulated lists for Id and Age fields, and type mismatch between the allowed values list and initial value from the record caused undesirable glitches.


[SecureADODB]: https://github.com/rubberduck-vba/examples/tree/master/SecureADODB
[SecureADODB fork]: https://github.com/pchemguy/RDVBA-examples
[Hybrid Interface]:  https://github.com/pchemguy/ContactEditor/wiki/Class-Module-Design-Convention
[DataTableModel]: https://github.com/pchemguy/ContactEditor/wiki/Data-Model#datatablemodel