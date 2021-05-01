---
layout: default
title: Worksheet
nav_order: 1
parent: Backends
permalink: /backends/worksheet
---

*DataTableWSheet* uses the "Contacts" worksheet filled with mock data. The *ConnectionString* parameter supplied to the backend factory should be in the form "\<Thisworkbook&#x2E;name\>!\<worksheet name\>" (e.g., "ContactEditor\.xls!Contacts"). The backend expects the table name as a globally scoped named range, so the worksheet name is not necessary for further processing. The first row of the table range should contain field names, and the first column should be the table's primary key ("id"). To simplify the VBA code, the backend expects two more named Ranges prefixed with the table name: a "Body" suffixed range containing only the data and an "Id" suffixed range with record IDs. 

For example, the named range "Contacts" contains the table, "ContactsBody" Range refers to the data area without the header, and "ContactsId" Range refers to the "id" column. Additionally, the "ContactsHeader" range refers to the header row only. (Note, because these ranges use the formula-based definition, the range name/address bar (top left corner) does not display their names.)

*DataRecordWSheet* uses the "ContactBrowser" worksheet to save the form data and populate it at the start. Each record field has an associated cell assigned a named range matching the field name. This convention makes it possible to use *FieldNames* from *DataTableModel*. Nevertheless, DataRecordWSheet collects the field names independently (to explore an alternative implementation). It goes through the *Names* collection of the Workbook object and keeps name members matching all of the following:

- Name.RefersTo starts with the name of the worksheet supplied to the *DataRecordWSheet* constructor;
- the target range is a single cell range;
- the target range has a "label" cell above or to the left containing the field name.

For example, a single cell range named "LastName" refers to ContactBrowser!E5, and ContactBrowser!D5 contains the same name.

The current implementation makes no attempts to handle unopened workbooks. The only confirmed tests ran on the file containing both the backend worksheet and the VBA code.