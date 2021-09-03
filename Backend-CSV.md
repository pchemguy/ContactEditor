---
layout: default
title: CSV
nav_order: 3
parent: Backends
permalink: /backends/csv
---

Excel VBA provides several functionalities for parsing delimiter-separated values text files:
- basic file I/O and the Split function (basic parser);
- Workbook.Open/Workbook.OpenText functions;
- Excel's QueryTable.

The basic parser is less flexible but at least 20x faster than Workbook.Open/Workbook.OpenText. Thus, the CSV backend employs the basic parser (adapted from [CSVParser][]).

The standard factory has the following signature:

```vb
Public Function Create(ByVal Model As DataTableModel, _
					   ByVal ConnectionString As String, _
					   ByVal TableName As String) As DataTableCSV
End Function
```

*ConnectionString* should be the path of the file. If blank, *ThisWorkbook.Path* will be used. *TableName* should contain a file name with an optional suffix "!sep=," (replace the comma with the actual single-character field separator. If blank, `ThisWorkbook.VBProject.Name & ".xsv!sep=,"` will be used. The default separator is the comma.

While performance tests using the CSV parser code demonstrated good performance of this approach, running Contact Editor with the CSV backend reveals a performance issue. To pinpoint the cause of this problem a further investigation is necessary.

[CSVParser]: https://github.com/pchemguy/CSVParser
