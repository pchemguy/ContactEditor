Excel VBA provides several functionalities for parsing text delimited files:
- basic file I/O and the Split function (basic parser);
- Workbook.Open/Workbook.OpenText functions;
- Excel's QueryTable.

The basic parser is less flexible but at least 20x faster than Workbook.Open/Workbook.OpenText. Thus, the CSV backend employs the basic parser (adapted from [CSVParser]).

While performance tests using the CSV parser code demonstrated good performance of this approach, running Contact Editor with the CSV backend reveals a performance issue. To pinpoint the cause of this problem a further investigation is necessary.

[CSVParser]: https://github.com/pchemguy/CSVParser