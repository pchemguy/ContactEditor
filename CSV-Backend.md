There are several ways to pull data from a text delimited file in VBA, e.g.:
- directly read the file into a string buffer and use Split function to split it into records and then into fields (basic parser)
- use Workbook.Open/Workbook.OpenText to open the file and copy the parsed data from the range
- using Excel's QueryTable.  
The first two are illustrated by [CSV Parser][CSV Parser]. Workbook.Open/Workbook.OpenText, while being more flexible then the basic parser, is at least 20x time slower. CSV backend is currently implemented via the first approach. While performance tests using the CSV Parser code demonstrated good performance of this approach, running Contact Editor with CSV backend show performance issue. The nature of this problem is not clear at present and would require further investigation.

[CSV Parser]: https://github.com/pchemguy/CSVParser