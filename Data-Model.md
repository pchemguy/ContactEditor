At the basis of the project model are two data data model classes, `DataRecordModel`, representing a single Record (table row) shown to the user on the UserForm, and `DataTableModel`, abstracting persistent storage. Model classes know nothing about data storage, which is the responsibility of the backend classes.  

Presently, there is one backend for DataRecordModel, `DataRecordWSheet`, which saves the last saved record to Excel Worksheet, populating the UserForm with this record at startup.

DataTableModel has three backends, `DataTableWSheet`, `DataTableCSV`, and `DataTableADODB`, responsible for handling a Worksheet, a delimited text file, and a generic RDBMS respectively.  
 
The backends correspondingly implement either IDataRecordStorage or IDataTableStorage interface are instantiated by their respective abstract factories: DataRecordFactory\IDataRecordFactory and  DataTableFactory\IDataTableFactory. In turn, DataRecordManager\IDataRecordManager and DataTableManager\IDataTableManager incorporate by composition one model and one backend class, yielding a "backend-managed" model.

![Base classes][Base classes]

DataRecordModel and DataTableModel work in concert, with DataRecordModel holding a piece of information from the DataTableModel. Data chunks need to travel between the two models; hence, the need for additional functionality, which should a part of a derived manager. Thus, `DataCompositeManager` class has been added. In VBA, composition is used for such purposes, and the composite manager can be implemented either via a composition of "backend-managed" classes, or via a composition of two model and two backend classes directly. The latter pathway has been chosen for  DataCompositeManager class, and it exposes the necessary features of the constituent classes and implements the "inter-model" functionality. Finally, the viewModel class ContactEditorModel incorporates DataCompositeManager.

![Composite classes][Composite classes]


[Composite classes]: https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram.svg
[Base classes]: https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Class%20Diagram%20-%20Table%20and%20Record.svg
