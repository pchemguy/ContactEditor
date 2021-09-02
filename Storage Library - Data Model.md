---
layout: default
title: Data model
nav_order: 1
parent: Storage Library overview
permalink: /storage-library/data-model
---

The [Data Container][Data manager application figure] component consists of three model classes. The high-level *ContactEditorModel* class is a part of the [MVP][Data manager application figure] component. *DataRecordModel* and *DataTableModel* are the two base model classes, which are part of the [Storage Library][Data manager application figure]. *ContactEditorModel* incorporates the two base classes via indirect composition through the *DataCompositeManager* class, as illustrated later in this section.

Storage backends comprise the other important part of the library. These classes encapsulate details of individual storage types. Each backend type, such as an Excel worksheet or a CSV file, has a corresponding backend class responsible for transferring the data between the model class and the respective persistent storage. Backend classes implement either *IDataRecordStorage* or *IDataTableStorage* interface providing a backend-independent means for retrieving and saving the data.

*DataRecordModel* holding a single record (data table row) is behind the "record editor" user form. This model class has one backend, *DataRecordWSheet*, which saves the data to Excel Worksheet and populates the model at application startup. Record backends implement the *IDataRecordStorage* interface, and the *DataRecordFactory* abstract factory, implementing the *IDataRecordFactory* interface, instantiates record backends. See [Fig. 1](#FigDataRecord).

<a name="FigDataRecord"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Class Diagram - Record.svg" alt="FigDataRecordModel" width="100%" />  
<p align="center"><b>Figure 1. DataRecord class diagram</b></p>  

*DataTableModel* represents a whole data table or a subset of rows, abstracting persistent storage. This model class has three backends, including *DataTableWSheet*, *DataTableCSV*, and *DataTableADODB*. They handle a Worksheet, a delimited text file, and a relational database, respectively. Table backends implement the *IDataTableStorage* interface, and the *DataTableFactory* abstract factory, implementing the *IDataTableFactory* interface, instantiates table backends. See [Fig. 2](#FigDataTable).

<a name="FigDataTable"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Class Diagram - Table.svg" alt="FigDataTableModel" width="100%" />  
<p align="center"><b>Figure 2. DataTable class diagram</b></p>  

Data Managers comprise the last part of the library. Each such class incorporates by composition a model class and an appropriate backend class, yielding "backend-managed" models. Applications using this library should typically instantiate and keep a reference to one of these classes. The library defines three manager classes, including two simple and one composite.

*DataRecordManager* ([Fig. 1](#FigDataRecord)) and *DataTableManager* ([Fig. 2](#FigDataTable)) implement *IDataRecordManager* and *IDataTableManager* interfaces and manage their respective model classes. Additionally, since *DataRecordModel* and *DataTableModel* may work cooperatively (one holding a record from the recordset held in the other), the data may need to be transferred between the two model classes. Thus, a third composite manager is required.

*DataCompositeManager* is used where *DataRecordModel* and *DataTableModel* work together, and it handles data transfers between the model classes. A composite manager can encapsulate either two backend-managed classes or two model and two backend classes directly. *DataCompositeManager* uses the latter option ([Fig. 3](#FigCompositeManager)).

<a name="FigCompositeManager"></a>  
<img src="https://github.com/pchemguy/ContactEditor/blob/master/Assets/Diagrams/Class Diagram.svg?raw=true" alt="Overview" width="100%" />  
<p align="center"><b>Figure 3. Composite manager</b></p>  

Lets us overlap the structural application schematic from the previous section and the class hierarchy, as shown in [Fig. 4](#FigFunctionalClassMapping). Note that the figure has only one DataTable backend class, one Data Manager class, while abstract factory classes are not shown. The arrows indicate the flow of data, and the classes placed near the middle of these arrows facilitate/implement such transfers (*ContactEditorPresenter* is responsible for "GUI&nbsp;&#x21D4;&nbsp;*DataRecordModel*" transfer). The figure also shows how the MVP and the Storage Library components are connected.

<a name="FigFunctionalClassMapping"></a>  
<img src="https://github.com/pchemguy/ContactEditor/blob/master/Assets/Diagrams/Overview Class Map.svg?raw=true" alt="Functional class mapping" width="100%" />  
<p align="center"><b>Figure 4. Functional class mapping.</b></p>  



<!-- References -->

[Data manager application figure]: https://pchemguy.github.io/ContactEditor/#FigDataManagerApp
