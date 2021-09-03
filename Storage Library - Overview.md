---
layout: default
title: Overview
nav_order: 1
parent: Storage Library
permalink: /storage-library/overview
---

### Functional structure

[Fig. 1](#FigDataFlow) illustrates a simplified data flow process in a basic data manager application. Storage Library (StoreLib), shown in the center of the diagram, mediates data transfer between the user (UI) and storage (Database). StoreLib, in turn, contains two core class families, *DataRecord* and *DataTable*. The user (via the UI) may work with a single data chunk at time (e.g., a record). In StoreLib, current data chunk is handled by the *DataRecord* family, which directly interacts with the UI and can save a copy locally in a file. The other StoreLib family, *DataTable*, interacts directly with the data storage and holds multiple chunks of data (e.g., a set of records).

<a name="FigDataFlow"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Data Flow Diagram.svg" alt="FigDataFlow" width="100%" />  
<p align="center"><b>Figure 1. Basic data flow diagram</b></p>  

<details><summary>Data management flow chart</summary>

<a name="FigDataManagementFlowChat"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Data Manager Flow Chart.svg" alt="FigDataManagementFlowChat" width="100%" />  
<p align="center"><b>Figure 2. Data management flow chart</b></p>  

</details>

First, the user requests certain data, e.g., a table. These data is loaded into the *DataTable* family, and one chunk, e.g., a record is loaded from *DataTable* into *DataRecord* and presented to the user. Next, whenever user requests a different chunk of data available from *DataTable*, confirmed edits to the current chunk are copied back into this chunk is copied from *DataTable* into *DataRecord*

### Data model

*DataRecordModel* holding a single record (data table row) is behind the "record editor" user form. This model class has one backend, *DataRecordWSheet*, which saves the data to Excel Worksheet and populates the model at application startup. Record backends implement the *IDataRecordStorage* interface, and the *DataRecordFactory* abstract factory, implementing the *IDataRecordFactory* interface, instantiates record backends. See [Fig. 3](#FigDataRecord).

<a name="FigDataRecord"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Class Diagram - Record.svg" alt="FigDataRecordModel" width="100%" />  
<p align="center"><b>Figure 3. DataRecord class diagram</b></p>  

*DataTableModel* represents a whole data table or a subset of rows, abstracting persistent storage. This model class has three backends, including *DataTableWSheet*, *DataTableCSV*, and *DataTableADODB*. They handle a Worksheet, a delimited text file, and a relational database, respectively. Table backends implement the *IDataTableStorage* interface, and the *DataTableFactory* abstract factory, implementing the *IDataTableFactory* interface, instantiates table backends. See [Fig. 4](#FigDataTable).

<a name="FigDataTable"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Class Diagram - Table.svg" alt="FigDataTableModel" width="100%" />  
<p align="center"><b>Figure 4. DataTable class diagram</b></p>  

Data Managers comprise the last part of the library. Each such class incorporates by composition a model class and an appropriate backend class, yielding "backend-managed" models. Applications using this library should typically instantiate and keep a reference to one of these classes. The library defines three manager classes, including two simple and one composite.

*DataRecordManager* ([Fig. 3](#FigDataRecord)) and *DataTableManager* ([Fig. 4](#FigDataTable)) implement *IDataRecordManager* and *IDataTableManager* interfaces and manage their respective model classes. Additionally, since *DataRecordModel* and *DataTableModel* may work cooperatively (one holding a record from the recordset held in the other), the data may need to be transferred between the two model classes. Thus, a third composite manager is required.

*DataCompositeManager* is used where *DataRecordModel* and *DataTableModel* work together, and it handles data transfers between the model classes. A composite manager can encapsulate either two backend-managed classes or two model and two backend classes directly. *DataCompositeManager* uses the latter option ([Fig. 5](#FigCompositeManager)).

<a name="FigCompositeManager"></a>  
<img src="https://github.com/pchemguy/ContactEditor/blob/master/Assets/Diagrams/Class Diagram.svg?raw=true" alt="Overview" width="100%" />  
<p align="center"><b>Figure 5. Composite manager</b></p>  

Lets us overlap the structural application schematic from the previous section and the class hierarchy, as shown in [Fig. 6](#FigFunctionalClassMapping). Note that the figure has only one DataTable backend class, one Data Manager class, while abstract factory classes are not shown. The arrows indicate the flow of data, and the classes placed near the middle of these arrows facilitate/implement such transfers (*ContactEditorPresenter* is responsible for "GUI&nbsp;&#x21D4;&nbsp;*DataRecordModel*" transfer).

<a name="FigFunctionalClassMapping"></a>  
<img src="https://github.com/pchemguy/ContactEditor/blob/master/Assets/Diagrams/Overview Class Map.svg?raw=true" alt="Functional class mapping" width="100%" />  
<p align="center"><b>Figure 6. Functional class mapping.</b></p>  

The *ContactEditorModel* (MVP) and *DataRecordModel*/*DataTableModel* (Storage Library) classes act as a container for the data. As illustrated in the figure, the former encapsulates the latter via indirect composition through the DataCompositeManager class. The four classes thus bridge the MVP and Storage Library components.
