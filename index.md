# Contact Editor

## Acknowledgments

This project is largely an exercise for me aimed at both learning and making VBA templates for the Model-View-Presenter (MVP) pattern backed by persistent storage. A special thanks goes to Mathieu Guindon, a co-founder of the [Rubber Duck VBA][Rubber Duck VBA] project and his [RDVBA blog][RDVBA blog]. Initially, I followed the [post][RDVBA No Worksheet] describing a possible approach to abstracting Worksheet-based persistent storage, as well as a template provided in the comments. Eventually, I started from scratch borrowing some code and design elements.

## Overview

Contact Editor is NOT a real application. Rather, it is a template and a demo, targeting the User <---> Database interaction workflow. Such a workflow would typically involve the user sending a query to the database, receiving a response table, browsing/editing record data with user form, and, possibly, updating the database, as shown in the figure below.

![Overview][Overview]

## Development environment and repo structure

Contact Editor demo is a VBA app, and it lives within an Excel Workbook [Contact Editor.xls][Contact Editor] available from the root of this repo. Additionally, all code modules and UserForms are available from the [Project][Project] folder (which acts as a container and corresponds to the VBA project root within the .xls file). The development process is greatly facilitated by the [Rubber Duck VBA][Rubber Duck VBA] add-in, and the project structure is exported/imported using [RDVBA Project Utils][RDVBA Project Utils] VBA module. Primarily, I use Excel 2002 for development and also run tests on Excel 2016.

## Running the demo with different backends

The main entry for the demo is `RunContactEditor` from \ContactEditor\Forms\Contact Editor\ContactEditorRunner.bas module. Basic backend configuration is a part of `ContactEditorPresenter.InitializeModel` (\ContactEditor\Forms\Contact Editor\ContactEditorPresenter.cls). The `DataTableBackEnd` variable found in the entry point sub determines the type of the main storage backend. Three different backends have been implemented, pulling data from an Excel Worksheet ("Worksheet"),  a text delimited file ("CSV"), and the most recent addition - ADODB backend ("ADODB"), which pulls data from a generic database. Both CSV and Worksheet based databases can be processed through the ADODB backend with appropriate configuration, but its primary purpose is to abstract relational database management systems. The latter is illustrated with a mock [SQLite database][ContactEditor.db], provided in the root of the repo.

## Documentation and further information

Documentation and technical details are available from the project [WiKi][WiKi].


[Rubber Duck VBA]: https://rubberduckvba.com
[RDVBA blog]: https://rubberduckvba.wordpress.com
[RDVBA No Worksheet]: https://rubberduckvba.wordpress.com/2017/12/08/there-is-no-worksheet
[RDVBA Project Utils]: https://github.com/pchemguy/RDVBA-Project-Utils
[Overview]: https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Overview.png
[Contact Editor]: https://github.com/pchemguy/ContactEditor/blob/master/ContactEditor.xls
[Project]: https://github.com/pchemguy/ContactEditor/tree/master/Project
[ContactEditor.db]:  https://github.com/pchemguy/ContactEditor/blob/master/ContactEditor.db
[WiKi]: https://github.com/pchemguy/ContactEditor/wiki
