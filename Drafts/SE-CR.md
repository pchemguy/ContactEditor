[ContactEditor][ContactEditor] uses Excel VBA as a prototyping platform for a personal information manager combining the "Model, View, Presenter" (MVP) pattern and a persistent storage manager. The design of this demo app was inspired by [RubberDuck VBA][]. I attempted to incorporate the ideas/coding patterns/tutorials from RubberDuck VBA blog's multiple great posts (such as [OOP in VBA][] or [VBA Class Modules][]) and demo projects ([SecureADODB][] and [OOP Battleship][]). At the same time, RubberDuck and associated [SE-CR][] posts only briefly discuss the possible approaches to combining MVP with persistent storage management (for example, [There is no worksheet][], [Abusing Excel VBA… to maintain data stored in a database][Excel VBA DB], [YARPI: Yet Another Repository Pattern Implementation][YARPI], and [Down the rabbit hole with MVP][]). Hence, this app is also a hands-on exercise filling the blanks in my understanding of the MVP pattern backed by persistent storage.




[ContactEditor]: https://pchemguy.github.io/ContactEditor/

[RubberDuck VBA]: https://rubberduckvba.com/

[OOP in VBA]: https://rubberduckvba.wordpress.com/2015/12/24/oop-in-vba/
[VBA Class Modules]: https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
[SecureADODB]: https://rubberduckvba.wordpress.com/2020/04/22/secure-adodb/
[OOP Battleship]: https://rubberduckvba.wordpress.com/2018/08/28/oop-battleship-part-1-the-patterns/

[SE-CR]: https://codereview.stackexchange.com/

[There is no worksheet]: https://rubberduckvba.wordpress.com/2017/12/08/there-is-no-worksheet/
[Excel VBA DB]: https://codereview.stackexchange.com/questions/57734/abusing-excel-vba-to-maintain-data-stored-in-a-database
[YARPI]: https://codereview.stackexchange.com/questions/57889/yarpi-yet-another-repository-pattern-implementation
[Down the rabbit hole with MVP]: https://codereview.stackexchange.com/questions/58348/down-the-rabbit-hole-with-mvp


[MVP]: https://rubberduckvba.wordpress.com/2018/05/08/apply-logic-for-userform-dialog/
[UserForm1.Show]: https://rubberduckvba.wordpress.com/2017/10/25/userform1-show/
[Modeless UserForm]: https://stackoverflow.com/questions/47357708/vba-destroy-a-modeless-userform-instance-properly
[Events]: https://rubberduckvba.wordpress.com/2019/03/27/everything-you-ever-wanted-to-know-about-events/


[Factories]: https://rubberduckvba.wordpress.com/2018/04/24/factories-parameterized-object-initialization/
[Private this As TSomething]: https://rubberduckvba.wordpress.com/2018/04/25/private-this-as-tsomething/





[FigDataManagerApp]: https://raw.githubusercontent.com/pchemguy/ContactEditor/develop/Assets/Diagrams/Data%20Management%20Overview.jpg

[FigDataTable]: https://raw.githubusercontent.com/pchemguy/ContactEditor/develop/Assets/Diagrams/Class%20Diagram%20-%20Table.jpg

[FigCompositeManager]: https://raw.githubusercontent.com/pchemguy/ContactEditor/develop/Assets/Diagrams/Class%20Diagram.jpg
