The following are the general class design conventions partly following those of [RubberDuck VBA examples][RubberDuck VBA examples].

**Custom class Factory/Constructor**. When a class employs parameterized instantiation, Predeclared attribute of such a class is set to true, and the class provides `Create` factory used to generate parameterized instances. A custom `Init` constructor with the same arguments signature is defined and is responsible for initializing the instance (`Create` factory is called on the default class instance, and it does not have access to the private fields of the newly generated instance). Often a `Class` property getter is defined, which returns a reference to the default (Predeclared) class instance. A `Self` getter returns a reference to the object itself. When a class implements a single interface, Self typically returns a reference of such an interface, and so does the `Create` factory.

**Abstract Factory**. When an Abstract Factory pattern is employed, `Create` factory called on the default instance of the AbstractFactoryClass generates a factory instance AbstractFactoryInstance, and `CreateInstance` method called on a factory instance generates an instance of the target class.

[](#hybrid-interface)
**Hybrid interface**. Some classes implementing a custom interface, may also provide specific functionality not exposed via such an interface, and, thus, only accessible via the default interface. When such functionality needs to be exposed on the same instance along with the custom interface, a second factory `CreateDefault` is added, which returns the default interface reference. Additionally, a getter `Self<Interface name>` is defined on the default class interface, returning a reference of the custom interface. In principle, multiple interfaces maybe exposed this way.


[RubberDuck VBA examples]: https://github.com/rubberduck-vba/examples