---
layout: default
title: Class design convention
nav_order: 2
permalink: /class-design
---

*Parametrized class instantiation*  
A pair of a factory and a custom constructor performs parametrized class instantiation. Default factory *Create* and default constructor *Init* are defined on the class's default interface only. Both methods have the same parameter signature but different return values. Factory should be a function returning a class instance, and constructor should be a sub with no return value. The factory method called on the default (predeclared) class's instance (enabled via the "Predeclared" attribute) generates a new object and then calls its constructor with all received arguments to complete initialization. To simulate rudimentary introspection, *Class* and *Self* getters can also be defined. The *Class* getter returns the class's default instance. If a class instance presents a non-default interface, *Self* should return it as well.

*Abstract Factory*  
An abstract factory class has two default factory methods. The *Create* method follows the same convention as for a regular class. It is available on the default predeclared instance of the abstract factory only and generates factory instances. The other factory is *CreateInstance*. It must be available on non-default factory instances, but it can also be available on the default instance. This factory generates instances of the target class.

*Multiple interfaces*  
Class interfaces enable better decoupling of components. Some classes implement multiple interfaces (analog of multiple inheritance). While their instances expose the implemented interfaces, a given reference exposes only one interface at a time. One possible approach to making all interfaces conveniently available is as follows. "Self\<Interface name\>" methods defined on the class's default interface return references to the instance's respective interfaces. Additionally, the *CreateDefault* constructor defined on the default interface generates class instances exposing the default interface.