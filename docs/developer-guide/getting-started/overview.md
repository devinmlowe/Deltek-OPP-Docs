# Deltek Open Plan® Developer's Guide - Overview

*Last Updated: December 20, 2020*

## About Deltek Open Plan

With Deltek Open Plan's project management tools, you can provide your organization with reliable estimates and solutions for long and short-term resource and budget allocations. Simplify manual processes with rapid data entry, quickly share and update project plan information through Project Portfolio Management (PPM) integration, and ensure projects get the resources they need.

## Integration Capabilities

Open Plan provides powerful integration capabilities through [OLE Automation](ole-automation-intro.md), allowing you to:

- Extend Open Plan's capabilities with custom solutions
- Integrate with other Windows applications
- Automate repetitive tasks
- Build custom workflows

## Key Concepts

### Objects

The fundamental unit of data or functionality exposed by Open Plan. Objects can include:

- The Open Plan application
- Projects or collections of projects
- Activities or collections of activities
- Resource definitions or assignments
- Project codes and code files
- Calendar information

Learn more: [Understanding Objects](../objects/README.md)

### Collections

Some objects are collections of other objects. Collection names are usually the plural form of the objects they contain. Examples:

- A collection of activity records
- A collection of barchart views
- A collection of resource description records

### Properties

Each object has a set of properties that represent specific attributes. Properties can be:

- Field values in an activity record
- The number of activities in a project
- The state of a window or view

Learn more: [Properties Reference](../properties/README.md)

### Methods

Methods are actions that you instruct objects to perform. Available methods include:

- Adding an activity
- Performing time analysis
- Displaying a view
- Opening a project
- Minimizing a window
- Setting a filter

Learn more: [Methods Reference](../methods/README.md)

## Getting Started

1. [Introduction to OLE Automation](ole-automation-intro.md)
2. [Accessing OLE Automation Objects](accessing-objects.md)
3. [Understanding the Object Hierarchy](../object-hierarchy/README.md)

## Modifying Existing Applications for Open Plan 8.x

If you have applications designed for earlier versions of Open Plan, you may need to make modifications for compatibility with Open Plan 8.x. Key changes include:

### Objects, Properties, and Methods No Longer Supported

Objects not supported:
- OPFCProcesses
- OPFCStructures

Properties no longer supported:
- OPPFormat
- ResidueFile

Methods not supported:
- FileOpenSpecialOpenPlan
- OpenSpecialODBC
- GlobalDir
- Rights
- Processes
- Structures

### Naming Conventions

Applications using early binding need to be modified to use the new application object's name. `OPCreateApplication` is also referred to as `OPApplication` object in this guide.

### Security Features

A security feature added to Open Plan 8.x requires users to login upon launching the application. Automation applications that launch Open Plan must call the [Login method](../methods/opcreateapplication33.md#login) immediately after instantiating the Open Plan application.

If a user's security permissions don't permit the use of a given feature, that feature will not be available for use by automation applications. Instead of disabling tab commands in Deltek EPM Security Administrator, set the menu items to invisible and leave them enabled. The user will not be able to access them, but automation applications will.

### Add-Ins Tab

In Open Plan 8.x, more parameters are available for applications added to the Add-Ins tab. These parameters can return context-specific data to your application or control how an Add-Ins tab tool is launched.

The data returned by Open Plan's special command line parameters in the Add-Ins tab is wrapped in quotes. You may need to modify your applications to handle this.

> **Note:** For a list of command line parameters and information about the Add-Ins tab, see the Open Plan Configuration Files documentation.

## See Also

- [Object Hierarchy Diagrams](../object-hierarchy/README.md)
- [Complete Objects Reference](../objects/README.md)
- [Properties Reference](../properties/README.md)
- [Methods Reference](../methods/README.md)
