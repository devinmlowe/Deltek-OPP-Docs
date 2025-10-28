# Accessing OLE Automation Objects

This guide describes the steps and syntax for accessing Open Plan OLE Automation objects with Visual Basic. The same principles apply to other programming environments.

## Getting the Application Object

The first step for an external application to access objects belonging to Open Plan is to get the Open Plan application object.

### Visual Basic Example

```vb
'Declare the Open Plan Application Object variable
Dim OPCreateApplication33 As Object

'Assign the Application object variable
Set OPCreateApplication33 = CreateObject("opp.application")
```

> **Note:** In Visual Basic, a line of code that begins with an apostrophe is a comment.

## Working with Object Properties and Methods

After retrieving the Open Plan application object, you can access its properties and methods in the following ways:

### Getting an Object Property

A property of an object is retrieved by setting a variable equal to the object variable, followed by a period (.), and followed by a property name.

**Syntax:**
```vb
Variable = ObjectVariable.PropertyName
```

**Example:**
```vb
Dim MyActivities As Object
Dim ActivityCount As Long
ActivityCount = MyActivities.Count
```

### Setting an Object Property

A property of an object is set by setting the object variable, followed by a period (.), and followed by a property name equal to a variable that represents the values of the property.

**Syntax:**
```vb
ObjectVariable.PropertyName = Variable
```

**Example:**
```vb
Dim MyActivities As Object
Dim ActID As String
ActID = MyActivities.ID
```

### Executing Object Methods

An OLE Automation method can either use no parameters, or one or more parameters. It can return either nothing or a value. For methods that return a value, a variable can be assigned to the returning value by placing the variable on the left side of the equal sign (=). Some methods return objects. For these methods, the returned value is set to an object variable.

#### Methods with No Parameters

**Syntax:**
```vb
ObjectVariable.MethodName
```

**Example:**
```vb
MyCodeView.Minimize
```

#### Methods with One or More Parameters

**Syntax:**
```vb
ObjectVariable.MethodName Parameter1, Parameter2
```

**Example:**
```vb
MyProject.ResourceSchedule "Time Limited", True
```

#### Methods that Return a Value

**Syntax:**
```vb
Variable = ObjectVariable.MethodName(Parameter1, Parameter2)
```

**Example:**
```vb
SuppressInListsFlag = AResource.Getfield("NO_LIST")
```

#### Methods that Return an Object

**Syntax:**
```vb
ObjectVariable1 = ObjectVariable2.MethodName.Item(Parameter1)
```

**Example:**
```vb
Set MyCalendarFile = OPCreateApplication33.Calendars.Item("MYCAL")
```

## Complete Workflow Example

Here's a complete example of creating an Open Plan automation application:

```vb
'1. Declare the application object variable
Dim OPApp As Object

'2. Create the application object
Set OPApp = CreateObject("opp.application")

'3. Login (required for Open Plan 8.x+)
OPApp.Login

'4. Open a project
Dim MyProject As Object
Set MyProject = OPApp.FileOpen("C:\Projects\MyProject.opp")

'5. Get the activities collection
Dim Activities As Object
Set Activities = MyProject.Activities

'6. Loop through activities
Dim Activity As Object
Dim i As Long
For i = 1 To Activities.Count
    Set Activity = Activities.Item(i)
    'Do something with each activity
    Debug.Print Activity.ID & ": " & Activity.Description
Next i
```

## Important Notes for Open Plan 8.x

### Login Requirement

A security feature added to Open Plan 8.x requires users to login upon launching the application. Automation applications that are intended to launch Open Plan must call the [Login method](../methods/opcreateapplication33.md#login) immediately after instantiating the Open Plan application.

**Example:**
```vb
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login 'Must call Login before accessing other features
```

### Security Permissions

If a user's security permissions in Open Plan do not permit the use of a given feature, that feature will not be available for use by automation applications.

**Best Practice:** Instead of disabling tab commands in Deltek EPM Security Administrator, set the menu items to invisible and leave them enabled. The user will not be able to access them, but automation applications will.

## Next Steps

- [Understanding the Object Hierarchy](../object-hierarchy/README.md)
- [OPCreateApplication33 Reference](../objects/opcreateapplication33.md)
- [Working with Projects](../objects/opproject.md)
- [Working with Activities](../objects/opactivity.md)

## See Also

- [Object Reference](../objects/README.md)
- [Properties Reference](../properties/README.md)
- [Methods Reference](../methods/README.md)
