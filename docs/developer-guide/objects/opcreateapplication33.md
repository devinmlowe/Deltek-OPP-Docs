# OPCreateApplication33 Object

## Description

The OPCreateApplication33 object contains properties and methods for accessing the Open Plan application. This is the main entry point for all Open Plan automation.

## How to Get This Object

```vb
'Declare the Open Plan Application Object variable
Dim OPCreateApplication33 As Object

'Assign the Application object variable
Set OPCreateApplication33 = CreateObject("opp.application")

'Login (required for Open Plan 8.x+)
OPCreateApplication33.Login
```

## Properties

| Property | Description | Type |
|----------|-------------|------|
| [EnableUndo](../properties/enableundo.md) | Controls whether undo functionality is enabled | Boolean |
| [Height](../properties/height.md) | Height of the application window | Long |
| [SilentMode](../properties/silentmode.md) | Controls whether the application runs in silent mode | Boolean |
| [Width](../properties/width.md) | Width of the application window | Long |
| [XPosition](../properties/xposition.md) | X-coordinate of the application window | Long |
| [YPosition](../properties/yposition.md) | Y-coordinate of the application window | Long |

## Methods

### File Operations
- **[FileNew](../methods/filenew.md)** - Create a new project
- **[FileOpen](../methods/fileopen.md)** - Open an existing project
- **[FileOpenEx](../methods/fileopenex.md)** - Open a project with extended options
- **[FilePrint](../methods/fileprint.md)** - Print the active view
- **[FilePrintPreview](../methods/fileprintpreview.md)** - Show print preview
- **[FilePrintSetup](../methods/fileprintsetup.md)** - Show print setup dialog
- **[FileRename](../methods/filerename.md)** - Rename a file

### Import/Export
- **[GeneralExport](../methods/generalexport.md)** - Export data
- **[GeneralImport](../methods/generalimport.md)** - Import data

### Collections Access
- **[ActiveProject](../methods/activeproject.md)** - Get the currently active [OPProject](opproject.md) object
- **[ActiveView](../methods/activeview.md)** - Get the currently active [OPView](opview.md) object
- **[Calendars](../methods/calendars.md)** - Access the [OPCalendars](collections.md#opcalendars-collection) collection
- **[Codes](../methods/codes.md)** - Access the [OPCodes](collections.md#opcodes-collection) collection
- **[FileCabinet](../methods/filecabinet.md)** - Access the [OPFileCabinet](opfilecabinet.md) object
- **[Projects](../methods/projects.md)** - Access the [OPProjects](collections.md#opprojects-collection) collection
- **[ReportingCalendars](../methods/reportingcalendars.md)** - Access the [OPReportingCalendars](collections.md#opreportingcalendars-collection) collection
- **[Resources](../methods/resources.md)** - Access the [OPResources](collections.md#opresources-collection) collection
- **[WebWindows](../methods/webwindows.md)** - Access the [OPWebWindows](collections.md#opwebwindows-collection) collection

### Window Management
- **[Conceal](../methods/conceal.md)** - Hide the application window
- **[Disable](../methods/disable.md)** - Disable the application window
- **[Enable](../methods/enable.md)** - Enable the application window
- **[Maximize](../methods/maximize.md)** - Maximize the application window
- **[Minimize](../methods/minimize.md)** - Minimize the application window
- **[Restore](../methods/restore.md)** - Restore the application window
- **[Show](../methods/show.md)** - Show the application window
- **[WindowMinimizeAll](../methods/windowminimizeall.md)** - Minimize all open windows
- **[WindowTile](../methods/windowtile.md)** - Tile windows horizontally
- **[WindowTileVertical](../methods/windowtilevertical.md)** - Tile windows vertically

### Other Methods
- **[ApplyFilter](../methods/applyfilter.md)** - Apply a filter
- **[CreateBrowserView](../methods/createbrowserview.md)** - Create a new browser view
- **[GetLastSecurityValidation](../methods/getlastsecurityvalidation.md)** - Get last security validation result
- **[Login](../methods/login.md)** - Login to Open Plan (required for 8.x+)
- **[PageSetup](../methods/pagesetup.md)** - Show page setup dialog
- **[SetPrinter](../methods/setprinter.md)** - Set the printer
- **[SetResourceSelection](../methods/setresourceselection.md)** - Set resource selection
- **[SysDir](../methods/sysdir.md)** - Get the system directory path
- **[User](../methods/user.md)** - Get current user information
- **[Version](../methods/version.md)** - Get Open Plan version
- **[WorkDir](../methods/workdir.md)** - Get the working directory path

## Usage Example

```vb
'Create and configure the application
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")

'Login (required for Open Plan 8.x+)
OPApp.Login

'Get version information
Dim version As String
version = OPApp.Version
Debug.Print "Open Plan Version: " & version

'Open a project
Dim proj As Object
Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")

'Access the active project
Dim activeProj As Object
Set activeProj = OPApp.ActiveProject

'Show the application window
OPApp.Show
```

## Important Notes for Open Plan 8.x

### Login Requirement

A security feature added to Open Plan 8.x requires users to login upon launching the application. Automation applications that launch Open Plan **must** call the [Login](../methods/login.md) method immediately after instantiating the Open Plan application.

### Security Permissions

If a user's security permissions in Open Plan do not permit the use of a given feature, that feature will not be available for use by automation applications.

## See Also

- [Getting Started with OLE Automation](../getting-started/accessing-objects.md)
- [OPProject Object](opproject.md)
- [OPProjects Collection](collections.md#opprojects-collection)
- [Methods Reference](../methods/README.md)
