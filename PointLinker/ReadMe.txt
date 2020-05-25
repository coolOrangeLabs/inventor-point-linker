================================================
         Plugin of the Month, July 2011
================================================
------------
Point Linker
------------

Description
-----------
Inventor can import point data from an Excel file, but once the 
data has been imported, the link with the Excel file is not
retained. If the user then updates the point data in the Excel file,
the data must be manually re-imported and the associated model
re-constructed. This is a time consuming and isnefficient exercise. 

The Point Linker for Inventor tool has been developed to allow
Inventor users to import points from an Excel file, and retain a link
to that file. The users can import data (as with the existing
command),  update with new data, or resolve to another Excel file. 


System Requirements
-------------------
The APIs used by the plugin the are only supported from Inventor
version 2009 onwards (currently 2009, 2010, 2011, 2012).

An installer package which works for both 32-bit and 64-bit systems
is provided.

The source code has been provided as a Visual Studio 2008 project
containing VB.NET code (not required to run the plugin).

Installation
------------
To install the add-in on a client machine, run the add-in setup. 
Follow the instructions from the setup wizard. You will only need to
select the installation folder for the add-in. There is no
particular restriction concerning the location of this installation
folder.

On 32-bit platforms, building the source project inside Visual Studio
should cause the plugin to be registered, providing the application
has been "enabled for COM Interop" via the project settings.

On 64-bit platforms, the plugin cannot be registered by Visual Studio
2008 because it is a 32-bit application. Registration must either be
performed via the installer or a batch file/command prompt using:

RegAsm.exe /codebase "addin_path"

The plug-in provides three UI options for the new command:
(a) in Sketch tab, either put the command alongside the existing
     Inventor command 'Points', or
(b) replace the existing command with the new PointLinker plug-in
     command, or
(c) in the Tools tab, add one isolated command (Part or Assembly
     documents only)

Usage
-----
Once loaded, the command "Link Points" will be available:
- in the Sketch command bar or Tools menu of the classic UI
- in the Sketch panel or the Tools tab of the Ribbon UI

When you run the "Link Points" command. A dialog will appear. It
allows you to manage the points imported into the sketch.


Uninstallation
--------------
You can unload the plugin without uninstalling it by unchecking the
"Load" checkbox associated with the plugin in the Inventor Add-In
dialog.

Unchecking "Load on Startup" causes the plugin not to be loaded in
future sessions of Inventor.

To remove the plugin completely, uninstall the application via your
system's Control Panel.

Known Issues
------------

The installer - while working on both 32- and 64-bit systems - was
developed using 32-bit technology and therefore suggests an
inappropriate default installation location on 64-bit systems:

i.e. "C:\Program Files (x86)" instead of "C:\Program Files".

Although this will work fine, it's adviseable to install the
application into the 64-bit Program File section of your system.


Possible Future Enhancements
----------------------------
1) Treat the excel file as part of the dependency hierarchy, and
    display a warning on open if the Part/Assembly is out of date
	with respect to the Excel file.
2) When the plug-in dialog is running, and the user is updating the
    Excel file simultaneously,the plug-in can respond to the update
	immediately, without the user having to re-open the dialog.  

Note For Developers
-------------------
The installer project PointLinkerSetup uses one custom dialog
template which was defined by the author. The template file is
enclosed with the package:
PointLinker\PointLinkerSetup\VsdInventorInstallerDlg.wid. 
Before you re-build PointLinkerSetup, please put the file into the 
C:\Program Files\Microsoft Visual Studio 9.0\Common7\Tools\
  Deployment\VsdDialogs\1033 folder.
If you are using Visual Studio 2008, '1033' means the English
version.  For more information on *.wid, please refer to MSDN help.


Further Reading
---------------
The "Point Linker for Inventor Addin - Quick Start" guide in the
"Doc" folder provides further details about the use of this plugin.

Feedback
--------
Email us at support@coolorange.com with feedback or requests for
enhancements.

Release History
---------------

1.0 Original release


(C) Copyright 2011 by coolOrange s.r.l