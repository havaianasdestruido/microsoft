README.TXT for Corporate Benefits Sample [Visual Basic Client Application]
Visual Studio '97

What Are the Different Files for?
=================================
blocvb.exe   - Benefits Application that locally uses the BenVB.DLL Server. 
blocvbr.exe  - Benefits Application that locally uses the BenVBR.EXE Server.
bloccpp.exe  - Benefits Application that locally uses the BenCPP.DLL Server.
bloccppr.exe  - Benefits Application that locally uses the BenCPPR.EXE Server.


How Do I Run the Client Application?
====================================
To start, you MUST install the server objects.
See the README.TXT in the \benefits\servers\install directory to find out how to do this.
You must also set up the database and DSN as described in the README.TXT in the \benefits\db directory.

Once the servers have been installed, you can:
1) Run from the source files, or
2) Run the .EXE file

1) Running the client application from the source files
-------------------------------------------------------
a. Confirm that the following Components are available (using Visual Basics Projects | References dialog box). Do not select/de-select them.
     - Corporate Benefits Sample [VB Local Server]
     - Corporate Benefits Sample [VB Remote Server]
     - Corporate Benefits Sample [CPP Local Server]
     - Corporate Benefits Sample [CPP Remote Server]
   If the required entries are not listed, then the Company Benefits Server Components have not been installed correctly.

b. You can now run the client application from within the Visual Basic environment.

2) Installing and running the client application stand-alone
-------------------------------------------------------------
a. Click on the desired client in the Corporate Benefits program group.
b. Before using the client you should ensure that the database has been created and is available.


HR Benefits Sample Information
==============================

Running the Client Application
------------------------------
To login, use "admin" (without the quotes) to get to all of the functionality of the sample.
However, if you wish to view the application in "end-user" mode, log on using the user ID from any employee from the sample database.
