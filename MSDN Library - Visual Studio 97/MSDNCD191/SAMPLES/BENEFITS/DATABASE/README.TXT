README.TXT for Corporate Benefits Sample [Database Component]
Visual Studio '97

Why Two Types Of Database Files?
---------------------------------
The sample will work with either an ODBC or an MDB database.
To use an ODBC database, Microsoft SQL Server version 6.5 must be installed.

How Do I Prepare The Microsoft Access Database?
------------------------------------------------
*Use or copy db\benefits.mdb for the MDB mode. Check the readonly property for write access.
	
How Do I Prepare The Microsoft SQL Server Database?
----------------------------------------------------
* Create new database on your server. Make sure that isql is in your path. Run these scripts on the server machine. You also need rights on the master database before running these scripts.

* Open up a command prompt to run the build scripts from.

* Run 'build1.bat dbName' (see NOTE below) - where dbName is the name of the database created in the above step - to create the table and support database objects.

* Run 'build2.bat dbName' (see NOTE below) - where dbName is the name of the database created in the above step - to create the stored procedures. 

* Run 'build3.bat dbName' (see NOTE below) - where dbName is the name of the database created in the above step - to load the sample data.

* Set up an ODBC data source with the ODBC Data Source Administrator (this must be done on the client).

NOTE: Use the buildXd.bat files to direct output to directory level log files instead of screen display.  To check for error messages search for the string "MSG" in all build.log files. 
