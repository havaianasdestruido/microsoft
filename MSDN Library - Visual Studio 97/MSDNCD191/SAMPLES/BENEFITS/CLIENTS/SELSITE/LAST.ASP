<!-- This page writes records of the employee's family to the database. . 
     Copyright 1996 by Microsoft Corporation. Written in Developer Studio. -->
<html>
<head><title>Using your benefits</title>

<link rel=stylesheet href="hr.css" type="text/css">

</head>

<body bgcolor=FFFFFF>
<CENTER><IMG SRC="images\ben_ani2.gif" WIDTH=300 HEIGHT=100>
<BR>
<P><P>
<h2>Thank you for visiting!</h2>
<% 
REM Delete the old records from the database;
PASS = "DELETE FROM tbDependant WHERE tbDependant.EmployeeID="  & Request.Form("empid")
varLong = 0
Session("DBConn").Execute(PASS), varLong, 1

 
REM Write all the records from the last page to the database.
REM We're not assuming they're contiguous.
set AddRecord = Server.CreateObject("ADODB.Recordset")
AddRecord.Open "SELECT * FROM tbDependant WHERE EmployeeID = " & Session("empid"), Session("DBConn"), 1, 3, 1
Dim FieldNames(6)
FieldNames(0) = "EmployeeID"
FieldNames(1) = "FirstName"
FieldNames(2) = "LastName"
FieldNames(3) = "DepTypeID"
FieldNames(4) = "DateOfBirth"
FieldNames(5) = "Gender"
FieldNames(6) = "Status"

Dim FieldValues(6)
FieldValues(0) = Request.Form("empid")

For i = 1 to Request.Form("howmany") 
	If (Request.Form( i & "FName") <> "") then
		varLong = 0
		FieldValues(1) = Request.Form( i & "FName")
		FieldValues(2) = Request.Form( i & "LName")
		FieldValues(3) = Request.Form( i & "Type")
		FieldValues(4) = Request.Form( i & "DateofBirth")
		FieldValues(5) = Request.Form( i & "Sex") 
		FieldValues(6) = Request.Form( i & "Status")
		AddRecord.AddNew FieldNames,FieldValues
	end if
Next
%>
<BR>Your benefit selections for 1997 have been recorded.

</CENTER>

</body></html>
<% Session.Abandon %>