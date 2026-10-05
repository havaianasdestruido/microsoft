<!-- This page writes benefit choices to the database and gets records of the employee's family. 
     Copyright 1996 by Microsoft Corporation. Written in Developer Studio. -->

<html>
<head><title>Saving choices</title>

<link rel=stylesheet href="hr.css" type="text/css"></head>

<body bgcolor="FFFFFF"  topmargin=0 leftmargin=0>
<img src="images\h_fade.gif" height=75 width=800>
<% if ( Session("ActiveXControls") = -1 ) then
	LIFE = request.form("LifeCovText")
	DIS = request.form("DisCovText")
else
	LIFE = request.form("LifeCovDisp")
	DIS = request.form("DisCovDisp")
end if %>


<% 
REM First check to see if the text values are in permitted limits.
REM I've submitted the calculated values in the previous form, but if you expect network transmission to be
REM slower than database access you could requery the tbBenefit table and recalculate here.

TESTRET = ""
TESTLIFE = ""
TESTDIS = ""
SEP = "<LI>"

REM Cast to singles, or these will be alphabetic comparisons...

If ( CSng(request.form("RetDedText")) > CSng(Session("RetMax")) ) then	
	TESTRET = SEP & "your retirement deduction was too high"
end if
If ( CSng(request.form("RetDedText")) < 0 ) then
	TESTRET = SEP & "your retirement deduction can't be negative. <strong>good try.</strong>"
end if
If  ( (CSng(LIFE) <> 0) AND  (CSng(LIFE) < CSng(Session("LifeMin"))) ) then
		TESTLIFE = SEP & "your life insurance coverage was nonzero, but below the minimum"
end if
if ( CSng(LIFE) > CSng(Session("LifeMax")) ) then
		TESTLIFE = SEP & "your life insurance coverage was too high"
end if
If ( CSng(DIS) <> 0 ) then
	if ( CSng(DIS) < CSng(Session("DisMin")) ) then
		TESTDIS = SEP & "your disability coverage was nonzero, but below the minimum"
	elseif ( CSng(DIS) > CSng(Session("DisMax")) ) then
		TESTDIS = SEP & "your disability coverage was too high"
	end if
end if %>
<table  width=600 cellpadding=12>
<%if ( TESTRET <> "" or TESTLIFE <> "" or TESTDIS <> "" ) then %>
<tr><td>I'm sorry:<UL>
<%response.write  TESTRET & TESTLIFE & TESTDIS %>
</UL>Please go back to the last	page.</p>
</td></tr></table>
<%else%>
	<tr><td><img src="images/moment.gif" width=350 height=40></td></tr>

<%	REM Delete any previous choices made by this employee; 
	varLong = 0
	Session("DBConn").Execute "DELETE FROM tbEmplBnf WHERE tbEmplBnf.EmployeeId="  & Session("empid"), varLong, 1
	response.write "<!-- " & varLong & " records deleted-->"

	REM and add the new choices.

	set AddRecord = Server.CreateObject("ADODB.Recordset")
	response.write "<!-- Selecting records -->"
	AddRecord.Open "SELECT * FROM tbEmplBnf WHERE EmployeeID = " & Session("empid"), Session("DBConn"), 1, 3, 1

	Dim FieldNames(4)
	FieldNames(0) = "EmployeeId"
	FieldNames(1) = "BenefitId"
	FieldNames(2) = "CoverageId"
	FieldNames(3) = "Amount"
	FieldNames(4) = "TaxId"

	Dim FieldValues(4)
	FieldValues(0) = Session("empid")
	If ( Request.Form("HealthIns") <> "-1") then
		FieldValues(1) = Request.Form("HealthIns")
		FieldValues(2) = Request.Form("HealthCov")
		FieldValues(3) = 0
		FieldValues(4) = 2
		AddRecord.AddNew FieldNames, FieldValues
	end if
	
	If (LIFE <> 0) then
		FieldValues(1) = 1
		FieldValues(2) = 1
		FieldValues(3) = LIFE
		FieldValues(4) = 2
		AddRecord.AddNew FieldNames, FieldValues
	end if

	REM Saving disability insurance choices
	if (DIS <> 0) then
		FieldValues(1) = 5
		FieldValues(2) = 1
		FieldValues(3) = DIS
		FieldValues(4) = Request.Form("DisTaxation")
		AddRecord.AddNew FieldNames, FieldValues
	end if

	REM Saving retirement choices. 
	if ( Request.Form("RetDedText") <> 0) then
		FieldValues(1) = 6
		FieldValues(2) = 1
		FieldValues(3) = Request.Form("RetDedText")
		FieldValues(4) = 2
		AddRecord.AddNew FieldNames, FieldValues
	end if%>
	

	<%
	REM Once again, we run through the existing perks, whatever they are - 
	REM in a real app, I might have hardcoded this, but doing something with
	REM an unknown number of records and choices seemed interesting.
	Set benefits = Server.CreateObject("ADODB.Recordset")
	benefits.Open "SELECT * FROM tbBenefit WHERE BnfTypeId = 5", Session("DBConn"), 0, 1, 1
	Do While Not benefits.EOF 
		perkname = "Perk" &  benefits("PKId") 
		if ( Request.Form(perkname) ) then
			FieldValues(1) = benefits("PKId")
			FieldValues(2) = 1
			FieldValues(3) = 0
			FieldValues(4) = 3
			AddRecord.AddNew FieldNames, FieldValues
 		end if
		benefits.MoveNext
	  Loop
	%>
	<!-- benefits saved successfully--></table>
	<% 
	REM Did insurance choices imply dependents/beneficiaries? Are there any currently on record?

	Set deps = Server.CreateObject ("ADODB.Recordset")
	deps.Open "SELECT * FROM tbDependant WHERE EmployeeId = " & Session("empid"), Session("DBConn"), 0, 1, 1


	If ( deps.RecordCount > 0  OR Request.Form("HealthCov") > 1 OR Request.Form("DentalCov") > 1 OR LIFE > 0   ) then
	REM read the table tbDepType only once, to feed arrays; pass the arrays to a function that 
	REM writes them out as dropdowns and can set exisitng selections.
	Set deptypes = Server.CreateObject("ADODB.Recordset")
	deptypes.Open "SELECT PKId, Descr FROM tbDepType", Session("DBConn"), 0, 1, 1
	Dim Index(40)
	Dim Text(40)
	ind = 0
	do while not deptypes.EOF and ind < 40
		Index(ind) = deptypes("PKId")
		Text(ind) = deptypes("Descr")
		ind = ind + 1
		deptypes.MoveNext
	loop%>

		<form name=depForm method=post action="last.asp">
		<table  width=600 cellpadding=12>
		<h2> Your family and beneficiaries:</h2><span class=gif><img src="images\rule.gif" height=2 width=600></span><p>
		</table>

<table width=600 cellpadding=1 border=12 frame=box bordercolor="FFFFFF" rules=none>
<tr>
			<td valign="bottom">FirstName</td><td valign="bottom">LastName</td>
			<td></td><td></td><td></td><td valign="bottom">Birthdate</td>
		</tr>
		<tr><td colspan=6><img src="images\rule.gif" height=2 width=600></td></tr>

		<% 
		i = 1
		Do While Not deps.EOF %>
			<tr>
			<td><input type=text name="<%= i%>FName" value=<%= deps("FirstName")%> size=10></td>
			<td><input type=text name="<%= i%>LName" value=<%= deps("LastName")%> size=10></td>
			<td><select name="<%= i%>Sex">
				<option value=1>Female
				<option value=0 
				<% if deps("Gender") = 0 then%>
						SELECTED
				<% end if%>>Male
			</select></td>
			<td><select name="<%= i%>Status">
				<option value=1> Beneficiary
				<option value=0
				<% if deps("Status") = 0 then%>
						SELECTED
				<% end if%>>Spouse/partner/dependent
			</select></td>
			<td><select name="<%= i%>Type">
			<% Call OptionsFromArray(Index, Text, ind, deps("DepTypeId"))%>
			</select></td>
			<td><input type=text name="<%= i%>DateofBirth" value="<%= deps("DateOfBirth")%>" size=10><br></td>
			</tr>
		<%	i = i + 1
			deps.MoveNext
		Loop
		j = i + 6
		Do While i < j%>
			<tr>
			<td><input type=text name="<%= i%>FName" size=10></td>
			<td><input type=text name="<%= i%>LName" size=10></td>
			<td><select name="<%= i%>Sex">
				<option value=0>Female
				<option value=1>Male
			</select></td>
			<td><select name="<%= i%>Status">
				<option value=0>Spouse/partner/dependent
				<option value=1>Beneficiary
			</select></td>
			<td><select name="<%= i%>Type">
			<% Call OptionsFromArray(Index, Text, ind, -1)%>
			</select></td>
			<td><input type=text name="<%= i%>DateofBirth" size=10><br></td>
			</tr>
		<% i = i + 1
		loop %>
		</table>
		<table width=600 cellpadding=12>
			<tr><td>

			<input type=hidden name=howmany value=<%= i %>>
			<input type=hidden name=empid value=<%= Session("empid") %>>
			<input type=submit>
			</td></tr>
		</table>

		</form>
<script language=vbscript runat=server>

Sub OptionsFromArray(arrInd, arrText, nArraySize, nInt)
	dummy = 0
	do while dummy < nArraySize
		response.write "<option value=" & arrInd(dummy)
		if ( arrInd(dummy) = nInt) then
			response.write " SELECTED "
		end if
		response.write ">" & arrText(dummy)
		dummy = dummy + 1
	loop
end sub

</script>


<% Else %>
	<table  width=600 cellpadding=12>
	<tr><td><h2>Done! Good day!</h2></td></tr></table>
<% Session.Abandon %>
<% End If %> <!-- ending the no-dependents-implied section -->


<%end if%> 



</body></html>


