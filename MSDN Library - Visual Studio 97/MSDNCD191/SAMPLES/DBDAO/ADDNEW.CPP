//  Sample AddNewX
/*	Notes:
	1)	_cgets() returns console input up to the maximum specified in szInput.
		If sscanf() places 1 non-blank string in szData, that string is placed
		in an empty COleVariant.
	2)	VT_BSTRT specifies an ANSI variant string instead of the default, UniCode.
		However, the datatype of the variant will be the generic VT_BSTR.
		Another way to code this is:  vFirst.SetString(szData, VT_BSTRT);
	3)	We're making an assumption here.  We expect a string to be returned
		so we treat it that way in the printf(). 
		We could also say:
			printf(_T("...%s..."), rs.GetField(_T("FirstName")).bstrVal...); 
	4)	Delete() is called to match the VB example.  If you want to see that the
		code works, comment out this statement, examine the database with a tool 
		like MS Access, then use that tool to delete the row.
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<conio.h>
#include	<stdio.h>

void AddName(	CdbRecordset	rst,
				LPVARIANT		lpvFirst,
				LPVARIANT		lpvLast)
{
//	Adds a new record to a recordset using the data passed by the 
//	calling procedure. The new record is then made the current record.
	rst.AddNew();
	rst.SetField(_T("FirstName"), lpvFirst);
	rst.SetField(_T("LastName"),  lpvLast);
	rst.Update();
	rst.SetBookmark( rst.GetLastModified());
}

void main(void)									// AddNewX
{
	CdbDBEngine		dben;
	CdbDatabase		db;
	CdbRecordset	rs;
	COleVariant		vFirst, vLast;
	char			szInput[80] = {77},	
					szData[80];			
	//
	try
	{
		db = dben.OpenDatabase(_T("employee.mdb"));
		rs = db.OpenRecordset(_T("Employees"), dbOpenDynaset);
			
		//	Get data from the user.
		printf(_T("\nEnter first name: "));
		if (1 == sscanf( _cgets( szInput ), _T("%77s"), szData))// note 1
			vFirst = COleVariant( szData, VT_BSTRT);			// note 2
		
		printf(_T("\nEnter last name: "));
		if (1 == sscanf( _cgets( szInput ), _T("%77s"), szData))
			vLast = COleVariant( szData, VT_BSTRT);
		
		printf(_T("\n"));

		if ((vFirst.vt == VT_BSTR) && (vLast.vt == VT_BSTR))
		{
			// Call the function that adds the record.
			AddName( rs, &vFirst, &vLast);
			
			// Show the newly added data.
			vFirst = rs.GetField(_T("FirstName"));
			vLast  = rs.GetField(_T("LastName"));
			printf(_T("New record: %s %s\n"),		// note 3
					(LPTSTR)vFirst.bstrVal, 
					(LPTSTR)vLast.bstrVal);

			// Delete the new record because this is a demonstration.
			rs.Delete();						// note 4
		}
		else
			printf(_T("You must input a string for the first and last name.\n"));
		//
		rs.Close();
		db.Close();

		printf(_T("Normal End\n"));	
	}
	catch (CdbException)
	{
		for (long ct = 0; ct < dben.Errors.GetCount(); ct++)
		{
			printf(_T("\t Error #%ld: #%ld -- %s\n"),
				ct,
				DBERR(dben.Errors[ct].GetNumber()),
				dben.Errors[ct].GetDescription());
		}
	}
	catch (CException *e)
	{
		printf(_T("MFC Exception\n"));
		e->Delete();
	}
	catch (...)
	{
		printf(_T("Win32 Exception\n"));
	}
}
