//  Sample EditX
/*	Notes:
	1)	_cgets() is allowed to read a maximum of 77 characters.
	2)	_cgets() returns console input up to the maximum specified in szInput.
		If sscanf() places 1 non-blank string in szData, that string is placed
		in an empty COleVariant.
	3)	VT_BSTRT specifies an ANSI variant string instead of the default, UniCode.
		However, the datatype of the variant will be the generic VT_BSTR.
		Another way to code this is: vFirst.SetString(szData, VT_BSTRT);
	4)	We're taking a shortcut here.  GetField() returns a COleVariant, which 
		could be one of several datatypes.  However, we expect a string to be 
		returned, so we treat it that way in the printf(). 
	5)	We diverge from the VB example here and use the existing routine.
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<conio.h>
#include	<stdio.h>

void EditName(	CdbRecordset	rstTemp,
				LPVARIANT		lpvFirst,
				LPVARIANT		lpvLast)
{
//	Make changes to record and set the bookmark to keep the same 
//	record current.
	rstTemp.Edit();
	rstTemp.SetField(_T("FirstName"), lpvFirst);
	rstTemp.SetField(_T("LastName"),  lpvLast);
	rstTemp.Update();
	rstTemp.SetBookmark( rstTemp.GetLastModified());
}

void main(void)									// EditX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstEmp;
	COleVariant		vFirst, vLast,
					vFirstOld, vLastOld;
	char			szInput[80] = {77},			// note 1
					szData[80],
					szMessage[] = 
					_T("Old data: %s %s\n")
					_T("New data: %s %s\n");			
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		rstEmp = dbsNwind.OpenRecordset(_T("Employees"), dbOpenDynaset);

		//	Store original data.
		vFirstOld = rstEmp.GetField(_T("FirstName"));
		vLastOld  = rstEmp.GetField(_T("LastName"));

		//	Get new data for the record.
		printf(_T("\nEnter first name: "));
		if (1 == sscanf( _cgets( szInput ), _T("%77s"), szData))	// note 2
			vFirst = COleVariant( szData, VT_BSTRT);			// note 3
		
		printf(_T("\nEnter last name: "));
		if (1 == sscanf( _cgets( szInput ), _T("%77s"), szData))
			vLast = COleVariant( szData, VT_BSTRT);
		
		printf(_T("\n"));

		if ((vFirst.vt == VT_BSTR) && (vLast.vt == VT_BSTR))
		{
			// Update record with new data.
			EditName( rstEmp, &vFirst, &vLast);
			
			// Show old and new data.
			vFirst = rstEmp.GetField(_T("FirstName"));
			vLast  = rstEmp.GetField(_T("LastName"));
			printf(szMessage,							// note 4
					vFirstOld.bstrVal,	vLastOld.bstrVal,
					vFirst.bstrVal,		vLast.bstrVal);

			// Restore original data because this is a demonstration.
			EditName( rstEmp, &vFirstOld, &vLastOld);	// note 5
		}
		else
			printf(_T("You must input a string for the first and last name.\n"));
		//
		rstEmp.Close();
		dbsNwind.Close();

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
