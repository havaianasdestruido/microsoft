//  Sample UpdateX2
/*	Notes:
	1)	GetField() returns a COleVariant, which could be one of
		several datatypes.  However, we expect a string to be 
		returned, so we treat it that way in the printf(). 
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<conio.h>
#include	<stdio.h>

void main(void)									// UpdateX2
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstEmp;
	COleVariant		vFirst(_T("Bill"), VT_BSTRT), 
					vLast(_T("Sornsin"), VT_BSTRT);
	char			szMessage[] =
						_T("AddNew in progress:\n")
						_T("\tData in buffer = %s %s\n\n")
						_T("Use Update to save buffer ")
						_T("to recordset?\n"),
					chYN;			
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		rstEmp = dbsNwind.OpenRecordset(_T("Employees"));
		
		rstEmp.AddNew();
		rstEmp.SetField(_T("FirstName"), &vFirst);
		rstEmp.SetField(_T("LastName"),  &vLast);

	//	Show contents of buffer and get user input.
		printf(szMessage,
				vFirst.bstrVal, vLast.bstrVal);
		scanf(_T("%c"), &chYN); 
		if (toupper(chYN) == _T('Y'))
		{
			rstEmp.Update();
		//	Go to the new record and show the resulting data.
			rstEmp.SetBookmark( rstEmp.GetLastModified() );

			vFirst = rstEmp.GetField(_T("FirstName"));
			vLast  = rstEmp.GetField(_T("LastName"));
			printf(_T("Data in recordset = %s %s\n"),	//note 1
				vFirst.bstrVal, vLast.bstrVal);

		//	Delete new data because this is a demonstration.
			rstEmp.Delete();
		}
		else
		{
			rstEmp.CancelUpdate();
			printf(_T("No new record added.\n"));
		}
		
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
