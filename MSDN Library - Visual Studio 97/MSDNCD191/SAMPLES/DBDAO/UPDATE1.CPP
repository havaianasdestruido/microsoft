//  Sample UpdateX1
/*	Notes:
	1)	We're taking a shortcut here.  GetField() returns a COleVariant, 
		which could be one of several datatypes.  However, we expect a 
		string to be returned, so we treat it that way in the printf(). 
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<conio.h>
#include	<stdio.h>

void main(void)									// UpdateX1
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstEmp;
	COleVariant		vFirst(_T("Linda"), VT_BSTRT), 
					vLast(_T("Kobara"), VT_BSTRT),
					vFirstOld, vLastOld;
	char			szMessage[] =
						_T("Edit in progress:\n")
						_T("\tOriginal data  = %s %s\n")
						_T("\tData in buffer = %s %s\n\n")
						_T("Use Update to replace original ")
						_T("data with the buffered data in ")
						_T("the recordset?\n"),
					chYN;			
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		rstEmp = dbsNwind.OpenRecordset(_T("Employees"));
		
	//	Store original data.
		vFirstOld = rstEmp.GetField(_T("FirstName"));
		vLastOld  = rstEmp.GetField(_T("LastName"));

		rstEmp.Edit();

	//	Change data in edit buffer.
		rstEmp.SetField(_T("FirstName"), &vFirst);
		rstEmp.SetField(_T("LastName"),  &vLast);

	//	Show contents of buffer and get user input.
		printf(szMessage);
		scanf(_T("%c"), &chYN); 
		if (toupper(chYN) == _T('Y'))
			rstEmp.Update();
		else
			rstEmp.CancelUpdate();

	//	Show the resulting data.
		vFirst = rstEmp.GetField(_T("FirstName"));
		vLast  = rstEmp.GetField(_T("LastName"));
		printf(_T("Data in recordset = %s %s\n"),	//note 1
			vFirst.bstrVal, vLast.bstrVal);

	//	Restore original data because this is a demonstration.
		if (!((vFirst == vFirstOld) && (vLast == vLastOld)))
		{
			rstEmp.Edit();
			rstEmp.SetField(_T("FirstName"), &vFirstOld);
			rstEmp.SetField(_T("LastName"),  &vLastOld);
			rstEmp.Update();
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
