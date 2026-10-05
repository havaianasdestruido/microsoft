//  Sample RepairDatabaseX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// RepairDatabaseX
{
	CdbDBEngine	dben;
	char		chYN;
	//
	try
	{	
		printf(_T("Repair the Northwind database? "));
		scanf(_T("%c"), &chYN);
		if (toupper(chYN) == _T('Y'))
			dben.RepairDatabase(_T("Northwind.mdb"));

		printf(_T("\nNormal End\n"));	
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