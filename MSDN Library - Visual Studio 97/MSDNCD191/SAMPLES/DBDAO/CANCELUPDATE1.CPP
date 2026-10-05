//  Sample CancelUpdateX1, w/ method AddNew
/*	Notes:
	1)	This is a shortcut.  GetField() returns a COleVariant 
		which could be one of many datatypes. However, we know
		it's a string so we'll treat it that way in the printf().
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<stdlib.h>

void main(void)						// CancelUpdateX1
{
	CdbDBEngine		dben;
	CdbDatabase		db;
	CdbRecordset	rs;
	COleVariant		vFirst, vLast;
	char			cCmd;			
	//
	try
	{
		db = dben.OpenDatabase(_T("Northwind.mdb"));
		rs = db.OpenRecordset(_T("Employees"), dbOpenDynaset);
		
		rs.AddNew();
		vFirst.SetString(_T("Kimberly"), VT_BSTRT);
		vLast.SetString(_T("Bowen"), VT_BSTRT);
		rs.SetField(_T("FirstName"), &vFirst);
		rs.SetField(_T("LastName"), &vLast);
											// note 1
		printf(_T("Add new record for '%s %s'? (Y/N): "),
					rs.GetField(_T("FirstName")).bstrVal,
					rs.GetField(_T("LastName")).bstrVal);
		scanf(_T("%c"), cCmd);					// note 2
		if (toupper(cCmd) == _T('Y'))
		{
			rs.Update();
			printf(_T("\nRecord added."));
		//	Delete new record because this is a demonstration.
			rs.SetBookmark( rs.GetLastModified() );
			rs.Delete();
		}
		else
		{
			rs.CancelUpdate();
			printf(_T("\nRecord not added."));
		}
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