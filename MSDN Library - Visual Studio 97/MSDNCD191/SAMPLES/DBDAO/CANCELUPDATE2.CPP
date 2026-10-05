//  Sample CancelUpdateX2, w/ method Edit
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

void main(void)						// CancelUpdateX2
{
	CdbDBEngine		dben;
	CdbDatabase		db;
	CdbRecordset	rs;
	COleVariant		vFirst, vLast, vFirstSav, vLastSav;
	char			cCmd;			
	//
	try
	{
		db = dben.OpenDatabase(_T("Northwind.mdb"));
		rs = db.OpenRecordset(_T("Employees"), dbOpenDynaset);
		
		vFirstSav = rs.GetField(_T("FirstName"));
		vLastSav  = rs.GetField(_T("LastName"));
		
		rs.Edit();
		vFirst.SetString(_T("Cora"), VT_BSTRT);
		vLast.SetString(_T("Edmonds"), VT_BSTRT);
		rs.SetField(_T("FirstName"), &vFirst);
		rs.SetField(_T("LastName"), &vLast);
											// note 1
		printf(_T("Replace current name with '%s %s'? (Y/N): "),
					rs.GetField(_T("FirstName")).bstrVal,
					rs.GetField(_T("LastName")).bstrVal);
		scanf(_T("%c"), cCmd);
		if (toupper(cCmd) == _T('Y'))
		{
			rs.Update();
			printf(_T("\nRecord modified."));
		//	Delete new record because this is a demonstration.
			rs.SetBookmark( rs.GetLastModified() );
			rs.Edit();
			rs.SetField(_T("FirstName"), &vFirstSav);
			rs.SetField(_T("LastName"), &vLastSav);
			rs.Update();
		}
		else
		{
			rs.CancelUpdate();
			printf(_T("\nRecord not modified."));
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