//  Sample DeleteX
/* Notes: 
	1)	This notation is a shortcut. We could be more rigorous
		and assign a variant, then check that its datatype is LONG.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void DeleteRecord( CdbRecordset rstTemp, long lSeek )
{
	COleVariant	vKey(lSeek, VT_I4);
	//
	rstTemp.Seek(_T("="), 1L, vKey);
	if ( rstTemp.GetNoMatch() )
		printf(_T("No employee #%ld in file!\n"), lSeek);
	else
	{
		rstTemp.Delete();
		printf(_T("Record for employee #%ld deleted!\n"), lSeek);
	}
}

void main(void)											// DeleteX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwnd;
	CdbRecordset	rstEmp;
	long			lEmpID;
	COleVariant		vFirst(_T("Janelle"), VT_BSTRT),
					vLast(_T("Tebbs"), VT_BSTRT);
	//
	try
	{	dbsNwnd = dben.OpenDatabase(_T("northwind.mdb"));
		rstEmp = dbsNwnd.OpenRecordset(_T("Employees"));

	//	Add temporary record to be deleted.
		rstEmp.SetIndex(_T("PrimaryKey"));
		rstEmp.AddNew();
		rstEmp.SetField(_T("FirstName"), &vFirst);
		rstEmp.SetField(_T("LastName"), &vLast);
		rstEmp.Update();
		rstEmp.SetBookmark( rstEmp.GetLastModified());
		lEmpID = rstEmp.GetField(_T("EmployeeID")).lVal;	// note 1

	//	Delete the employee record with the specified ID number.
		DeleteRecord( rstEmp, lEmpID );

		rstEmp.Close();
		dbsNwnd.Close();

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