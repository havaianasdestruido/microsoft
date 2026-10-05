//  Sample CompactDatabaseX1
/*	Notes: 
	1)	You can correlate the value of CollatingOrder with the dbSortX  
		constants in DBDAOINT.H.
	2)	We differ from the VB sample here and code for final result.
	3)	The 3rd and 5th parameters take a string that can include several keyword 
		parameters.  The dbLangX constants defined in DBDAOINT.H specify the 
		collating order.
	4)	(LPCTSTR) is really a unary operator that converts a CString to string.
	5)	The returned value for collating order should be 1036, dbSortKorean.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)								// CompactDatabaseX1
{
	CdbDBEngine		dben;
	CdbDatabase		dbs;
	//
	try
	{	
		dbs = dben.OpenDatabase(_T("Northwind.mdb"));

		// Show the properties of the original database.
		printf(	_T("%s, version %s, CollatingOrder = %ld\n"),
				(LPCTSTR)dbs.GetName(), 
				(LPCTSTR)dbs.GetVersion(),
				dbs.GetCollatingOrder());					// note 1
		dbs.Close();

		//	Make sure there isn't already a file with the name of the 
		//	compacted database.
		remove(_T("NWindKorean.mdb"));						// note 2

		//	This statement creates a compact version of the Northwind database that
		//	uses a Korean language collating order. 
		dben.CompactDatabase(_T("Northwind.mdb"),
							 _T("NWindKorean.mdb"),
							 dbLangKorean);					// note 3
							
		dbs = dben.OpenDatabase(_T("NWindKorean.mdb"));

		//	Show the properties of the compacted database.
		printf(	_T("%s, version %s, CollatingOrder = %ld\n"),		
				(LPCTSTR)dbs.GetName(),						// note 4
				(LPCTSTR)dbs.GetVersion(),
				dbs.GetCollatingOrder());					// note 5

		dbs.Close();

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