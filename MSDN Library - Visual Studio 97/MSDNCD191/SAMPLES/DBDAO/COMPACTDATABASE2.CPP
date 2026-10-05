//  Sample CompactDatabaseX2
/*	Notes: 
	1)	You can correlate the value of CollatingOrder with the dbSortX  
		constants in DBDAOINT.H.
	2)	We differ from the VB sample here and code for final result.
	3)	The 4th parameter is a long which can be the composite of several
		values, including encryption and database version.
	4)	(LPCTSTR) is really an operator that takes a CString and returns 
		a pointer to a 'C' string.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)								// CompactDatabaseX2
{
	CdbDBEngine		dben;
	CdbDatabase		dbs;
	CdbProperty		prp;
	COleVariant		vVal;
	long			lCt, lMaxCt;
	//
	try
	{	
		dbs = dben.OpenDatabase(_T("Nwind11.mdb"));

		// Show the properties of the original database.
		printf(	_T("%s, version %s, CollatingOrder = %ld\n"),
				(LPCTSTR)dbs.GetName(), 
				(LPCTSTR)dbs.GetVersion(),
				dbs.GetCollatingOrder());					// note 1
		dbs.Close();
		
		//	Make sure there isn't already a file with the name of the 
		//	compacted database.
		remove(_T("Nwind20.mdb"));							// note 2

		//	This statement creates a compact and encrypted Micrososft Jet 
		//	2.0 version of a Micrososft Jet version 1.1 database.

		dben.CompactDatabase(_T("Nwind11.mdb"),
							 _T("Nwind20.mdb"),NULL,
							 (dbEncrypt | dbVersion20));	// note 3
		dbs = dben.OpenDatabase(_T("Nwind20.mdb"));

		//	Show the properties of the compacted database.
		printf(	_T("%s, version %s\n"),
				(LPCTSTR)dbs.GetName(),						// note 4
				(LPCTSTR)dbs.GetVersion());
		lMaxCt = dbs.Properties.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			try 
			{
				prp = dbs.Properties[lCt];
				vVal = prp.GetValue();
				if (vVal.vt != VT_EMPTY)
				{
					vVal.ChangeType(VT_BSTR);
					printf(_T("\t%s = '%S'\n"), 
						(LPCTSTR)prp.GetName(),
						(LPCTSTR)vVal.bstrVal);	
				}
			}
			catch (CdbException) {}
		}
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