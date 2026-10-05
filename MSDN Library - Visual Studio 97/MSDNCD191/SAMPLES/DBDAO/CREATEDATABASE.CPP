//  Sample CreateDatabaseX
/*	Note:
	1)	We differ from the VB sample here and code for final result.
	2)	We differ from the VB sample here. GetCount() is a little 
		more efficient.
	3)	We'll ignore any write-only properties.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)										// CreateDatabaseX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsDef;
	CdbDatabase		dbsNew;
	CdbProperty		prp;
	long			lMaxCt, lCt;
	COleVariant		vVal;
	//
	try
	{
	//	Get default workspace.
		wsDef = dben.Workspaces[0L];

	//	Make sure there isn't already a file with the name of the 
	//	new database.
		remove(_T("NewDB.mdb"));						// note 1

	//	Create a new encrypted database with the specified collating order.
		dbsNew = wsDef.CreateDatabase(_T("NewDB.mdb"), 
							dbLangGeneral, dbEncrypt);

	//	Enumerate the Propeties collection of the new database object.
		printf(_T("Properties of %s\n"), (LPCTSTR)dbsNew.GetName());
		lMaxCt = dbsNew.Properties.GetCount();			// note 2
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			try											// note 3
			{
				prp = dbsNew.Properties[ lCt ];
				vVal = prp.GetValue();
				if (vVal.vt != VT_EMPTY)
				{					
					vVal.ChangeType(VT_BSTR);
					printf(_T("\t%s = '%S'\n"), 
						(LPCTSTR)prp.GetName(), vVal.bstrVal);
				}
			}
			catch (CdbException) {}
		}
		dbsNew.Close();

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