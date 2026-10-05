//  Sample OpenDatabaseX
/*	Notes:
	1)	(LPCTSTR) is actually a unary operator that takes a 
		CString and returns a pointer to C string.
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// OpenDatabaseX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsJet;
	CdbDatabase		dbsNwind, dbsPubs1, dbsPubs2, dbs;
	CdbProperty		prp;
	COleVariant		vVal;
	CString			strColl(_T("Connection"));
	long			lCt1, lCt2, lMaxCt1, lMaxCt2;
	//
	try
	{
	//	Create MS Jet workspace object.

		wsJet = dben.CreateWorkspace(_T(""), _T("admin"), _T(""), dbUseJet);

	//	Open Database object from saved MS Jet database for exclusive use. 	
		printf(_T("Opening Northwind...\n"));
		dbsNwind = wsJet.OpenDatabase(_T("Northwind.mdb"), TRUE);

	//	Open read-only Database object based on information in the connect string. 
		printf(_T("Opening pubs...\n"));
		dbsPubs1 = wsJet.OpenDatabase(_T("Publishers"), 
					dbDriverNoPrompt,
					TRUE, 
					_T("ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers"));

	/*	Open read-only Database object by entering only the missing information
		in the ODBC Driver Manager dialog box. 
	*/
		printf(_T("Opening second copy of pubs...\n"));
		dbsPubs2 = wsJet.OpenDatabase(_T("Publishers"), 
					dbDriverCompleteRequired,
					TRUE, 
					_T("ODBC;DATABASE=pubs;DSN=Publishers"));

	//	Enumerate the Database collection.
		lMaxCt1 = wsJet.Databases.GetCount();
		for ( lCt1 = 0; lCt1 < lMaxCt1; lCt1++ )
		{
			dbs = wsJet.Databases[lCt1];
			printf(_T("Database properties for %s:\n"),
				(LPCTSTR)dbs.GetName());	// note 1
	//	Enumerate the Properties collection of each Database.
			try
			{
			lMaxCt2 = dbs.Properties.GetCount();
			for ( lCt2 = 0; lCt2 < lMaxCt2; lCt2++ )
				{
				prp = dbs.Properties[lCt2];
				vVal = prp.GetValue();
				if (prp.GetName() == strColl)
					{
				// Property actually returns a Connection object
					printf(_T("\tConnection[.Name] = %s\n"),
						(LPCTSTR)dbs.GetConnection().GetName());
					}
				else
					{
					vVal.ChangeType(VT_BSTR);
					printf(_T("\t%s = '%S'\n"),
						(LPCTSTR)prp.GetName(),
						(LPSTR)vVal.bstrVal);
					}
				}
			}
			catch(CdbException) {}
		}

		dbsNwind.Close();
		dbsPubs1.Close();
		dbsPubs2.Close();
		wsJet.Close();
		
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