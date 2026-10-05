//  Sample OpenRecordsetX
//
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void OpenRecordsetOutput( CdbRecordset rstOutput )
{
	COleVariant	vFld0, vFld1;

	//	Enumerate the specified recordset object.
	while (!rstOutput.GetEOF())
	{
		vFld0 = rstOutput.GetField(0L);
		vFld0.ChangeType(VT_BSTR);
		vFld1 = rstOutput.GetField(1L);
		vFld1.ChangeType(VT_BSTR);
		printf(_T("\t%s\t%s\n"), 
			(LPTSTR)vFld0.bstrVal, (LPTSTR)vFld1.bstrVal);
		rstOutput.MoveNext();
	}
}

void main(void)							// OpenRecordsetX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsODBC, wsJet;
	CdbDatabase		dbs;
	CdbConnection	conPubs;
	CdbRecordset	rstTemp1, rstTemp2;
	LPTSTR			szTitle1 = _T("Opening forward-only-type ")
						_T("recordset where the source is ")
						_T("a QueryDef object...\n"),
					szTitle2 = _T("Opening read-only dynaset-type ")
						_T("recordset where the source is an ")
						_T("SQL statement...\n"),
					szTitle3 = _T("Opening recordset from existing ")
						_T("recordset object to filter records\n"),
					szTitle4 = _T("Opening dynamic-type recordset from ")
						_T("an ODBC connection...\n"),
					szTitle5 = _T("Opening snapshot-type recordset ")
						_T("based on asynchronous query to ")
						_T("ODBC connection...\n");
	//
	try
	{
	//	Open MS Jet and ODBCDirect workspaces, MS Jet database, and ODBCDirect
	//	connection.
 		wsJet   = dben.CreateWorkspace(_T(""), _T("admin"), _T(""), dbUseJet);
		wsODBC  = dben.CreateWorkspace(_T(""), _T("admin"), _T(""), dbUseODBC);
		
		dbs     = wsJet.OpenDatabase(_T("Northwind.mdb"));
		conPubs = wsODBC.OpenConnection(_T(""),-1L,FALSE,
					_T("ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers"));

	//	Open five different Recordset objects and display the contents of each.
		printf(szTitle1);
		rstTemp1 = dbs.OpenRecordset(_T("Ten Most Expensive Products"),
						dbOpenForwardOnly);
		OpenRecordsetOutput( rstTemp1 );

		printf(szTitle2);
		rstTemp1 = dbs.OpenRecordset(_T("SELECT * FROM Employees"),
					dbOpenDynaset);
		OpenRecordsetOutput( rstTemp1 );

	//	Use the Filter property to retrieve only certain records with the
	//	next OpenRecordset call.
		printf(szTitle3);
		rstTemp1.SetFilter(_T("LastName >= 'M'"));
		rstTemp2 = rstTemp1.OpenRecordset();
		OpenRecordsetOutput( rstTemp2 );

		printf(szTitle4);
		rstTemp1 = conPubs.OpenRecordset(_T("SELECT * FROM Stores"),
					dbOpenDynamic);
		OpenRecordsetOutput( rstTemp1 );

	//	Use the StillExecuting property to determine when the
	//	recordset is ready for manipulation.
		printf(szTitle5);
		rstTemp1 = conPubs.OpenRecordset( _T("Publishers"),
					dbOpenSnapshot, dbRunAsync);
		while(rstTemp1.GetStillExecuting())
			printf(_T("\t[still executing...]\n"));
		OpenRecordsetOutput( rstTemp1 );

		rstTemp1.Close(); 
		rstTemp2.Close();
		dbs.Close();
		conPubs.Close();
		wsJet.Close();
		wsODBC.Close();

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