//  Sample OpenConnectionX
/*	Notes:
	1)	We've already printed this connection's name.
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// OpenConnectionX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsODBC;
	CdbConnection	con1, con2, con3, conLoop;
	long			lCt, lMaxCt;
	//
	try
	{
	//	Create an ODBCDirect workspace object.

		wsODBC = dben.CreateWorkspace(_T("NewODBCWorkspace"),
			_T("admin"), _T(""), dbUseODBC);

	/*	Open Connection object using supplied information in connect string.
		If this information were insufficient, you could trap for an error
		rather than go to the ODBC Driver Manager dialog box. 	
	*/
		printf(_T("Opening Connection 1...\n"));
		con1 = wsODBC.OpenConnection(_T("Connection1"), 
					dbDriverNoPrompt,
					FALSE, 
					_T("ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers"));

	/*	Open read-only Connection object based on information you may enter in
		the ODBC Driver Manager dialog box. 
	*/
		printf(_T("Opening Connection 2...\n"));
		con2 = wsODBC.OpenConnection(_T("Connection2"), 
					dbDriverPrompt,
					TRUE, 
					_T("ODBC;DSN=Publishers"));

	/*	Open read-only Connection object by entering only the missing information
		in the ODBC Driver Manager dialog box. 
	*/
		printf(_T("Opening Connection 3...\n"));
		con3 = wsODBC.OpenConnection(_T("Connection3"), 
					dbDriverCompleteRequired,
					TRUE, 
					_T("ODBC;DATABASE=pubs;DSN=Publishers"));

	//	Enumerate the Connections collection.
		lMaxCt = wsODBC.Connections.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			conLoop = wsODBC.Connections[lCt];
	//	Print property values by explicitly calling each Property object; the 
	//	Connection object does not support a Properties collection. 

			printf(_T("Connection properties for %s:\n"), (LPCTSTR) conLoop.GetName());
			printf(_T("\tConnect = %s\n"), (LPCTSTR) conLoop.GetConnect());
			printf(_T("\tDatabase[.Name] = %s\n"),
				(LPCTSTR) conLoop.GetDatabase().GetName());
			printf(_T("\tName = %s\n"), (LPCTSTR) conLoop.GetName());	// note 1
			printf(_T("\tQueryTimeout = %d\n"), conLoop.GetQueryTimeout());
			printf(_T("\tRecordsAffected = %ld\n"), conLoop.GetRecordsAffected());
			printf(_T("\tStillExecuting = %s\n"), 
				(conLoop.GetStillExecuting() ? _T("TRUE") : _T("FALSE")));
			printf(_T("\tTransactions = %s\n"),
				(conLoop.GetDatabase().GetTransactions() ? _T("TRUE") : _T("FALSE")));
			printf(_T("\tUpdatable = %s\n"), 
				(conLoop.GetUpdatable() ? _T("TRUE") : _T("FALSE")));
		}

		con1.Close();
		con2.Close();
		con3.Close();
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