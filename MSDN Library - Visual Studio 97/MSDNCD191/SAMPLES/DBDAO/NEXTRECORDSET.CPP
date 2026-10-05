//  Sample NextRecordsetX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// NextRecordsetX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsODBC;
	CdbConnection	conPubs;
	CdbRecordset	rs;
	COleVariant		vFld;
	int				nCount;
	BOOLEAN			bNext;
	LPTSTR			szSQL = 
					_T("SELECT * FROM authors; ")
					_T("SELECT * FROM stores; ")
					_T("SELECT * FROM jobs");
	//
	try
	{
	/*	Create an ODBCDirect workspace and open the Connection
		object. The DefaultCursorDriver setting is required when
		using compound SQL statements.
	*/
		wsODBC = dben.CreateWorkspace(_T(""), _T("admin"), _T(""), 
					dbUseODBC);
		wsODBC.SetDefaultCursorDriver( dbUseODBCCursor );
		conPubs = 
			wsODBC.OpenConnection(_T("Publishers"),-1L,FALSE,
			_T("ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers"));
	
	//	Construct compound SELECT statement.
		rs = conPubs.OpenRecordset(szSQL);

	//	Try printing results from each of the the three SELECT statments.
		for( nCount = 1, bNext = TRUE; bNext; nCount++ )
		{
			printf(_T("Contents of recordset #%d\n"), nCount);
			while(!rs.GetEOF())
			{
				vFld = rs.GetField(0L);
				vFld.ChangeType(VT_BSTR);
				printf(_T("\t'%s'"),   (LPTSTR)vFld.bstrVal);

				vFld = rs.GetField(1L);
				vFld.ChangeType(VT_BSTR);
				printf(_T("\t'%s'\n"), (LPTSTR)vFld.bstrVal);

				rs.MoveNext();
			}
			bNext = rs.NextRecordset();
			printf(_T("\trs.NextRecordset() = %s\n"), 
							(bNext ? _T("TRUE"):_T("FALSE")));
		}

		rs.Close();
		conPubs.Close();
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