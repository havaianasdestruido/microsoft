//  Sample RefreshLinkX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void RefreshLinkedOutput( CdbDatabase& dbsTemp )
{
	CdbRecordset	rstRmt;
	int				nCount;
	COleVariant		vFld0, vFld1;

	//	Open linked table.
		rstRmt = dbsTemp.OpenRecordset(_T("AuthorsTable"));

	//	Enumerate recordset object, but stop at 50 records.
		for (nCount = 0; (!rstRmt.GetEOF()) && (nCount < 50); nCount++)
		{
			vFld0 = rstRmt.GetField(0L);
			vFld0.ChangeType(VT_BSTR);
			vFld1 = rstRmt.GetField(1L);
			vFld1.ChangeType(VT_BSTR);
			printf(_T("\t%s\t%s\n"), 
				(LPTSTR)vFld0.bstrVal, (LPTSTR)vFld1.bstrVal);
			rstRmt.MoveNext();
		}
	
	//	Finish up.
		if (!rstRmt.GetEOF())
			printf(_T("\t[more records]\n"));
		rstRmt.Close();
}

void main(void)							// RefreshLinkX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsCurr;
	CdbTableDef		tdfLinked;
	//
	try
	{
	//	Open a database to which a linked table can be appended.
		dbsCurr = dben.OpenDatabase(_T("DB1.mdb"));

	//	Create a linked table that points to an SQL Server database.
		tdfLinked = dbsCurr.CreateTableDef(_T("AuthorsTable"));
		tdfLinked.SetConnect(_T("ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers"));
		tdfLinked.SetSourceTableName(_T("authors"));
		dbsCurr.TableDefs.Append( tdfLinked );
	
	//	Display the contents of the linked table.
		printf(_T("Data from linked table connected to first source: \n"));
		RefreshLinkedOutput( dbsCurr );

	//	Change connection information for linked table and refresh the connection
	//	in order to make the new data available.
		tdfLinked.SetConnect(_T("ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=NewPublishers"));
		tdfLinked.RefreshLink();
	
	//	Display the contents of the linked table.
		printf(_T("Data from linked table connected to second source: \n"));
		RefreshLinkedOutput( dbsCurr );

	//	Delete linked table because this is a demonstration.
		dbsCurr.TableDefs.Delete( (LPCTSTR)tdfLinked.GetName());
		
		dbsCurr.Close();
		
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