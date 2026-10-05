//  Sample CreatQueryDefX
//
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void GetrstTemp( CdbQueryDef qd )
{
	CdbRecordset	rs;
	//
	printf(_T("'%s'\t'%s'\n"),
		(LPCTSTR) qd.GetName(), (LPCTSTR) qd.GetSQL());

	//	Open recordset from querydef.
	rs = qd.OpenRecordset(dbOpenSnapshot);

	//	Populate recordset and print number of records.
	rs.MoveLast();
	printf(_T("\tNumber of records = %ld\n\n"),
		rs.GetRecordCount());
	rs.Close();
}

void main(void)									// CreatQueryDefX
{
	CdbDBEngine		dben;
	CdbDatabase		db;
	CdbQueryDef		qdTemp, qdNew;
	//
	try
	{
		db = dben.OpenDatabase(_T("Northwind.mdb"));
		
		// Create temporary QueryDef
		qdTemp = db.CreateQueryDef(_T(""), _T("SELECT * FROM EMPLOYEES"));

		// Open recordset and print report.
		GetrstTemp( qdTemp );

		// Create permanent QueryDef
		qdNew = db.CreateQueryDef(_T("NewQueryDef"), _T("SELECT * FROM CATEGORIES"));

		// Open recordset and print report.
		GetrstTemp( qdNew );

		// Delete new QueryDef because this is a demonstration
		db.QueryDefs.Delete( (LPCTSTR)qdNew.GetName());

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