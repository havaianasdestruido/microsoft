//  Sample AppendX
/*	Notes:
	1)	APPEND or DELETE
	2)	Optional: Ignored when Command is DELETE.
	3)	Size is implicit, based on Type.
	4)	The documentation says the terminating value is "count-1",
		because the VB "for" statement tests "<=". The C idiom is "<", 
		so the terminating value is just "count".
	5)	This cast is really an operater that takes a CString and returns  
		a pointer to a string.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void AppendDeleteField(	CdbTableDef	tdx,
						LPCTSTR 	szCmd,	// note 1
						LPCTSTR 	szName,
						long		lType,	// note 2
						long		lSize)
{
	CdbField	fld;
	// Check first to see if TableDef object is updatable.  If it isn't, 
	// control is passed back to the calling procedure.
	if (!tdx.GetUpdatable())
	{
		printf(_T("TableDef not Updatable! Unable to complete task.\n"));
		return;
	}
	else
	{
	//	Depending on the passed data, append or delete a field to 
	//	the Fields collection of the specified TableDef object.
		if (stricmp(szCmd, "APPEND") == 0)
		{
			fld = tdx.CreateField(szName, lType, lSize);
			tdx.Fields.Append(fld);
		}
		else
			tdx.Fields.Delete(szName);
	}
}

void main(void)								//AppendX
{
	CdbDBEngine	dben;
	CdbDatabase	db;
	CdbTableDef	td;
	CdbField	fld;
	long		lCt, lMaxCt;
	//
	try
	{
		db = dben.OpenDatabase(_T("employee.mdb"));
		td = db.TableDefs[_T("Employees")];
		
		AppendDeleteField( td, "APPEND", "E-mail", dbText,    50 );
		AppendDeleteField( td, "APPEND", "Http",   dbText,    80 );
		AppendDeleteField( td, "APPEND", "Quota",  dbInteger, 0 );	// note 3

		printf(_T("\nFields after Append\nType\tSize\tName\n"));
		lMaxCt = td.Fields.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)							// note 4
		{
			fld = td.Fields[ lCt ];
			printf(_T("%d\t%ld\t%s\n"),
				fld.GetType(),
				fld.GetSize(),
				(LPCTSTR) fld.GetName());							// note 5
		}
		//
		AppendDeleteField( td, "DELETE", "E-mail", 0, 0 );
		AppendDeleteField( td, "DELETE", "Http",   0, 0 );
		AppendDeleteField( td, "DELETE", "Quota",  0, 0 );

		printf(_T("\nFields after Delete\nType\tSize\tName\n"));
		lMaxCt = td.Fields.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)  
		{
			fld = td.Fields[ lCt ];
			printf(_T("%d\t%ld\t%s\n"),
				fld.GetType(),
				fld.GetSize(),
				(LPCTSTR) fld.GetName());
		}

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

