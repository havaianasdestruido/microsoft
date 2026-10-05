//  Sample RegisterDatabaseX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// RegisterDatabaseX
{
	CdbDBEngine	dben;
	CdbDatabase	dbsReg;
	CString		strDescription, strAttributes;
	char		szInput[80],
				szRegMsg[] = _T("Use regedit.exe to view changes: ")
					_T("HKEY_CURRENT_USER\\Software\\ODBC\\ODBC.INI"),
				szPrompt[] = _T("Enter a description for the database ")
					_T("to be registered: ");
	//
	try
	{
	//	Build keywords string.
		printf(szPrompt);
		scanf(_T("%79s"), szInput);	
		strDescription = szInput;

		strAttributes = 
			CString(_T("Database=pubs")) + _T('\r') +
			_T("Description=") + strDescription + _T('\r') + 
			_T("OemToAnsi=No") + _T('\r') + 
			_T("Server=Server1");
	
	//	Update Windows Registry.
		dben.RegisterDatabase(_T("Publishers"), _T("SQL Server"), 
			TRUE, (LPCTSTR)strAttributes);
		printf(szRegMsg);
		
		printf(_T("Normal End\n"));	
	}
	//	Notify user of any errors that result from the invalid data.
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