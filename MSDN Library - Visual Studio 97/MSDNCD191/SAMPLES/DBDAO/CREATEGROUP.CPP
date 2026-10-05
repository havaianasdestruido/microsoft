//  Sample CreateGroupX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<stdlib.h>

void main(void)							// CreateGroupX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsDef;
	CdbGroup		grpNew, grpTemp;
	CdbProperty		prp;
	CdbUser			usr;
	COleVariant		vVal;
	long			lCt, lMaxCt;
	//
	try
	{
		wsDef = dben.Workspaces[0L];
	
	//	Create and append a new group.
		grpNew = wsDef.CreateGroup(_T("NewGroup"), _T("AAA123456789"));
		wsDef.Groups.Append( grpNew );
	
	//	Make the user "admin" a member of the group NewGroup by creating and adding
	//	the appropriate Group object to the Users' Groups collection.
		grpTemp = wsDef.Users[_T("admin")].CreateGroup(_T("NewGroup"));
		wsDef.Users[_T("admin")].Groups.Append( grpTemp );
		
		printf(_T("Properties of %s\n"), (LPCTSTR)grpNew.GetName());
	//	Enumerate the properties collection of NewGroup. 
	//	The PID property is not readable.
		lMaxCt = grpNew.Properties.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			try
			{
			prp = grpNew.Properties[lCt];
			vVal = prp.GetValue();
			if (vVal.vt != VT_EMPTY)
				{
				vVal.ChangeType(VT_BSTR);
				printf(_T("\t%s = '%S'\n"),
					(LPCTSTR)prp.GetName(),
					(LPTSTR)vVal.bstrVal); 
				}
			}
			catch (CdbException) {};
		}
		printf(_T("Users collection of %s\n"), 
			(LPCTSTR)grpNew.GetName());

	//	Enumerate the Users collection of NewGroup.
		lMaxCt = grpNew.Users.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			usr = grpNew.Users[lCt];
			printf(_T("\t%s\n"), (LPCTSTR)usr.GetName());
		}

	//	Delete the new Group object because this is a demonstration.
		wsDef.Groups.Delete(_T("NewGroup"));

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