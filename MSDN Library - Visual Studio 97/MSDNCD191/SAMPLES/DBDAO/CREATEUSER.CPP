//  Sample CreateUserX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<stdlib.h>

void main(void)							// CreateUserX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsDef;
	CdbGroup		grpNew, grp;
	CdbUser			usrNew, usr;
	COleVariant		vVal;
	CdbProperty		prp;
	long			lCt, lMaxCt;
	//
	try
	{
		wsDef = dben.Workspaces[0L];
	
	//	Create and append a new User.
		usrNew = wsDef.CreateUser(_T("NewUser"));
		usrNew.SetPID(_T("AAA12345679"));
		usrNew.SetPassword(_T("NewPassword"));
		wsDef.Users.Append( usrNew );
	
	//	Create and append new Group.
		grpNew = wsDef.CreateGroup(_T("NewGroup"), _T("AAA123456789"));
		wsDef.Groups.Append( grpNew );

	//	Make the user "NewUser" a member of the group "NewGroup" by creating
	//	and adding the appropriate user object to the Groups' Users collection.
		usr = wsDef.Groups[_T("NewGroup")].CreateUser(_T("NewUser"));
		wsDef.Groups[_T("NewGroup")].Users.Append( usr );
		
		printf(_T("Properties of %s\n"), (LPCTSTR)usrNew.GetName());
	//	Enumerate the properties collection of NewUser. 
	//	The PID property is not readable.
		lMaxCt = usrNew.Properties.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			try
			{
			prp = usrNew.Properties[lCt];
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

		printf(_T("Groups collection of %s\n"), (LPCTSTR) usrNew.GetName());
	//	Enumerate the Groups collection of NewUser.
		lMaxCt = usrNew.Groups.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			grp = usrNew.Groups[ lCt ];
			printf(_T("\t%s\n"), (LPCTSTR)grp.GetName());
		}

	//	Delete the new User and Group object because this is a demonstration.
		wsDef.Users.Delete(_T("NewUser"));
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