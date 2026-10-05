//  Sample NewPasswordX
/*	Notes:
	1)	The maximum input length allowed is 77.
	2)	This is more in the style of C than the VB example.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<stdlib.h>
#include	<conio.h>

void main(void)						// NewPasswordX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsDef;
	CdbGroup		grpNew, grpMember;
	CdbUser			usrNew;
	COleVariant		vPwd;
	CString			strPwd;
	int				nLen = 0;
	char			szInput[80] = {77},	// note 1
					szData[80]  = {0};
	//
	try
	{
	//	Get default workspace.
		wsDef = dben.Workspaces[0L];
	
	//	Create and append new user.
		usrNew = wsDef.CreateUser(_T("Pat Smith"), _T("abc123DEF456"), 
									_T("Password1"));
		wsDef.Users.Append( usrNew );

	//	Create and append new group.
		grpNew = wsDef.CreateGroup(_T("Accounting"), _T("UVW987xyz654"));
		wsDef.Groups.Append( grpNew );

	//	Make the new user a member of the new group.
		grpMember = usrNew.CreateGroup(_T("Accounting"));
		usrNew.Groups.Append( grpMember );
		
	//	Ask user for new password. If input is too long, ask again.
		do
		{
			printf(_T("Enter new password (or CR to exit)\n"));
			if (1 == sscanf( _cgets( szInput ), _T("%77s"), szData))
				nLen = strlen( szData );
			if (nLen > 14)			// note 2
				printf(_T("Password too long!\n"));
			else if (nLen > 0)		// nLen is 1 to 14
			{
				usrNew.NewPassword(_T("Password1"),	szData);
				printf(_T("Password changed!\n"));
				break;
			}
		} while (nLen);

	//	Delete the new Group object because this is a demonstration.
		wsDef.Users.Delete(_T("Pat Smith"));
		wsDef.Groups.Delete(_T("Accounting"));

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