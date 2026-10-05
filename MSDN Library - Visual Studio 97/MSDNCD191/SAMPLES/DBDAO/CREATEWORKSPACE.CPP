//  Sample CreateWorkspaceX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// CreateWorkspaceX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsODBC, wsJet;
	CdbProperty		prp;
	COleVariant		vVal;
	long			lCt, lMaxCt;
	//
	try
	{
	//	Create an ODBCDirect workspace. Until you create the 
	//	MS Jet workspace, the MS Jet Database Engine will not 
	//	be loaded into memory.

		wsODBC = dben.CreateWorkspace(_T("ODBCWorkspace"),
			_T("admin"), _T(""), dbUseODBC);
		dben.Workspaces.Append( wsODBC);

	//	Create an unnamed Workspace object of the type specified
	//	by the DefaultType property of DBEngine (dbUseJet).

		dben.SetDefaultType( dbUseJet );
		wsJet = dben.CreateWorkspace(_T(""), _T("admin"), _T(""));
	
	//	Enumerate the Workspaces collection.
		printf(_T("Workspace objects in Workspaces collection:\n"));
		lMaxCt = dben.Workspaces.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
			printf(_T("\t%s\n"), (LPCTSTR)dben.Workspaces[lCt].GetName());

	//	Enumerate Properties collection of ODBCDirect workspace.
		printf(_T("Properties of %s\n"), (LPCTSTR)wsODBC.GetName());
		lMaxCt = wsODBC.Properties.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			try
			{
			prp = wsODBC.Properties[lCt];
			vVal = prp.GetValue();
			vVal.ChangeType(VT_BSTR);
			printf(_T("\t%s = '%S'\n"),
				(LPCTSTR)prp.GetName(), (LPTSTR)vVal.bstrVal);
			}
			catch (CdbException) {}
		}

	//	Enumerate Properties collection of Microsoft Jet workspace.
		printf(_T("Properties of unnamed Microsoft Jet workspace\n"));
		lMaxCt = wsJet.Properties.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			try
			{
			prp = wsJet.Properties[lCt];
			vVal = prp.GetValue();
			vVal.ChangeType(VT_BSTR);
			printf(_T("\t%s = '%S'\n"),
				(LPCTSTR)prp.GetName(), (LPTSTR)vVal.bstrVal);
			}
			catch (CdbException) {}
		}
			
		wsODBC.Close();
		wsJet.Close();

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