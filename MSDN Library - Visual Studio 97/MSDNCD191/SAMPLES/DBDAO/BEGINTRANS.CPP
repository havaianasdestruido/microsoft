//  Sample BeginTransX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<conio.h>

void main(void)							// BeginTransX
{
	CdbDBEngine		dben;
	CdbWorkspace	wsDef;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstEmp;
	COleVariant		vFirst, vLast, vTitle,
					vAcctExec(_T("Account Executive"), VT_BSTRT),
					vSalesRep(_T("Sales Representative"), VT_BSTRT);
	char			chYN;
	LPTSTR			szMessage = 
						_T("Employee: %s, %s\n")
						_T("Change title to Account Executive? ");
	//
	try
	{
	//	Get default workspace.
		wsDef = dben.Workspaces[0L];
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		rstEmp = dbsNwind.OpenRecordset(_T("Employees"));

	//	Start of outer transaction.
		wsDef.BeginTrans();
	//	Start of main transaction.
		wsDef.BeginTrans();

	//	Loop through recordset and ask if the user wants to
	//	change the title for a specified employee.
		while(!rstEmp.GetEOF())
		{
			vTitle = rstEmp.GetField(_T("Title"));
			if (vTitle == vSalesRep)
			{
			//	Change the title for the specified employee.
				vFirst = rstEmp.GetField(_T("First"));
				vLast  = rstEmp.GetField(_T("Last"));

				vFirst.ChangeType(VT_BSTR);
				vLast.ChangeType(VT_BSTR);

				printf(szMessage, 
							vLast.bstrVal, 
							vFirst.bstrVal);
				scanf(_T("%c"), &chYN);	printf(_T("\n"));
				if (toupper(chYN) == _T('Y'))
				{
					rstEmp.Edit();
					rstEmp.SetField(_T("Title"), &vAcctExec);
					rstEmp.Update();
				}
				rstEmp.MoveNext();
			}
		}
	//	Ask if the user wants to commit all the changes 
	//	made above.
		printf(_T("Save all changes? "));
		scanf(_T("%c"), &chYN); printf(_T("\n"));
		if (toupper(chYN) == _T('Y'))
			wsDef.CommitTrans();
		else
			wsDef.Rollback();

	//	Print current data in recordset.
		rstEmp.MoveFirst();
		while(!rstEmp.GetEOF())
		{
			vFirst = rstEmp.GetField(_T("First"));
			vLast  = rstEmp.GetField(_T("Last"));
			vTitle = rstEmp.GetField(_T("Title"));

			vFirst.ChangeType(VT_BSTR);
			vLast.ChangeType(VT_BSTR);
			vTitle.ChangeType(VT_BSTR);

			printf(_T("%s, %s - %s\n"), 
						vLast.bstrVal, 
						vFirst.bstrVal,
						vTitle.bstrVal);
			rstEmp.MoveNext();
		}

	//	Roll back any changes made by the user since this is
	//	a demonstration
		wsDef.Rollback();
		rstEmp.Close();
		dbsNwind.Close();
			
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