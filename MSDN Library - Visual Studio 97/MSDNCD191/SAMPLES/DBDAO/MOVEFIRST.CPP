//  Sample MoveFirstX
//

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<conio.h>

LPTSTR		szAlreadyAt = _T("Already at %s of recordset!\n");

void MoveAny(int nChoice, CdbRecordset rstTemp)
{
//	Use specified method, trapping BOF and EOF.
	switch (nChoice) 
	{
	case 1:
		rstTemp.MoveFirst();
		break;
	case 2:
		rstTemp.MoveLast();
		break;
	case 3:
		rstTemp.MoveNext();
		if (rstTemp.GetEOF())
		{
		printf(szAlreadyAt, _T("end"));
		rstTemp.MoveLast();
		}
		break;
	case 4:
		rstTemp.MovePrevious();
		if (rstTemp.GetBOF())
		{
		printf(szAlreadyAt, _T("beginning"));
		rstTemp.MoveFirst();
		}
		break;
	default:
		break;
	}
}

void main(void)							// MoveFirstX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstEmp;
	COleVariant		vAbsPos, vRcdCt, vFirst, vLast;
	int				nCmd;
	LPTSTR			szSelect = 
						_T("SELECT FirstName, LastName ")
						_T("FROM Employees ORDER BY LastName"),
					szMessage = 
						_T("Name: %s %s\n")
						_T("Record (%ld) of %ld\n")
						_T("[1 - MoveFirst, 2 - MoveLast, \n")
						_T(" 3 - MoveNext,  4 - MovePrevious]\n")
						_T("(Any other digit to exit)\n");
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		rstEmp = dbsNwind.OpenRecordset(szSelect, dbOpenSnapshot);

	//	Populate Recordset.
		rstEmp.MoveLast();
		rstEmp.MoveFirst();

	//	Show current record information and get user's 
	//	method choice.
		for(;;)
		{
			vFirst = rstEmp.GetField(_T("FirstName"));
			vLast  = rstEmp.GetField(_T("LastName"));
			vAbsPos = rstEmp.GetAbsolutePosition();
			vRcdCt = rstEmp.GetRecordCount();

			vFirst.ChangeType(VT_BSTR);
			vLast.ChangeType(VT_BSTR);
			vAbsPos.ChangeType(VT_I4);
			vRcdCt.ChangeType(VT_I4);

			printf(	szMessage, 
					vFirst.bstrVal, vLast.bstrVal);
					vAbsPos.lVal + 1, vRcdCt.lVal,
			scanf(_T("%d"), &nCmd);
			if ((nCmd < 1) || (nCmd > 4)) break;

		//	Call method based on user's input.
			MoveAny( nCmd, rstEmp);
		}

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