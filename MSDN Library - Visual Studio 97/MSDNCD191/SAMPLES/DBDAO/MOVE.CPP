//  Sample MoveX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<conio.h>

void main(void)							// MoveX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstSupp;
	CdbBookmark		bmBook;
	COleVariant		vAbsPos, vRcdCt, vCo, vCity, vCntry;
	char			szInput[80] = {77};
	long			lMove;
	LPTSTR			szSelect = 
						_T("SELECT CompanyName, City, Country ")
						_T("FROM Suppliers ORDER BY CompanyName"),
					szMessage = 
						_T("Record (%ld) of %ld\n")
						_T("Company: %s\n")
						_T("Location: %s, %s\n\n")
						_T("Enter number of records to move ")
						_T("(positive or negative).\n")
						_T("(CR to exit)\n"),
					szTooFar = 
						_T("Too far %s! ")
						_T("Returning to current record.\n");
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		rstSupp = dbsNwind.OpenRecordset(szSelect, dbOpenDynaset);

	//	Populate recordset.
		rstSupp.MoveLast();
		rstSupp.MoveFirst();

		for(;;)
		{
		//	Display information about current record and ask how 
		//	many records to move.

			vAbsPos = rstSupp.GetAbsolutePosition();
			vRcdCt = rstSupp.GetRecordCount();
			vCo = rstSupp.GetField(_T("CompanyName"));
			vCity  = rstSupp.GetField(_T("City"));
			vCntry = rstSupp.GetField(_T("Country"));

			vAbsPos.ChangeType(VT_I4);
			vRcdCt.ChangeType(VT_I4);
			vCo.ChangeType(VT_BSTR);
			vCity.ChangeType(VT_BSTR);
			vCntry.ChangeType(VT_BSTR);

			printf(szMessage, 
						vAbsPos.lVal + 1, vRcdCt.lVal,
						vCo.bstrVal, 
						vCity.bstrVal, 
						vCntry.bstrVal);
		//	Move method requires parameter of datatype Long.
			if (1 != sscanf( _cgets( szInput ), _T("%ld"), &lMove))
				break;

		//	Store bookmark in case the Move doesn't work.
			bmBook = rstSupp.GetBookmark();
			rstSupp.Move( lMove );

		//	Trap for BOF and EOF.
			if (rstSupp.GetBOF())
			{
				printf(szTooFar, _T("backward"));
				rstSupp.SetBookmark( bmBook );
			}
			if (rstSupp.GetEOF())
			{
				printf(szTooFar, _T("forward"));
				rstSupp.SetBookmark( bmBook );
			}
		}

		rstSupp.Close();
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