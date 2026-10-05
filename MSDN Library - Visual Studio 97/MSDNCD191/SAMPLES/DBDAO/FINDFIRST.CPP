//  Sample FindFirstX
//
/*	Notes:
	1)	(LPCTSTR) is really a unary operator that takes  
		a CString and returns a pointer to a string.
	2)	This corresponds to the VB sample. You may want to 
		optimize.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<conio.h>
//
BOOL FindAny(int nChoice, CdbRecordset rstTemp,
			 CString strFind)
{
//	Use Find method based on user input.
	switch (nChoice) 
	{
	case 1:
		rstTemp.FindFirst( (LPCTSTR)strFind );	// note 1
		break;
	case 2:
		rstTemp.FindLast( (LPCTSTR)strFind );
		break;
	case 3:
		rstTemp.FindNext( (LPCTSTR)strFind );
		break;
	case 4:
		rstTemp.FindPrevious( (LPCTSTR)strFind );
		break;
	default:
		break;
	}
//	Set return value based on NoMatch property.
	return( !rstTemp.GetNoMatch());
}

void main(void)							// FindFirstX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstCust;
	CString			strCountry;
	CdbBookmark		bmBook;
	COleVariant		vCo, vCity, vCntry;
	int				nCmd;
	char			szInput[80] = {77},
					szData[80];
	LPTSTR			szSelect = 
						_T("SELECT CompanyName, City, Country ")
						_T("FROM Customers ORDER BY CompanyName"),
					szMessage = 
						_T("Company: %s\n")
						_T("Location: %s, %s\n\n")
						_T("%s\n\n")			// Search criteria
						_T("[1 - FindFirst, 2 - FindLast, \n")
						_T(" 3 - FindNext,  4 - FindPrevious]\n")
						_T("(Any other digit to exit)\n");
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		rstCust = dbsNwind.OpenRecordset(szSelect, dbOpenSnapshot);

		for(;;)									// note 2
		{
		//	Get user input and build search string.
			printf(_T("Enter country to search (CR to exit):\n"));
			if (1 == sscanf( _cgets( szInput ), _T("%77s"), szData))
				strCountry = szData;
			else
				break;
			strCountry = _T("Country = '") + strCountry + _T("'");
		
		//	Populate recordset.
			rstCust.MoveLast();
		
		//	Find first record satisfying search string. 
		//	Exit loop if no such record exists.

			rstCust.FindFirst( (LPCTSTR)strCountry );
			if (rstCust.GetNoMatch())
			{
				printf(_T("No records found with %s.\n"), (LPCTSTR)strCountry);
				break;
			}

			for(;;)
			{
			//	Store bookmark of current record.
				bmBook = rstCust.GetBookmark();

			//	Get user choice of which method to use.
				vCo = rstCust.GetField(_T("CompanyName"));
				vCity  = rstCust.GetField(_T("City"));
				vCntry = rstCust.GetField(_T("Country"));

				vCo.ChangeType(VT_BSTR);
				vCity.ChangeType(VT_BSTR);
				vCntry.ChangeType(VT_BSTR);

				printf(	szMessage, 
						vCo.bstrVal, vCity.bstrVal, vCntry.bstrVal, 
						(LPCTSTR)strCountry );
				scanf(_T("%d"), &nCmd);
				if ((nCmd < 1) || (nCmd > 4)) break;

			//	Use selected Find method. If the Find fails
			//	return to the last current record.
				if (FindAny( nCmd, rstCust, strCountry) == FALSE)
				{
					rstCust.SetBookmark( bmBook );
					printf(_T("No match--returning to current record.\n"));
				}
			}
			break;
		}

		rstCust.Close();
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