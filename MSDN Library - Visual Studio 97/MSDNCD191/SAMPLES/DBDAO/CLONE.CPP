//  Sample CloneX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<conio.h>

void main(void)							// CloneX()
{
	CdbDBEngine		dben;
	CdbDatabase		db;
	CdbRecordset	arst[3];			// array of recordsets
	long			lCt;
	CString			strFind, 
					strMessage;
	char			szInput[80] = {77},
					szData[80];
	COleVariant		vVar[3];
	LPTSTR			lpszSelect = 
						_T("SELECT ProductName FROM Products ORDER BY Productname"),
					lpszStr =	
						_T("\nRecordsets from Products table:\n")
						_T(" 1 - Original - record pointer at %s\n")
						_T(" 2 - Clone    - record pointer at %s\n") 
						_T(" 3 - Clone    - record pointer at %s\n")
						_T("Enter search string for #%ld [null to exit]: ");
	//
	try
	{	db = dben.OpenDatabase(_T("northwind.mdb"));

	//	If the following SQL statement will be used often, creating a  
	//	permanent QueryDef will result in better performance.
		arst[0] = db.OpenRecordset(lpszSelect, dbOpenSnapshot);

	//	Create two clones of the original recordset.
		arst[1] = arst[0].Clone();
		arst[2] = arst[0].Clone();

		for (;;) 
		{
		//	Loop through the array so that on each pass the user is 
		//	searching a different copy of the same recordset.
			for (lCt = 0; lCt < 3; lCt++) 
			{
			//	Ask for search string while showing where the current 
			//	record pointer is for each recordset.
				vVar[0] = arst[0].GetField(_T("ProductName"));
				vVar[1] = arst[1].GetField(_T("ProductName"));
				vVar[2] = arst[2].GetField(_T("ProductName"));

				printf(	lpszStr, 
						(LPTSTR) vVar[0].bstrVal, 
						(LPTSTR) vVar[1].bstrVal, 
						(LPTSTR) vVar[2].bstrVal,
						lCt+1);
				if (1 != sscanf( _cgets( szInput ), _T("%77s"), szData))
					goto exitLoop;
				strFind = szData;
				strMessage = _T("ProductName >= '") + strFind + _T("'");

			//	Find the search string; if there's no match, jump 
			//	to the last record.
				arst[lCt].FindFirst( (LPCTSTR)strMessage );
				if (arst[lCt].GetNoMatch()) arst[lCt].MoveLast();
			}
			printf(_T("\n"));
		}
exitLoop:
		arst[0].Close();
		arst[1].Close();
		arst[2].Close();
		db.Close();
	//
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