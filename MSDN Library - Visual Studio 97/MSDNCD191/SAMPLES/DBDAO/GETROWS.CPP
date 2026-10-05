//  Sample GetRowsX
/*	Notes:
	1)	GetRows() returns a rectangular array of records. Each 
		record consists of 3 fields (FirstName, LastName, Title) 
		numbered 0 to 2. There are at best lNumber records 
		numbered 0 to lNumber-1. 
	2)	The upper bound of the second dimension plus 1 is the number
		of records returned.
	3)	lIndex[0] accesses each field of each record.
		lIndex[1] accesses each record.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

BOOL GetRowsOK(	CdbRecordset rstTemp,
				long		 lNumber,
				COleVariant	 vArray)
{
	long	lUBound;

//	Store results of GetRows method in array.
	vArray = rstTemp.GetRows( lNumber );	// note 1

//	Return FALSE only if fewer than the desired number
//	of rows were returned, but not because of the end 
//	of the recordset was reached.
											// note 2
	SafeArrayGetUBound( vArray.parray, 2, &lUBound);
	return !((lNumber > (lUBound+1)) && (!rstTemp.GetEOF()));
}

void main(void)								// GetRowsX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstEmp;
	COleVariant		vFirst, vLast, vTitle, vRecords;
	long			lIndex[2], lRows, lUBound;
	char			chYN,
					szSelect[] = 
						_T("SELECT FirstName, LastName, Title  ")
						_T("FROM Employees ORDER BY LastName"),
					szMessage[] = 
						_T("Enter number of rows to retrieve.\n")
						_T("(enter 0 to exit)\n"),
					szNotEnough[] = 
						_T("(Not enough records in recordset to ")
						_T("retrieve %ld rows.)\n"),
					szFailed[] = 
						_T("GetRows failed!\n"),
					szNotRestart[] = 
						_T("Recordset not Restartable!\n"),
					szRetry[] =
						_T("GetRows failed--retry?\n"),
					szRcdsFound[] = 
						_T("%ld records found.\n");
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		rstEmp = dbsNwind.OpenRecordset(szSelect, dbOpenSnapshot);

		for(;;)
		{
		//	Get user input for number of rows.

			printf(szMessage);
			scanf(_T("%ld"), &lRows);
			if (lRows == 0) break;

		//	If GetRowsOK is successful, print the results noting
		//	if the end of file was reached.

			if (GetRowsOK( rstEmp, lRows, vRecords))
			{
				SafeArrayGetUBound(vRecords.parray, 2, &lUBound);
				lUBound++;
				if (lRows > lUBound)
					printf(szNotEnough, lRows);
				printf(szRcdsFound, lUBound);

			//	Print the retrieved records.		// note 3
				for (lIndex[1] = 0; lIndex[1] < lUBound; lIndex[1]++)
				{
					lIndex[0] = 0;	
					SafeArrayGetElement(vRecords.parray, lIndex, &vFirst);
					lIndex[0] = 1;
					SafeArrayGetElement(vRecords.parray, lIndex, &vLast);
					lIndex[0] = 2;
					SafeArrayGetElement(vRecords.parray, lIndex, &vTitle);
					printf(_T("\t%s %s, %s\n"),
						vFirst.bstrVal, vLast.bstrVal, vTitle.bstrVal);
				}
			}
			else
			{
			//	Assuming the GetRows error was due to data changes by another
			//	user, use Requery to refresh the recordset and start over.

				if (rstEmp.GetRestartable())
				{
					printf(szRetry);
					scanf(_T("%c"), &chYN);
					if (toupper(chYN) == _T('Y')) 
						rstEmp.Requery();
					else
					{
						printf(_T("%s\n"), szFailed);
						break;
					}
				}
				else
				{
					printf(_T("%s%s\n"), szFailed, szNotRestart);
					break;
				}
			}

		//	Because using GetRows leaves the current record pointer at
		//	the last record accessed, move pointer back to the beginning
		//	of the recordset before looping back for another search.
			rstEmp.MoveFirst();
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