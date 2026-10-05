//  Sample IdleX
/*	Notes:
	1)	To match the corresponding VB example, we need the Microsoft Jet
		Engine to be globally accessible.
	2)	_cgets() is allowed to read a maximum of 77 characters.
	3)	_cgets() returns console input up to the maximum specified in
		szInput. sscanf() places 1 non-blank string in szData.
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
#include	<stdlib.h>
#include	<conio.h>

CdbDBEngine		dben;							// note 1

void IdleOutput( CdbRecordset rst,
				 CString	  strTemp)
{
	COleVariant	vData;

	//	Call the Idle method to release unneeded locks, force pending writes,
	//	and refresh the memory with the current data in the .mdb file.

	dben.Idle( dbRefreshCache );

	//	Enumerate the Recordset object.
	printf(_T("Orders from %s: \n"), (LPCTSTR) strTemp);
	printf(_T("\tOrderID\tCustomerID\tOrderDate\n"));
		
	while( !rst.GetEOF() )
	{
		vData = rst.GetField(_T("OrderID"));
		vData.ChangeType(VT_BSTR);
		printf(_T("\t%s"), vData.bstrVal);

		vData = rst.GetField(_T("CustomerID"));
		vData.ChangeType(VT_BSTR);
		printf(_T("\t%s"), vData.bstrVal);

		vData = rst.GetField(_T("OrderDate"));
		vData.ChangeType(VT_BSTR);
		printf(_T("\t%s\n"), vData.bstrVal);

		rst.MoveNext();
	}
}

void main(void)									// IdleX
{
	CdbDatabase		dbs;
	CdbRecordset	rstOrders;
	char			szInput[80] = {77},			// note 2
					szData[80];	
	CString			strCountry, strSQL;
	//
	try
	{
		dbs = dben[0L].OpenDatabase(_T("Northwind.mdb"));
			
		//	Get the name of the country from the user and build SQL
		//	statement with it.
		printf(_T("\nEnter country: "));
		if (1 == sscanf( _cgets( szInput ), _T("%77s"), szData))	
			strCountry = szData;				// note 3
		printf(_T("\n"));

		strSQL = _T("SELECT * FROM Orders WHERE ShipCountry = '") +
			strCountry + _T("' ORDER BY OrderID");

		//	Open Recordset object with SQL statement.
		rstOrders = dbs.OpenRecordset(strSQL);

		//	Display contents of Recordset object.
		IdleOutput( rstOrders, strCountry );

		rstOrders.Close();
		dbs.Close();

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
