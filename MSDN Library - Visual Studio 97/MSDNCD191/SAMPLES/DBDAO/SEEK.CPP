//  Sample SeekX
/*	Notes:
	1)	ProductID is an AutoNumber (long) field.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// SeekX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbRecordset	rstProd;
	CdbBookmark		bmBook;
	COleVariant		vProdID, vProdName, vVal;
	long			lFirst, lLast, lVal = 0;
	LPTSTR			szMessage = 
						_T("Product ID: %S\n")
						_T("Name: %s\n\n")
						_T("Enter a product ID between %ld and %ld\n")
						_T("(enter 0 to exit)\n");
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
	
	//	You must open a table-type recordset to use an index, and hence 
	//	the Seek method.
		rstProd = dbsNwind.OpenRecordset(_T("Products"), dbOpenTable);

	//	Set the index.
		rstProd.SetIndex(_T("PrimaryKey"));

	//	Get the highest and lowest product IDs.
		rstProd.MoveLast();
		vVal = rstProd.GetField(_T("ProductID"));
		vVal.ChangeType(VT_I4);
		lLast = vVal.lVal;

		rstProd.MoveFirst();
		vVal = rstProd.GetField(_T("ProductID"));
		vVal.ChangeType(VT_I4);
		lFirst = vVal.lVal;

	//	Display current record information and ask the user for 
	//	an ID number.
		for(;;)
		{
			vProdID = rstProd.GetField(_T("ProductID"));
			vProdName  = rstProd.GetField(_T("ProductName"));

			vProdID.ChangeType(VT_BSTR);
			vProdName.ChangeType(VT_BSTR);

			printf(szMessage, 
						vProdID.bstrVal,
						vProdName.bstrVal,
						lFirst, lLast);
 			scanf(_T("%ld"), &lVal);
			if (lVal == 0) break;
			vVal = COleVariant(lVal, VT_I4);

		//	Store bookmark in case the Seek fails.
			bmBook = rstProd.GetBookmark();
			rstProd.Seek( _T("="), 1, vVal );	// note 1

		//	Return to the current record if the Seek fails.
			if (rstProd.GetNoMatch())
			{
				printf(_T("ID not found\n"));
				rstProd.SetBookmark( bmBook );
			}
		}

		rstProd.Close();
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