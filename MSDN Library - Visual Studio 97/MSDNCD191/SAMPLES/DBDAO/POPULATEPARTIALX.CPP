//  Sample PopulatePartialX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)				// PopulatePartialX
{
	CdbDBEngine dben;		
	CdbDatabase	dbsTemp;
	CdbTableDef	tdfCustomers;
	COleVariant	vFilter(_T("Region = 'CA'"), VT_BSTRT);
	LPCTSTR		szSalesFY96   = _T("C:\\SALES\\FY96.MDB"),
				szSalesFY96CA = _T("F:\\SALES\\FY96CA.MDB"),
				szCustomers   = _T("Customers");
	//
	try						
	{
	//	Open the partial replica in exclusive mode.
		dbsTemp = dben.OpenDatabase(szSalesFY96CA, TRUE);
		tdfCustomers = dbsTemp.TableDefs[szCustomers];

	//	Synchronize with full replica before setting replica filter.
		dbsTemp.Synchronize(szSalesFY96);
		
		tdfCustomers.SetReplicaFilter( &vFilter );

	//	Populate records from the full replica.
		dbsTemp.PopulatePartial(szSalesFY96);
		dbsTemp.Close();
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