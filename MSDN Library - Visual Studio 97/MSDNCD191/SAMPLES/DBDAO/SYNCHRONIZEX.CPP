//  Sample SynchronizeX
/*	Notes:
	1) This sample combines all four of the VB examples.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
							// note 1
void main(void)				// SynchronizeX
{
	CdbDBEngine dben;		
	CdbDatabase	dbsNwind, dbsTemp;
	try						
	{
	//	SendChangeToReplicaX:
	//	Opens the replicable database, Northwind.mdb.
		dbsNwind = dben.OpenDatabase( _T("Northwind.mdb") );
	//	Sends data or structural changes to replica.
		dbsNwind.Synchronize(_T("Nwreplica.mdb"),
			dbRepExportChanges);
		dbsNwind.Close();

	//	ReceiveChangeX:
		dbsNwind = dben.OpenDatabase( _T("Northwind.mdb") );
	//	Sends changes from replica to design Master.
		dbsNwind.Synchronize(_T("Nwreplica.mdb"),
			dbRepImportChanges);
		dbsNwind.Close();

	//	TwoWayExchangeX:
		dbsNwind = dben.OpenDatabase( _T("Northwind.mdb") );
	//	Sends changes made in each replica to the other.
		dbsNwind.Synchronize(_T("Nwreplica.mdb"),
			dbRepImpExpChanges);
		dbsNwind.Close();

	//	InternetSynchronizeX:
		dbsTemp = dben.OpenDatabase( _T("C:\\Data\\ordEntry.mdb") );
	//	Synchronize the local database with the replica on
	//	the internet server.
		dbsTemp.Synchronize(
			_T("www.mycompany.myserver.com\\files\\Orders.mdb"),
			(dbRepImpExpChanges | dbRepSyncInternet));
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

