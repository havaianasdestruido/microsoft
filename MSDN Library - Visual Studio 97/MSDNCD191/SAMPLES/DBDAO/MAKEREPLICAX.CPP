//  Sample MakeReplicaX
/*	Notes:
	1)	This function is used in a larger program. Assume 
		that a CdbDBEngine object has been started in 
		that program.  This declaration produces an 
		shared instance of that engine.
*/
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>
							// MakeReplicaX
long MakeAdditionalReplica(	LPCTSTR szReplicableDB,
							LPCTSTR	szNewReplica,
							long	lOptions)
{
	CdbDBEngine dben;		// note 1
	CdbDatabase	dbsTemp;
	long		lErr = 1;	// general error
	try 
	{
		dbsTemp = dben.OpenDatabase( szReplicableDB );
	/*	If you want no options specified in MakeReplica,
		pass -1L, which defaults to a full, read/write  
		replica. Otherwise, specify a valid value for 
		lOptions.
	*/
		dbsTemp.MakeReplica( 
			szNewReplica,
			(strcat(_T("Replica of "), szReplicableDB)),
			lOptions);
		dbsTemp.Close();
		lErr = 0;			// success
	}
	catch (CdbException ex)
	{
		for (long ct = 0; ct < dben.Errors.GetCount(); ct++)
		{
			printf(_T("\t Error #%ld: #%ld -- %s\n"),
				ct,
				DBERR(dben.Errors[ct].GetNumber()),
				dben.Errors[ct].GetDescription());
		}
		lErr = (long)ex.m_hr;	// specific error
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
	return lErr;				// return status
}

