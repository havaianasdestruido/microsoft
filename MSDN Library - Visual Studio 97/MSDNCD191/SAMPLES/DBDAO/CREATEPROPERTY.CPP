//  Sample CreatePropertyX
/*	Note:
	1)	A try-catch block is explicitly required here. Since  
		it's available, we can re-throw the exception to the  
		outer handler.
	2)	We never reach this point if an exception is thrown.
	3)	This creates a shared instance of the engine.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void SetProperty(	CdbDatabase dbsTemp, 
					LPTSTR		szName,
					BOOL		bTemp)
{
	CdbProperty	prp;
	COleVariant	vVal((short)bTemp, VT_BOOL);
	//
	try										// note 1
	{
	//	Attempt to set the specified property.
		prp = dbsTemp.Properties[ szName ];
		prp.SetValue( &vVal );				// note 2
	}
	catch (CdbException)
	{
		CdbDBEngine	dben;					// note 3

		//	Error (3270) means the property was not found.
		if (dben.Errors[0L].GetNumber() == E_DAO_VtoPropNotFound)
		{
		//	Create the propety, set its value, and append   
		//	it to the Properties collection.
			prp = dbsTemp.CreateProperty( szName, 
					dbBoolean, &vVal );
			dbsTemp.Properties.Append( prp );
		}
		else
		{
		//	If different error has occurred, display message.
			throw;
		}
	}
}

void main(void)								// CreatePropertyX()
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbProperty		prp;
	COleVariant		vVal;
	long			lCt, lMaxCt;
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("northwind.mdb"));
		
	//	Set the Archive property to TRUE.
		SetProperty( dbsNwind, _T("Archive"), TRUE);

	//	Enumerate Properties collection of the Northwind database.
		printf(_T("Properties of %s\n"), (LPCTSTR)dbsNwind.GetName());
		lMaxCt = dbsNwind.Properties.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			prp = dbsNwind.Properties[ lCt ];
			vVal = prp.GetValue();
			if (vVal.vt != VT_EMPTY)
			{
				vVal.ChangeType(VT_BSTR);
				printf(_T("\t%s = '%S'\n"), 
					(LPCTSTR)prp.GetName(), vVal.bstrVal);
			}
		}
	//	Delete new property since this is a demonstration.
		dbsNwind.Properties.Delete(_T("Archive"));
		
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