//  Sample RefreshX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)						// RefreshX
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbTableDef		tdfEmp;
	CdbField		fld;
	COleVariant		vOrd, vName, vOrdTemp, vNameTemp;
	SAFEARRAYBOUND	sab;
	long			lCt, lMaxCt;
	short			sOrd;
	//
	try
	{	
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		tdfEmp = dbsNwind.TableDefs[_T("Categories")];

	//	Display original OrdinalPosition data and store it 
	//	in an array.
		printf(_T("Original OrdinalPosition data in TableDef.\n"));

		lMaxCt = tdfEmp.Fields.GetCount();
		sab.cElements = lMaxCt;
		vOrd.vt  = VT_I2   | VT_ARRAY;
		vName.vt = VT_BSTR | VT_ARRAY;
		vOrd.parray  = SafeArrayCreate(	VT_I2,   1, &sab);
		vName.parray = SafeArrayCreate(	VT_BSTR, 1, &sab);

		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			vOrdTemp  = tdfEmp.Fields[lCt].GetOrdinalPosition();
			SafeArrayPutElement(vOrd.parray, &lCt, &vOrdTemp );

			vNameTemp = tdfEmp.Fields[lCt].GetName();
			SafeArrayPutElement(vName.parray, &lCt, &vNameTemp );

			printf(_T("\t%d\t%s\n"), vOrdTemp.iVal, vNameTemp.bstrVal);
		}

	//	Change OrdinalPosition data.
		lMaxCt = tdfEmp.Fields.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			fld = tdfEmp.Fields[lCt];
			sOrd = fld.GetOrdinalPosition();
			fld.SetOrdinalPosition(100 - sOrd);
		}

	//	Print new data.
		printf(_T("New OrdinalPosition data before Refresh.\n"));
		lMaxCt = tdfEmp.Fields.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			fld = tdfEmp.Fields[lCt];
			printf(_T("\t%d\t%s\n"), 
				fld.GetOrdinalPosition(), (LPCTSTR)fld.GetName());
		}

	//	Do the Refresh.
		tdfEmp.Fields.Refresh();

	//	Print new data, showing how the field order has been changed.
		printf(_T("New OrdinalPosition data after Refresh.\n"));
		lMaxCt = tdfEmp.Fields.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			fld = tdfEmp.Fields[lCt];
			printf(_T("\t%d\t%s\n"), 
				fld.GetOrdinalPosition(), (LPCTSTR)fld.GetName());
		}

	//	Restore original OrdinalPosition data.
		lMaxCt = tdfEmp.Fields.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			SafeArrayGetElement(vOrd.parray, &lCt, &vOrd);
			SafeArrayGetElement(vName.parray,&lCt, &vName);
			tdfEmp.Fields[(LPCTSTR)vName.bstrVal].SetOrdinalPosition( vOrd.iVal );
		}

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