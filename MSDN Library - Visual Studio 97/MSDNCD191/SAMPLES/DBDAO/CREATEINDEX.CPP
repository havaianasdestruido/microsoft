//  Sample CreateIndexX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void CreateIndexOutput(CdbIndex idxTemp)
{
	CdbField	fld;
	CdbProperty	prp;
	COleVariant	vVal;
	long		lCt, lMaxCt;

//	Enumerate Fields collection of Index object.
	printf(_T("Fields in %s\n"), (LPCTSTR)idxTemp.GetName());
	lMaxCt = idxTemp.Fields.GetCount();
	for (lCt = 0; lCt < lMaxCt; lCt++)
	{
		fld = idxTemp.Fields[ lCt ];
		printf(_T("\t%s\n"), (LPCTSTR)fld.GetName());
	}

//	Enumerate Properties collection of Index object.
	printf(_T("Properties in %s\n"), (LPCTSTR)idxTemp.GetName());
	lMaxCt = idxTemp.Properties.GetCount();
	for (lCt = 0; lCt < lMaxCt; lCt++)
	{
		prp = idxTemp.Properties[ lCt ];
		vVal = prp.GetValue();
		if (vVal.vt == VT_EMPTY)
			printf(_T("\t%s - [empty]\n"), (LPCTSTR)prp.GetName());											// note 3
		else
		{
			vVal.ChangeType(VT_BSTR);
			printf(_T("\t%s - '%S'\n"), 
				(LPCTSTR)prp.GetName(), vVal.bstrVal);
		}
	}
}

void main(void)										// CreateIndexX()
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbTableDef		tdfEmp;
	CdbIndex		idxCntry, idxFirst, idx;
	CdbField		fld;
	CdbProperty		prp;
	long			lCt, lMaxCt;
	COleVariant		vVal;
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("northwind.mdb"));
		tdfEmp = dbsNwind.CreateTableDef(_T("NewTableDef"));
		
	//	Create first Index object, create and append Field objects
	//	to the Index object, and then append the Index object to
	//	the Index collection of the TableDef.

		idxCntry = tdfEmp.CreateIndex(_T("CountryIndex"));
		fld = idxCntry.CreateField(_T("Country"));
		idxCntry.Fields.Append( fld );
		fld = idxCntry.CreateField(_T("LastName"));
		idxCntry.Fields.Append( fld );
		fld = idxCntry.CreateField(_T("FirstName"));
		idxCntry.Fields.Append( fld );
		tdfEmp.Indexes.Append( idxCntry );

	//	Create second Index object, create and append Field objects to the
	//	Index object, and then append the Index object to the Indexes
	//	collection of the TableDef.

		idxFirst = tdfEmp.CreateIndex();
		idxFirst.SetName(_T("FirstNameIndex"));
		fld = idxFirst.CreateField(_T("FirstName"));
		idxFirst.Fields.Append( fld );
		fld = idxFirst.CreateField(_T("LastName"));
		idxFirst.Fields.Append( fld );
		tdfEmp.Indexes.Append( idxFirst );

	//	Refresh collection so that you can access new Index objects.
		tdfEmp.Indexes.Refresh();

	//	Enumerate Indexes collection.
		printf(_T("%ld Indexes in %s TableDef\n"), 
			tdfEmp.Indexes.GetCount(), (LPCTSTR)tdfEmp.GetName());
		lMaxCt = tdfEmp.Indexes.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			idx = tdfEmp.Indexes[ lCt ];
			printf(_T("\t%s\n"), (LPCTSTR)idx.GetName());
		}

		CreateIndexOutput( idxCntry );
		CreateIndexOutput( idxFirst );

	//	Delete new Index objects because this is a demonstration.
		tdfEmp.Indexes.Delete( (LPCTSTR)idxCntry.GetName());
		tdfEmp.Indexes.Delete( (LPCTSTR)idxFirst.GetName());
		
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