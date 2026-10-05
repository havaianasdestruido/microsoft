//  Sample CreateTableDefX
/*	Notes: 
	1)	(LPCTSTR) is really a unary operator that returns
		a pointer to a C string.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)						// CreateTableDefX
{
	CdbDBEngine	dben;
	CdbDatabase	dbsNwind;
	CdbTableDef	tdfNew;
	CdbProperty	prp;
	CdbField	fld;
	COleVariant	vVal;
	long		lCt, lMaxCt;
	char		szTitle[] = 
					_T("Properties of new TableDef object")
					_T("%s appending to collection:\n");
	//
	try
	{	dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));

	//	Create a new TableDef object.
		tdfNew = dbsNwind.CreateTableDef(_T("Contacts"));

	/*	Create fields and append them to the new TableDef object. This must
		be done before appending the TableDef object to the TableDefs collection
		of the Northwind database. 
	*/
		fld = tdfNew.CreateField(_T("FirstName"), dbText);
		tdfNew.Fields.Append( fld );
		fld = tdfNew.CreateField(_T("LastName"), dbText);
		tdfNew.Fields.Append( fld );
		fld = tdfNew.CreateField(_T("Phone"), dbText);
		tdfNew.Fields.Append( fld );
		fld = tdfNew.CreateField(_T("Notes"), dbMemo);
		tdfNew.Fields.Append( fld );

	//	Enumerate Properties collection of the new TableDef object.
		printf(szTitle, _T("before"));
		lMaxCt = tdfNew.Properties.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			try 
			{	
				prp = tdfNew.Properties[lCt];
				vVal = prp.GetValue();
				if (vVal.vt != VT_EMPTY)
				{
				vVal.ChangeType(VT_BSTR);
				printf(_T("\t%s = %S\n"),			// note 1
					(LPCTSTR)prp.GetName(), vVal.bstrVal);	
				}
			}
			catch(CdbException) {}
		}

	//	Append the new TableDef object to the Northwind database.
		dbsNwind.TableDefs.Append( tdfNew);
		
	//	Enumerate Properties collection of the new TableDef object.
		printf(szTitle, _T("after"));
		lMaxCt = tdfNew.Properties.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			try 
			{	
				prp = tdfNew.Properties[lCt];
				vVal = prp.GetValue();
				if (vVal.vt != VT_EMPTY)
				{
				vVal.ChangeType(VT_BSTR);
				printf(_T("\t%s = %S\n"),			// note 1
					(LPCTSTR)prp.GetName(), vVal.bstrVal);	
				}
			}
			catch(CdbException) {}
		}

	//	Delete new TableDef object since this is a demonstration.
		dbsNwind.TableDefs.Delete(_T("Contacts"));
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