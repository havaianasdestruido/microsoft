//  Sample CreateFieldX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)										// CreateFieldX
{
	CdbDBEngine		dben;
	CdbDatabase		dbs;
	CdbTableDef		tdf;
	CdbField		fld;
	CdbProperty		prp;
	long			lNumOfPrp,
					lNumOfFld;
	COleVariant		vVal;
	//
	try
	{
		dbs = dben.OpenDatabase(_T("Northwind.mdb"));
		tdf = dbs.CreateTableDef(_T("NewTableDef"));
		
		// Create and append new Field objects for the new TableDef object.
		// The CreateField method will set a default size if one is not specified.

		fld = tdf.CreateField(_T("TextField"), dbText);
		tdf.Fields.Append( fld );
		fld = tdf.CreateField(_T("IntegerField"), dbInteger);
		tdf.Fields.Append( fld );
		fld = tdf.CreateField(_T("DateField"), dbDate);
		tdf.Fields.Append( fld );

		dbs.TableDefs.Append( tdf );

		printf(_T("Properties of new Fields in %s\n"), (LPCTSTR)tdf.GetName());
		lNumOfFld = tdf.Fields.GetCount();
		for (long lFldNdx = 0; lFldNdx < lNumOfFld; lFldNdx++)
		{
			fld = tdf.Fields[ lFldNdx ];
			printf(_T("\t%s:\n"), (LPCTSTR)fld.GetName());
	
			lNumOfPrp = fld.Properties.GetCount();
			for (long lPrpNdx = 0; lPrpNdx < lNumOfPrp; lPrpNdx++)
			{
				prp = fld.Properties[ lPrpNdx ];

		//	Properties that are invalid in the context of TableDefs will 
		//	trigger an error if an attempt is made to use their values.
				try
				{
					vVal = prp.GetValue();
					printf(_T("\t\t%s - "), (LPCTSTR)prp.GetName());
					if (vVal.vt == VT_EMPTY)
						printf(_T("[empty]\n"));
					else
					{
						vVal.ChangeType(VT_BSTR);
						printf(_T("'%S'\n"), vVal.bstrVal);
					}
				}
				catch (CdbException)
				{
				}
			}
		}

		// Delete new TableDef since this is a demonstration.

		dbs.TableDefs.Delete(tdf.GetName());
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