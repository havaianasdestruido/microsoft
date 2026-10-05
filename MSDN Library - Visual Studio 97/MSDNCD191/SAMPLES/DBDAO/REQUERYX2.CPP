//  Sample RequeryX2
//
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// RequeryX2
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbQueryDef		qdfTemp;
	CdbRecordset	rstView, rstChange;
	CdbParameter	prmCountry;
	COleVariant		vUSA(_T("USA"), VT_BSTRT),
					vFirst, vLast, vCountry;
	LPCTSTR			szQuery = 
						_T("PARAMETERS ViewCoutry Text; ")
						_T("SELECT FirstName, LastName, Country ")
						_T("FROM Employees WHERE ")
						_T("Country = [ViewCountry]")
						_T("ORDER BY LastName");

	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));
		qdfTemp = dbsNwind.CreateQueryDef(_T(""), szQuery);
		prmCountry = qdfTemp.Parameters[_T("ViewCountry")];

		qdfTemp.Parameters[_T("ViewCountry")].SetValue(&vUSA);
		
		printf(_T("Data after initial query, [ViewCountry] = USA\n"));
		rstView = qdfTemp.OpenRecordset();

		while (!rstView.GetEOF())
		{
			vFirst = rstView.GetField(_T("FirstName"));
			vLast  = rstView.GetField(_T("LastName"));
			vCountry = rstView.GetField(_T("Country"));
			vFirst.ChangeType(VT_BSTR);
			vLast.ChangeType(VT_BSTR);
			vCountry.ChangeType(VT_BSTR);
			printf(_T("\t%s\t%s\t%s\n"), 
				(LPCTSTR)vFirst.bstrVal,
				(LPCTSTR)vLast.bstrVal,
				(LPCTSTR)vCountry.bstrVal);
			rstView.MoveNext();
		}

	//	Change query parameter.
		qdfTemp.Parameters[_T("ViewCountry")].SetValue(&COleVariant(_T("UK"), VT_BSTRT));

	//	QueryDef argument must be included so that the resulting recordset reflects
	//	the change in the query parameter.
		rstView.Requery(&qdfTemp);

		printf(_T("Requery after changing parameter, [ViewCountry] = UK\n"));
		rstView = qdfTemp.OpenRecordset();
		while (!rstView.GetEOF())
		{
			vFirst   = rstView.GetField(_T("FirstName"));
			vLast    = rstView.GetField(_T("LastName"));
			vCountry = rstView.GetField(_T("Country"));
			vFirst.ChangeType(VT_BSTR);
			vLast.ChangeType(VT_BSTR);
			vCountry.ChangeType(VT_BSTR);
			printf(_T("\t%s\t%s\t%s\n"), 
				(LPCTSTR)vFirst.bstrVal,
				(LPCTSTR)vLast.bstrVal,
				(LPCTSTR)vCountry.bstrVal);
			rstView.MoveNext();
		}

		rstView.Close();
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