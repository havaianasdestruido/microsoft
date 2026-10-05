//  Sample RequeryX1
//
#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// RequeryX1
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbQueryDef		qdfTemp;
	CdbRecordset	rstView, rstChange;
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

	//	Change underlying data.
		rstChange = dbsNwind.OpenRecordset(_T("Employees"));
		rstChange.AddNew();

		vFirst.SetString(_T("Nina"), VT_BSTRT);
		rstChange.SetField(_T("FirstName"), &vFirst);

		vLast.SetString(_T("Roberts"), VT_BSTRT);
		rstChange.SetField(_T("LastName"), &vLast);

		vCountry.SetString(_T("USA"), VT_BSTRT);
		rstChange.SetField(_T("Country"), &vCountry);

		rstChange.Update();
		rstView.Requery();

		printf(_T("Requery after changing underlying data\n"));
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

	//	Restore original data because this is only a demonstration.
		rstChange.SetBookmark( rstChange.GetLastModified());
		rstChange.Delete();
		rstChange.Close();

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