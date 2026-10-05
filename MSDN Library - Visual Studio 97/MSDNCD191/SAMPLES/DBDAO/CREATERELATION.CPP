//  Sample CreateRelationX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// CreateRelationX()
{
	CdbDBEngine		dben;
	CdbDatabase		dbsNwind;
	CdbTableDef		tdfEmp, tdfNew;
	CdbIndex		idxNew, idx;
	CdbRelation		relNew;
	CdbField		fld;
	CdbProperty		prp;
	long			lCt, lMaxCt;
	//
	try
	{
		dbsNwind = dben.OpenDatabase(_T("Northwind.mdb"));

	//	Add new field to Employees table.

		tdfEmp = dbsNwind.TableDefs[_T("Employees")];
		fld = tdfEmp.CreateField(_T("DeptID"), dbInteger, 2);
		tdfEmp.Fields.Append( fld );

	//	Create new Departments table.

		tdfNew = dbsNwind.CreateTableDef(_T("Departments"));

	//	Create and append Field objects to Fields collection
	//	of the new TableDef object.

		fld = tdfNew.CreateField(_T("DeptID"), dbInteger, 2);
		tdfNew.Fields.Append( fld );
		fld = tdfNew.CreateField(_T("DeptName"), dbText, 20);
		tdfNew.Fields.Append( fld );

	//	Create an Index object for the Departments table.
		idxNew = tdfNew.CreateIndex(_T("DeptIDIndex"));

	//	Create and append Field object to Fields collection
	//	of the new Index object.

		fld = idxNew.CreateField(_T("DeptID"));
		idxNew.Fields.Append( fld );

	//	The index in the primary table must be Unique in order
	//	to be part of a Relation.

		idxNew.SetUnique(TRUE);
		tdfEmp.Indexes.Append( idxNew );
		
		dbsNwind.TableDefs.Append( tdfNew );

	//	Create EmployeesDepartments Relation object using
	//	the names of the two tables in the relation.

		relNew = dbsNwind.CreateRelation(_T("EmployeesDepartments"),
					(LPCTSTR)tdfNew.GetName(),
					(LPCTSTR)tdfEmp.GetName(),
					dbRelationUpdateCascade);

	//	Create Field object for the Fields collection of the new
	//	Relation object.  Set the Name and ForeignName properties
	//	based on the fields to be used for the relation.

		fld = relNew.CreateField(_T("DeptID"));
		relNew.Fields.Append( fld );
		relNew.Fields[_T("DeptID")].SetForeignName(_T("DeptID"));
		dbsNwind.Relations.Append( relNew );

	//	Print report.
		printf(_T("Properties of %s Relation\n"), 
			(LPCTSTR)relNew.GetName());
		printf(_T("\tTable = %s\n"), 
			(LPCTSTR)relNew.GetTable());
		printf(_T("\tForeignTable = %s\n"), 
			(LPCTSTR)relNew.GetForeignTable());
		
		printf(_T("Fields of %s Relation\n"), 
			(LPCTSTR)relNew.GetName());

		fld = relNew.Fields[_T("DeptID")];
		printf(_T("\t%s\n"), (LPCTSTR)fld.GetName());
		printf(_T("\t\tName = %s\n"), (LPCTSTR)fld.GetName());
		printf(_T("\t\tForeignName = %s\n"), (LPCTSTR)fld.GetForeignName());

		printf(_T("Indexes in %s TableDef\n"), (LPCTSTR)tdfEmp.GetName());
		lMaxCt = tdfEmp.Indexes.GetCount();
		for (lCt = 0; lCt < lMaxCt; lCt++)
		{
			idx = tdfEmp.Indexes[lCt];
			printf(_T("\t%s, Foreign = %s\n"),
				(LPCTSTR)idx.GetName(),
				(idx.GetForeign() ? "TRUE":"FALSE"));
		}

	//	Delete new objects becuse this is a demonstration.
		dbsNwind.Relations.Delete((LPCTSTR)relNew.GetName());
		dbsNwind.TableDefs.Delete((LPCTSTR)tdfNew.GetName());
		tdfEmp.Fields.Delete(_T("DeptID"));
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