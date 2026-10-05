//  Sample AppendChunkX
/*	Notes:
	1)	This field is an OLE Object so we must use FieldSize() instead of
		GetSize().
	2)	We'll diverge from the VB sample to use the more compact C idiom.
	3)	This sample removes any evidence that it worked. But if you have
		a tool like Microsoft Access, you can comment out Delete() then 
		use the tool to examine the database. When you're done, use the
		tool to manually delete the new record.
*/

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

//	Set size of chunk in bytes.
const long lChunkSize = 32768;

void CopyLargeField( CdbField fldSource,
					 CdbField fldDestination )
{
//	Copy the photo from one recordset to the other in 32K chunks 
//	until the entire field is copied.
	long	lOffset,
			lTotalSize = fldSource.FieldSize();	// note 1
												// note 2
	for ( lOffset = 0; lOffset < lTotalSize; lOffset += lChunkSize ) 
		fldDestination.AppendChunk( &(fldSource.GetChunk(lOffset, lChunkSize)) );
}

void main(void)									// AppendChunkX
{
	CdbDBEngine		dben;
	CdbDatabase		db;
	CdbRecordset	rsEmp1,
					rsEmp2;
	//
	try
	{
		db = dben.OpenDatabase(_T("northwind.mdb"));
		rsEmp1 = db.OpenRecordset(_T("Employees"), dbOpenDynaset);
		rsEmp2 = rsEmp1.Clone();

	//	Add a new record to the first recordset and copy the data from
	//	a record in the second recordset.
		rsEmp1.AddNew();

		rsEmp1.SetField(_T("FirstName"), &(rsEmp2.GetField(_T("FirstName"))) );
		rsEmp1.SetField(_T("LastName"),  &(rsEmp2.GetField(_T("LastName"))) );
		CopyLargeField( rsEmp2.Fields[_T("Photo")], rsEmp1.Fields[_T("Photo")] );

		rsEmp1.Update();

	//	Delete new record because this is a demonstration.
		rsEmp1.SetBookmark( rsEmp1.GetLastModified());
		rsEmp1.Delete();						// note 3

		rsEmp1.Close();
 		rsEmp2.Close();
		db.Close();

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