//  Sample SetOptionX

#include	<afxole.h>
#include	<dbdao.h>
#include	<dbdaoerr.h>
#include	<stdio.h>

void main(void)							// SetOptionX
{
	CdbDBEngine	dben;
	int			nExclusiveDelay, nSharedDelay;
	COleVariant	vVal;
	char		szPrompt1[] = 
					_T("Enter a new value for the ")
					_T("ExclusiveAsyncDelay registry key \n")
					_T("(in milliseconds): "),
				szPrompt2[] = 
					_T("Enter a new value for the ")
					_T("SharedAsyncDelay registry key \n")
					_T("(in milliseconds): ");
	//
	try
	{
	//	Get user input for new values of ExclusiveAsyncDelay and 
	//	SharedAsyncDelay registry keys.
		printf(szPrompt1);
		scanf(_T("%d"), &nExclusiveDelay);
		printf(_T("\n"));

		printf(szPrompt2);
		scanf(_T("%d"), &nSharedDelay);	
		printf(_T("\n"));

		if ((nExclusiveDelay > 0) && (nSharedDelay > 0))
		{
		//	Change the values of registry keys.
			vVal = (short)nExclusiveDelay;
			dben.SetOption(dbExclusiveAsyncDelay, &vVal);
			vVal = (short)nSharedDelay;
			dben.SetOption(dbSharedAsyncDelay,    &vVal);
			printf(_T("Registry keys changed to new values.\n"));
		}
		else
		{
			printf(_T("Registry keys left unchanged.\n"));
		}
		
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