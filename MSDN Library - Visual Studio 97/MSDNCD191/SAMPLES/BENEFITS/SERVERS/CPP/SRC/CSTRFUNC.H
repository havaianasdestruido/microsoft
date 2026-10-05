//cstrfunc.h     - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for CStrFnc

#ifndef __CBSTR_H__
#define __CBSTR_H__

#include <windows.h>
#include <comutil.h>

class CStrFnc
{
protected:
	//Private member variables
	BOOL		bInit;
	BSTR		m_bstrString;
	long		m_oid;
	SAFEARRAY *	m_psa;

	//Private static member methods
	static BSTR vaFormat(BSTR bstrFormat, va_list argList);

	//Private member methods
	void Init();

public:
	//Public static member methods
	static	BOOL Boolean(BSTR bstrString);
	static	CURRENCY BstrToCy(BSTR bstrString);
	static	DATE FormatDate(BSTR Date);
	static	BSTR FormatDate(DATE Date, BSTR Format);
	static	BOOL IsEmpty(BSTR bstrValue);
	static	BOOL IsValidDate(DATE Date);
	static	BOOL IsValidDate(BSTR DateValue);
	static	BSTR LCase(BSTR bstrString);
	static	BSTR Left(BSTR bstrString, int x = -1);
	static	BSTR LoadResString(UINT iResID);
	static	BSTR LTrim(BSTR bstrString);
	static	BSTR Right(BSTR bstrString, int x = -1);
	static	BSTR RTrim(BSTR bstrString);
	static	BSTR Trim(BSTR bstrString);
	static	BSTR UCase(BSTR bstrString);
	static	long Val(BSTR bstrString);
	static	long VariantBstrToLong(VARIANT vValue);
	static	int	 BSTRCmpi(BSTR bstr1, BSTR bstr2);

	//Public member methods
	CStrFnc();
	CStrFnc(long lSize);
	CStrFnc(BSTR bstrString);
	~CStrFnc();

	BOOL	Boolean();
	CURRENCY BstrToCy();
	BSTR	Copy();
	void	Empty();
	BSTR	GetBstr();
	void	Format(BSTR Format, ...);
	BSTR	Left(int x = -1);
	BSTR	UCase();
	BSTR	LCase();
	int		Length();
	BSTR	Trim();
	BSTR	LTrim();
	BSTR	Right(int x = -1);
	BSTR	RTrim();
	long	Val();

	CStrFnc operator=(BSTR rhs);
};

typedef CStrFnc *PCStrFnc;

#endif //#ifndef __CBSTR_H__
